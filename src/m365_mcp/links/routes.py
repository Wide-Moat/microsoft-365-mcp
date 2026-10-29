"""The two public routes: GET /d/<link> streams an object out, PUT /u/<link> streams one in.

Nothing else is served on the links listener. Every refusal is a 403 with one fixed sentence (or 413/415 for an
upload that breaks its own limits, which names the limit) — never a 500, never a detail of the object.
"""

from __future__ import annotations

import asyncio
import base64
from html import escape as html_escape
import mimetypes
import tempfile
import unicodedata
from collections.abc import AsyncIterator
from typing import Any
from urllib.parse import quote, unquote

import anyio
import httpx
from starlette.applications import Starlette
from starlette.requests import Request
from starlette.responses import JSONResponse, Response, StreamingResponse
from starlette.routing import Route

from m365_mcp.core.context import runtime
from m365_mcp.core.html import prepare_outgoing
from m365_mcp.graph.client import Graph
from m365_mcp.graph.errors import GraphError, to_person_error
from m365_mcp.links.tokens import Link, LinkRefused
from m365_mcp.observability import LINK_BYTES, LINK_TRANSFERS

SPOOL_MEMORY = 8 * 1024 * 1024
READ_CHUNK = 64 * 1024
ATTACH_SINGLE_LIMIT = 3 * 1024 * 1024
ATTACH_CHUNK = 10 * 320 * 1024          # under Graph's 4 MB per range for attachments
DRIVE_SINGLE_LIMIT = 4 * 1024 * 1024
DRIVE_CHUNK = 32 * 320 * 1024           # a multiple of 320 KiB, as drive upload sessions require
# Every upload is spooled before it goes to Graph. Uploads running at once share this much spool, so many
# parallel PUTs, each under its own link's limit, cannot fill the disk; past it a body waits for a retry.
SPOOL_TOTAL = 1024 * 1024 * 1024
SPOOL_RETRY_AFTER = 30
# A body that sends nothing for this long, or is still arriving after the longest a link lives, is cut off,
# so a stalled client does not hold its spool and its link forever.
BODY_IDLE_TIMEOUT = 30.0
BODY_TOTAL_TIMEOUT = 900.0


class SpoolBudget:
    """The bytes all running uploads hold in their spools, against one limit for the process."""

    def __init__(self, limit: int) -> None:
        self.limit = limit
        self.used = 0

    def take(self, n: int) -> bool:
        if self.used + n > self.limit:
            return False
        self.used += n
        return True

    def give(self, n: int) -> None:
        self.used = max(0, self.used - n)


SPOOL = SpoolBudget(SPOOL_TOTAL)


def _graph(link: Link) -> Graph:
    rt = runtime()
    kw: dict[str, Any] = {"transport": rt.transport}
    if rt.sleep is not None:
        kw["sleep"] = rt.sleep
    return Graph(link.token, link.person, **kw)


# One sentence for every refusal: it tells the agent what to do and nothing about which check failed.
REFUSED = "This link is not valid: it expired, was used up by a restart, or is wrong. Ask for a new one.\n"


def _refused() -> Response:
    return Response(REFUSED, status_code=403, media_type="text/plain")


def _verify(request: Request, op: str) -> Link:
    signer = runtime().signer
    if signer is None:
        raise LinkRefused()
    return signer.verify(request.path_params["value"], op=op)


THROTTLED_WAIT = 30  # seconds, when Microsoft throttles without saying how long


def _graph_failed(e: GraphError, status: int) -> Response:
    """Graph's failure as the sandbox's curl sees it. Throttling stays throttling, with its wait, so a
    retry loop can honour it; anything else is the status given, never a detail of the object."""
    if e.status in (429, 503):
        wait = e.retry_after or THROTTLED_WAIT
        return JSONResponse({"error": to_person_error(e).text}, status_code=e.status, headers={"Retry-After": str(wait)})
    return JSONResponse({"error": to_person_error(e).text}, status_code=status)


def _download_path(link: Link) -> str:
    t = link.target
    if link.kind == "attachment":
        return f"/me/messages/{t['m']}/attachments/{t['a']}/$value"
    if link.kind == "drive":
        base = f"/drives/{t['d']}/items/{t['i']}/content" if t.get("d") else f"/me/drive/items/{t['i']}/content"
        return base
    if link.kind == "recording":
        return f"/me/onlineMeetings/{t['m']}/recordings/{t['r']}/content"
    if link.kind == "transcript":
        return f"/me/onlineMeetings/{t['m']}/transcripts/{t['t']}/content"
    if link.kind == "note_resource":
        return f"/me/onenote/resources/{t['r']}/$value"
    raise LinkRefused()


async def download(request: Request) -> Response:
    try:
        link = _verify(request, "d")
    except LinkRefused:
        LINK_TRANSFERS.labels("d", "unknown", "refused").inc()
        return _refused()
    graph = _graph(link)
    # Identity: OneDrive gzips text for a client that accepts it, and the sandbox's curl would save the
    # gzip bytes (measured). Asked plainly, the bytes and their Content-Length are the file's own.
    headers: dict[str, str] = {"Accept-Encoding": "identity"}
    if request.headers.get("range"):
        headers["Range"] = request.headers["range"]
    if link.kind == "transcript":
        headers["Accept"] = link.target.get("a") or "text/vtt"
    params = {"format": "pdf"} if link.kind == "drive" and link.target.get("f") == "pdf" else None
    ctx = graph.stream(_download_path(link), headers=headers, params=params)
    try:
        upstream = await ctx.__aenter__()
    except GraphError as e:
        await graph.aclose()
        LINK_TRANSFERS.labels("d", link.kind, "graph_error").inc()
        if params and e.status in (400, 406, 415, 422):
            # Measured: an empty workbook's conversion answers 400 from the conversion host. The file is
            # there; the PDF is what cannot be made.
            return JSONResponse({"error": "Microsoft 365 could not convert this file to PDF; it may be empty or "
                                          "damaged. Download the original instead."}, status_code=422)
        return _graph_failed(e, 502 if e.status >= 500 else 404)
    except httpx.HTTPError:
        await graph.aclose()
        LINK_TRANSFERS.labels("d", link.kind, "upstream_error").inc()
        return JSONResponse({"error": "Microsoft 365 could not be reached. Try again."}, status_code=502)
    except BaseException:
        await graph.aclose()
        raise

    released = False

    async def release() -> None:
        # Once, from whichever path gets here first: the body's end, the response's end, or a refusal.
        # Shielded: under ASGI spec 2.3 Starlette cancels the body when the client leaves mid-stream, and
        # the cancellation, still pending, would end the close at its first await with the flag already
        # set, so no later caller would close the upstream either.
        nonlocal released
        if released:
            return
        released = True
        with anyio.CancelScope(shield=True):
            try:
                await ctx.__aexit__(None, None, None)
            finally:
                await graph.aclose()

    sent = 0

    async def body() -> AsyncIterator[bytes]:
        nonlocal sent
        try:
            # Raw bytes when the body is not encoded, so Range and Content-Length stay true; an encoded
            # body is decoded here, and its encoding and length are not passed on.
            chunks = upstream.aiter_bytes(READ_CHUNK) if encoded else upstream.aiter_raw(READ_CHUNK)
            async for chunk in chunks:
                sent += len(chunk)
                yield chunk
            LINK_TRANSFERS.labels("d", link.kind, "ok").inc()
        finally:
            await release()

    encoded = upstream.headers.get("content-encoding", "identity").lower() not in ("", "identity")
    if encoded and upstream.status_code == 206:
        # A range of an encoded body cannot be decoded on its own, and its offsets are not the file's.
        await release()
        LINK_TRANSFERS.labels("d", link.kind, "upstream_error").inc()
        return JSONResponse({"error": "Microsoft 365 sent this part compressed. Download the whole file, "
                                      "without a Range header."}, status_code=502)
    passed = ("content-type", "content-range", "accept-ranges") if encoded else (
        "content-type", "content-length", "content-range", "accept-ranges")
    out = {k: v for k, v in upstream.headers.items() if k.lower() in passed}
    name = link.target.get("n")
    if name:
        out["Content-Disposition"] = f"attachment; filename*=UTF-8''{quote(name)}"
    if request.method == "HEAD":
        # The headers are the answer; the file is not pulled from Microsoft to be thrown away.
        await release()
        return Response(status_code=upstream.status_code, headers=out)

    async def done() -> None:
        LINK_BYTES.labels("d").inc(sent)
        await release()
    return _Released(body(), done, status_code=upstream.status_code, headers=out)


class _Released(StreamingResponse):
    """A streamed download whose upstream is released when the response ends, however it ends: a client
    gone before the first chunk never starts the body, so the body's own finally would never run."""

    def __init__(self, content: AsyncIterator[bytes], done: Any, **kw: Any) -> None:
        super().__init__(content, **kw)
        self._done = done

    async def __call__(self, scope: Any, receive: Any, send: Any) -> None:
        try:
            await super().__call__(scope, receive, send)
        finally:
            await self._done()


def _size(n: int) -> str:
    return f"{n / (1024 * 1024):g} MB" if n >= 1024 * 1024 else f"{n / 1024:g} KB" if n >= 1024 else f"{n} bytes"


async def _spool(request: Any, limit: int | None) -> tuple[tempfile.SpooledTemporaryFile, int] | Response:
    """The body, spooled. On success the caller owns `size` bytes of SPOOL and gives them back; on any
    refusal here they are already given back and the spool is closed."""
    spool = tempfile.SpooledTemporaryFile(max_size=SPOOL_MEMORY)
    size = 0
    budget = SPOOL
    loop = asyncio.get_running_loop()
    deadline = loop.time() + BODY_TOTAL_TIMEOUT
    chunks = request.stream().__aiter__()

    def refuse(body: dict[str, str], status: int, headers: dict[str, str] | None = None) -> Response:
        spool.close()
        budget.give(size)
        return JSONResponse(body, status_code=status, headers=headers)

    while True:
        left = deadline - loop.time()
        try:
            chunk = await asyncio.wait_for(chunks.__anext__(), timeout=max(0.0, min(BODY_IDLE_TIMEOUT, left)))
        except StopAsyncIteration:
            break
        except TimeoutError:
            return refuse({"error": "The upload stalled or took too long, so it was cut off. Nothing was saved. "
                                    "Send it again."}, 408)
        except BaseException:
            spool.close()
            budget.give(size)
            raise
        if limit is not None and size + len(chunk) > limit:
            # The limit is the link's own, not anyone's data: naming it is what lets the file be fixed.
            return refuse({"error": f"The upload is larger than this link accepts: at most {limit} bytes "
                                    f"({_size(limit)}). Nothing was saved."}, 413)
        if not budget.take(len(chunk)):
            return refuse({"error": "The server is receiving too many uploads at once. Nothing was saved. Send it "
                                    f"again in {SPOOL_RETRY_AFTER} seconds with the same link."}, 503,
                          {"Retry-After": str(SPOOL_RETRY_AFTER)})
        try:
            spool.write(chunk)
        except OSError:
            # The disk under the spool is full (or failing): the chunk's bytes go back with the rest.
            budget.give(len(chunk))
            return refuse({"error": "The server has no room to receive this upload now. Nothing was saved. Send "
                                    f"it again in {SPOOL_RETRY_AFTER} seconds with the same link."}, 507,
                          {"Retry-After": str(SPOOL_RETRY_AFTER)})
        size += len(chunk)
    spool.seek(0)
    return spool, size


async def _chunks(spool: Any, size: int, step: int) -> AsyncIterator[bytes]:
    sent = 0
    while sent < size:
        data = spool.read(step)
        if not data:
            break
        sent += len(data)
        yield data


async def upload(request: Request) -> Response:
    try:
        link = _verify(request, "u")
    except LinkRefused:
        LINK_TRANSFERS.labels("u", "unknown", "refused").inc()
        return _refused()
    signer = runtime().signer
    value = request.path_params["value"]
    if signer is None or not signer.claim(value):
        LINK_TRANSFERS.labels("u", link.kind, "busy").inc()
        return JSONResponse({"error": "This link is already receiving an upload; one link takes one file, once. "
                                      "Nothing was saved."}, status_code=409)
    # Spent only once a file was delivered; a body the link refused leaves it for the corrected one.
    delivered = False
    try:
        response = await _receive(request, link)
        delivered = response.status_code == 200
        return response
    finally:
        if delivered:
            signer.consume(value, link.expires)
        else:
            signer.release(value)


async def _receive(request: Request, link: Link) -> Response:
    ctype = request.headers.get("content-type", "").split(";")[0].strip().lower()
    if link.content_type and ctype != link.content_type:
        return JSONResponse({"error": f"This link takes {link.content_type} only: send it with the header "
                                      f"Content-Type: {link.content_type}. Nothing was saved."}, status_code=415)
    got = await _spool(request, link.max_bytes)
    if isinstance(got, Response):
        LINK_TRANSFERS.labels("u", link.kind, "too_large" if got.status_code == 413 else "refused").inc()
        return got
    spool, size = got
    graph = _graph(link)
    try:
        result = await _UPLOADERS[link.kind](graph, link, spool, size, request)
        LINK_TRANSFERS.labels("u", link.kind, "ok").inc()
        LINK_BYTES.labels("u").inc(size)
        return JSONResponse(result)
    except GraphError as e:
        LINK_TRANSFERS.labels("u", link.kind, "graph_error").inc()
        return _graph_failed(e, 502 if e.status >= 500 else 400)
    except _Rejected as r:
        return JSONResponse({"error": r.reason}, status_code=422)
    except (KeyError, ValueError, httpx.HTTPError):
        # A malformed answer from Microsoft or a broken connection: nothing for the sandbox to fix.
        LINK_TRANSFERS.labels("u", link.kind, "upstream_error").inc()
        return JSONResponse({"error": "Microsoft 365 did not accept the upload. Try again."}, status_code=502)
    finally:
        spool.close()
        SPOOL.give(size)
        await graph.aclose()


class _Rejected(Exception):
    def __init__(self, reason: str) -> None:
        super().__init__("rejected")
        self.reason = reason


_BIDI_CONTROLS = frozenset("\u202a\u202b\u202c\u202d\u202e\u2066\u2067\u2068\u2069\u200e\u200f")


def _header_text(raw: str) -> str:
    """A header value as the person wrote it. Starlette reads header bytes as Latin-1, while curl sends
    the terminal's UTF-8: "Reunión" arrived as "ReuniÃ³n". Bytes that are valid
    UTF-8 are read as UTF-8; anything else stays Latin-1. RFC 5987's UTF-8''<percent-encoded> is the
    form for a client that cannot send raw bytes."""
    if raw[:7].lower() == "utf-8''":
        return unquote(raw[7:], encoding="utf-8", errors="strict")
    try:
        return raw.encode("latin-1").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return raw


def _file_name(request: Request) -> str:
    try:
        raw = _header_text(request.headers.get("x-file-name", "").strip())
    except UnicodeDecodeError:
        raise _Rejected("The X-File-Name header is not valid UTF-8 percent-encoding.") from None
    # Control characters and bidirectional overrides go (a name that reads differently than it is);
    # joiners and spaces of any script stay, so a family emoji or a Persian name is kept whole.
    raw = "".join(ch for ch in raw if unicodedata.category(ch) != "Cc" and ch not in _BIDI_CONTROLS)
    name = raw.replace("/", "_").replace("\\", "_").strip(". ")
    if not name:
        raise _Rejected("Send the file name in an X-File-Name header.")
    return name[:200]


async def _draft_body(graph: Graph, link: Link, spool: Any, size: int, request: Request) -> dict[str, Any]:
    from m365_mcp.tools.mail import _attach_inline, _mark_value, _part_id, insert_new_part

    prepared = prepare_outgoing(spool.read())
    if prepared.errors:
        raise _Rejected(" ".join(prepared.errors))
    draft_id = link.target["d"]
    d = await _unsent_draft(graph, draft_id, params={"$select": "id,isDraft,body"},
                            headers={"Prefer": 'outlook.body-content-type="html"'})
    html = insert_new_part((d.get("body") or {}).get("content") or "", prepared.html, _part_id(_mark_value(d)))
    await graph.patch(f"/me/messages/{draft_id}", {"body": {"contentType": "html", "content": html}})
    await _attach_inline(graph, draft_id, prepared.inline_images)
    return {"draft_id": draft_id, "html": prepared.report()}


async def _unsent_draft(graph: Graph, draft_id: str, **kw: Any) -> dict[str, Any]:
    """The draft an upload writes into, refused when it was sent or is being sent: a file or a body added
    after the person said send would go out unseen, or land on the copy in Sent Items."""
    from m365_mcp.tools.mail import DRAFT_MARK, SENT_MARK, _mark_value

    params = kw.pop("params", {"$select": "id,isDraft"})
    params["$expand"] = f"singleValueExtendedProperties($filter=id eq '{DRAFT_MARK}')"
    d = await graph.get(f"/me/messages/{draft_id}", params=params, **kw)
    if not d.get("isDraft"):
        raise _Rejected("That message is not a draft any more. Nothing was changed.")
    if _mark_value(d) == SENT_MARK:
        raise _Rejected("This draft was already sent, so it is not a draft any more. Nothing was changed.")
    return d


async def _draft_attachment(graph: Graph, link: Link, spool: Any, size: int, request: Request) -> dict[str, Any]:
    name = _file_name(request)
    draft_id = link.target["d"]
    await _unsent_draft(graph, draft_id)
    ctype = (request.headers.get("content-type") or "").split(";")[0].strip().lower()
    if ctype in ("", "application/octet-stream"):
        # A type from the name when the sender gave none: every attachment arrived as octet-stream,
        # which recipients' clients open worse.
        ctype = mimetypes.guess_type(name)[0] or "application/octet-stream"
    if size < ATTACH_SINGLE_LIMIT:
        await graph.post(f"/me/messages/{draft_id}/attachments", {
            "@odata.type": "#microsoft.graph.fileAttachment", "name": name, "contentType": ctype,
            "contentBytes": base64.b64encode(spool.read()).decode()})
    else:
        session = await graph.post(f"/me/messages/{draft_id}/attachments/createUploadSession", {
            "AttachmentItem": {"attachmentType": "file", "name": name, "size": size, "contentType": ctype}}) or {}
        await graph.upload_session(session["uploadUrl"], _chunks(spool, size, ATTACH_CHUNK), size)
    return {"draft_id": draft_id, "attached": name, "size": size}


async def _drive_upload(graph: Graph, link: Link, spool: Any, size: int, request: Request) -> dict[str, Any]:
    t = link.target
    name = t.get("n") or _file_name(request)
    conflict = t.get("c", "fail")
    parent = f"/drives/{t['d']}/items/{t['p']}" if t.get("d") else f"/me/drive/items/{t['p']}"
    target = f"{parent}:/{quote(name)}:"
    if size <= DRIVE_SINGLE_LIMIT:
        resp = await graph.request("PUT", f"{target}/content", params={"@microsoft.graph.conflictBehavior": conflict},
                                   content=spool.read(), headers={"Content-Type": "application/octet-stream"})
        item = resp.json()
    else:
        session = await graph.post(f"{target}/createUploadSession", {
            "item": {"@microsoft.graph.conflictBehavior": conflict, "name": name}}) or {}
        item = await graph.upload_session(session["uploadUrl"], _chunks(spool, size, DRIVE_CHUNK), size)
    return {"id": item.get("id"), "name": item.get("name"), "size": item.get("size"), "web_url": item.get("webUrl")}


async def _note_body(graph: Graph, link: Link, spool: Any, size: int, request: Request) -> dict[str, Any]:
    prepared = prepare_outgoing(spool.read())
    if prepared.errors:
        raise _Rejected(" ".join(prepared.errors))
    t = link.target
    if t.get("page"):
        await graph.patch_note_page(t["page"], [{"target": "body", "action": t.get("a", "append"),
                                                 "content": prepared.html}])
        return {"page_id": t["page"], "html": prepared.report()}
    title = html_escape(t.get("title") or "Untitled")
    doc = f"<!DOCTYPE html><html><head><title>{title}</title></head><body>{prepared.html}</body></html>"
    resp = await graph.request("POST", f"/me/onenote/sections/{t['s']}/pages", content=doc.encode(),
                               headers={"Content-Type": "text/html"}, retry_server_errors=False)
    page = resp.json()
    return {"page_id": page.get("id"), "html": prepared.report()}


_UPLOADERS = {
    "draft_body": _draft_body,
    "draft_attachment": _draft_attachment,
    "drive_upload": _drive_upload,
    "note_body": _note_body,
}


def links_app() -> Starlette:
    return Starlette(routes=[
        Route("/d/{value:path}", download, methods=["GET", "HEAD"]),
        Route("/u/{value:path}", upload, methods=["PUT"]),
    ])
