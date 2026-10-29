"""The only HTTP client to Microsoft Graph.

Everything that makes Graph quietly wrong is handled here, once: ids that change when a
message moves, times in the wrong zone, throttling, paging, and error bodies that carry the
person's data.
"""

from __future__ import annotations

import asyncio
import html
import logging
import random
import re
import time
import uuid
from collections import defaultdict
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from dataclasses import dataclass
from datetime import datetime
from typing import Any
from urllib.parse import urlsplit

import httpx

from m365_mcp import __version__
from m365_mcp.graph.errors import GraphError, scope_from_message  # noqa: I001

# httpx logs every request URL at INFO, and Graph URLs carry search strings, addresses and
# subjects in their query. Nothing below WARNING from the HTTP stack reaches the logs.
for _name in ("httpx", "httpcore"):
    logging.getLogger(_name).setLevel(logging.WARNING)

GRAPH = "https://graph.microsoft.com/v1.0"
GRAPH_HOST = "graph.microsoft.com"
# Where Graph sends a download (302) or an upload session: Microsoft's file, media and mail hosts.
# A URL elsewhere is not followed, though it came from Graph's own answer: the bytes, and for an upload
# the person's file, would go to a host nobody chose.
FILE_HOSTS = ("sharepoint.com", "onedrive.com", "1drv.com", "svc.ms", "microsoft.com", "office.com",
              "office365.com", "office.net", "outlook.com")


def on_file_host(url: str, hosts: tuple[str, ...] = FILE_HOSTS) -> bool:
    """https on one of `hosts` or a subdomain of one; `evilsharepoint.com` is neither."""
    parts = urlsplit(url)
    host = (parts.hostname or "").lower()
    return parts.scheme == "https" and any(host == h or host.endswith("." + h) for h in hosts)

# Paths a DELETE may target. Messages are never deleted — they are moved to Deleted Items.
_DELETE_ALLOWED = [
    re.compile(p) for p in (
        r"^/me/drive/items/[^/]+$",
        r"^/drives/[^/]+/items/[^/]+$",
        r"^/me/todo/lists/[^/]+/tasks/[^/]+$",
        r"^/planner/tasks/[^/]+$",
        r"^/me/onenote/pages/[^/]+$",
        r"^/me/contacts/[^/]+$",
        r"^/me/contactFolders/[^/]+/contacts/[^/]+$",
        r"^/sites/[^/]+/lists/[^/]+/items/[^/]+$",
        r"^/me/messages/[^/]+/attachments/[^/]+$",   # removing an attachment from a draft
        r"^/chats/[^/]+/members/[^/]+$",              # removing someone from a group chat (save_chat)
        r"^/drives/[^/]+/items/[^/]+/permissions/[^/]+$",
        r"^/me/drive/items/[^/]+/permissions/[^/]+$",
        r"^/me/events/[^/]+$",   # an event nobody is invited to (cancel_event); it goes to Deleted Items
    )
]

_ENCODED_SEPARATOR = re.compile(r"%(2e|2f|5c|23|3f)", re.I)


class UnsafePath(GraphError):
    def __init__(self) -> None:
        super().__init__(400, "invalidRequestPath", "")


def _check_path(path: str) -> None:
    """Ids and names from the model end up inside these paths. A "#", a "?", a "." or ".." segment,
    or an encoded separator would let one id rewrite the request: "D/send#" turns a reply into a
    send, ".." turns a task into its whole list. Queries go through `params`, never the path."""
    if any(ch in path for ch in "#?\\") or any(ord(ch) < 0x21 for ch in path):
        raise UnsafePath()
    # An empty id leaves an empty segment ("/me/events/"), which Graph answers with the collection.
    if "//" in path or path.endswith("/"):
        raise UnsafePath()
    if _ENCODED_SEPARATOR.search(path):
        raise UnsafePath()
    for seg in path.split("/"):
        if seg in (".", "..") or seg.strip(":") in (".", ".."):
            raise UnsafePath()


NOTE_PAGE_ATTEMPTS = 6
NOTE_TITLE_AT_RISK_SECONDS = 60
_NOTE_TITLE = re.compile(r"<title[^>]*>(.*?)</title>", re.I | re.S)
MAX_ATTEMPTS = 3
MAX_WAIT = 30.0


def _retry_after(resp: httpx.Response) -> int | None:
    raw = resp.headers.get("Retry-After")
    if raw and raw.isdigit():
        return int(raw)
    return None


class Pacer:
    """Per-person limits that keep us under Graph's throttling instead of collecting 429s."""

    def __init__(self) -> None:
        self._sems: dict[tuple[str, str], asyncio.Semaphore] = {}
        self._next_slot: dict[tuple[str, str], float] = defaultdict(float)
        self._locks: dict[tuple[str, str], asyncio.Lock] = defaultdict(asyncio.Lock)

    CONCURRENCY = {"outlook": 4, "drive": 4, "teams": 1, "onenote": 1, "other": 4, "bookings": 4}
    MIN_INTERVAL = {"teams": 1.0, "onenote": 0.5}

    def sem(self, person: str, family: str) -> asyncio.Semaphore:
        key = (person, family)
        if key not in self._sems:
            self._sems[key] = asyncio.Semaphore(self.CONCURRENCY.get(family, 4))
        return self._sems[key]

    async def space(self, person: str, family: str, sleep=asyncio.sleep, clock=time.monotonic) -> None:
        interval = self.MIN_INTERVAL.get(family)
        if not interval:
            return
        key = (person, family)
        async with self._locks[key]:
            now = clock()
            wait = self._next_slot[key] - now
            if wait > 0:
                await sleep(wait)
                now += wait
            self._next_slot[key] = now + interval


PACER = Pacer()


def family_of(path: str) -> str:
    if path.startswith(("/me/messages", "/me/mailFolders", "/me/events", "/me/calendar", "/me/sendMail",
                        "/me/mailboxSettings", "/me/contacts", "/me/contactFolders", "/me/findMeetingTimes",
                        "/me/calendarView", "/me/calendars")):
        return "outlook"
    if path.startswith(("/me/drive", "/drives/", "/sites/", "/shares/")):
        return "drive"
    if path.startswith(("/me/chats", "/chats/", "/teams/", "/me/joinedTeams")):
        return "teams"
    if path.startswith("/me/onenote"):
        return "onenote"
    if path.startswith("/solutions/bookingBusinesses"):
        return "bookings"
    return "other"


@dataclass
class Page:
    items: list[dict[str, Any]]
    next_link: str | None


class Graph:
    """One instance per tool call, carrying the person's token."""

    def __init__(self, token: str, person_key: str, *, timezone: str | None = None,
                 transport: httpx.AsyncBaseTransport | None = None, sleep=asyncio.sleep,
                 pacer: Pacer = PACER) -> None:
        self._token = token
        self._person = person_key
        self.timezone = timezone
        self._sleep = sleep
        self._pacer = pacer
        self._client = httpx.AsyncClient(
            transport=transport,
            timeout=httpx.Timeout(30.0, connect=10.0),
            follow_redirects=False,
        )

    async def aclose(self) -> None:
        await self._client.aclose()

    async def __aenter__(self) -> Graph:
        return self

    async def __aexit__(self, *exc: object) -> None:
        await self.aclose()

    # -- request core -------------------------------------------------------------------

    def _headers(self, extra: dict[str, str] | None) -> dict[str, str]:
        # Times stay in UTC on the wire; the server converts to the person's zone itself. Asking
        # Graph for the mailbox zone would return local wall times that the server could only read
        # back correctly for zones it can name.
        prefer = ['IdType="ImmutableId"']
        h = {
            "Authorization": f"Bearer {self._token}",
            "Prefer": ", ".join(prefer),
            "client-request-id": str(uuid.uuid4()),
            "User-Agent": f"m365-mcp/{__version__}",
            "Accept": "application/json",
        }
        if extra:
            for k, v in extra.items():
                if k.lower() == "prefer":
                    h["Prefer"] = h["Prefer"] + ", " + v
                else:
                    h[k] = v
        return h

    @staticmethod
    def _url(path_or_url: str) -> str:
        if path_or_url.startswith("https://"):
            parts = urlsplit(path_or_url)
            # A next link or a monitor URL from Graph; never anything else, so a cursor
            # cannot point the server at an arbitrary host.
            if parts.hostname != GRAPH_HOST or parts.fragment:
                raise ValueError("refusing a non-Graph URL")
            return path_or_url
        if not path_or_url.startswith("/"):
            raise ValueError("Graph paths start with /")
        _check_path(path_or_url)
        return GRAPH + path_or_url

    async def request(self, method: str, path: str, *, params: dict[str, Any] | None = None,
                      json: Any = None, content: bytes | None = None, headers: dict[str, str] | None = None,
                      retry_server_errors: bool | None = None) -> httpx.Response:
        url = self._url(path)
        if method == "DELETE" and not any(r.match(path) for r in _DELETE_ALLOWED):
            raise ValueError("DELETE is not allowed on this path")
        family = family_of(urlsplit(url).path.removeprefix("/v1.0"))
        idempotent = method in ("GET", "HEAD", "PUT", "DELETE") if retry_server_errors is None else retry_server_errors
        waited = 0.0
        for attempt in range(1, MAX_ATTEMPTS + 1):
            async with self._pacer.sem(self._person, family):
                await self._pacer.space(self._person, family, sleep=self._sleep)
                resp = await self._client.request(method, url, params=params, json=json, content=content,
                                                  headers=self._headers(headers))
            if resp.status_code < 400:
                return resp
            # 429 means Graph did not take the request, so any method may repeat it. A 5xx may have
            # been half-done: only requests that are safe to repeat are retried.
            retry = resp.status_code == 429 or (idempotent and resp.status_code in (502, 503, 504))
            if retry and attempt < MAX_ATTEMPTS:
                ra = _retry_after(resp)
                delay = float(ra) if ra is not None else min(2 ** (attempt - 1) + random.random(), 8.0)
                if waited + delay > MAX_WAIT:
                    raise self._error(resp, method)
                waited += delay
                await self._sleep(delay)
                continue
            raise self._error(resp, method)
        raise AssertionError("unreachable")

    @staticmethod
    def _error(resp: httpx.Response, method: str = "") -> GraphError:
        code, message, inner = "", "", ""
        try:
            body = resp.json()
            err = body.get("error", {}) if isinstance(body, dict) else {}
            code = str(err.get("code", ""))
            message = str(err.get("message", ""))
            inner = str((err.get("innerError") or err.get("innererror") or {}).get("code", "") or "")
        except (ValueError, AttributeError):
            pass
        if message.strip() == "ParsingFailed":
            # Teams says an id did not parse in the message, under the generic code BadRequest; the
            # message is a fixed word, never the person's data, so it is kept as the inner code.
            inner = "ParsingFailed"
        missing = scope_from_message(message) if resp.status_code == 403 else None
        return GraphError(resp.status_code, code, resp.headers.get("request-id", ""), _retry_after(resp), missing,
                          inner_code=inner, method=method)

    async def get(self, path: str, *, params: dict[str, Any] | None = None, headers: dict[str, str] | None = None) -> dict[str, Any]:
        resp = await self.request("GET", path, params=params, headers=headers)
        # 204 with no body is a real answer: a new mailbox reads its unset time zone that way.
        return resp.json() if resp.content else {}

    async def post(self, path: str, json: Any = None, *, headers: dict[str, str] | None = None) -> dict[str, Any] | None:
        resp = await self.request("POST", path, json=json, headers=headers)
        return resp.json() if resp.content and "json" in resp.headers.get("content-type", "") else None

    async def patch(self, path: str, json: Any, *, if_match: str | None = None, headers: dict[str, str] | None = None) -> dict[str, Any] | None:
        h = dict(headers or {})
        if if_match:
            h["If-Match"] = if_match
        resp = await self.request("PATCH", path, json=json, headers=h)
        return resp.json() if resp.content and "json" in resp.headers.get("content-type", "") else None

    async def patch_note_page(self, page: str, ops: list[dict[str, Any]]) -> None:
        """Change a OneNote page. A page OneNote has just created is not always ready to change: on the
        test tenant one PATCH in three right after creation answered 404 for a page that existed. A 404
        is tried again every 2 seconds, six attempts in all, before it counts as missing.

        A body change to a page made in the last minute, or to one whose listed title is empty, also
        sends the page's own title. Measured: a body PATCH within seconds of creation left the listed
        title empty for good in two tries of three (the content kept it), the same PATCH carrying a
        title replace kept it three times of three, and a later title replace mended a page already
        hit. Older pages are left alone, so a rename made meanwhile elsewhere is not undone."""
        if not any(op.get("target") == "title" for op in ops):
            title = await self._note_title_to_keep(page)
            if title:
                ops = [{"target": "title", "action": "replace", "content": title}, *ops]
        for attempt in range(1, NOTE_PAGE_ATTEMPTS + 1):
            try:
                await self.patch(f"/me/onenote/pages/{page}/content", ops)
                return
            except GraphError as e:
                if e.status != 404 or attempt == NOTE_PAGE_ATTEMPTS:
                    raise
                await self._sleep(2)

    async def _note_title_to_keep(self, page: str) -> str | None:
        """The title, as HTML, that a body change to this page must carry; None when it needs none."""
        try:
            meta = await self.get(f"/me/onenote/pages/{page}", params={"$select": "title,createdDateTime"})
        except GraphError as e:
            if e.status != 404:
                raise
            meta = {}  # a page just made can answer 404 for a moment: it is as fresh as they come
        try:
            made = datetime.fromisoformat(str(meta.get("createdDateTime") or "").replace("Z", "+00:00"))
            fresh = time.time() - made.timestamp() < NOTE_TITLE_AT_RISK_SECONDS
        except ValueError:
            fresh = True
        if meta.get("title") and not fresh:
            return None
        found = _NOTE_TITLE.search(await self.note_page_html(page))
        # The title is text: read back as text and escaped again, so nothing in it can be markup.
        text = html.unescape(found.group(1)).strip() if found else ""
        return html.escape(text) if text else None

    async def note_page_html(self, page: str) -> str:
        """The HTML of a OneNote page before it is changed, with the same patience as patch_note_page:
        the content of a page just made answers 404 for a few seconds too."""
        for attempt in range(1, NOTE_PAGE_ATTEMPTS + 1):
            try:
                resp = await self.request("GET", f"/me/onenote/pages/{page}/content", headers={"Accept": "text/html"})
                return resp.text
            except GraphError as e:
                if e.status != 404 or attempt == NOTE_PAGE_ATTEMPTS:
                    raise
                await self._sleep(2)
        raise AssertionError("unreachable")

    async def delete(self, path: str, *, if_match: str | None = None) -> None:
        await self.request("DELETE", path, headers={"If-Match": if_match} if if_match else None)

    async def page(self, path_or_next: str, *, params: dict[str, Any] | None = None,
                   headers: dict[str, str] | None = None) -> Page:
        if path_or_next.startswith("https://"):
            params = None  # the next link carries its own query
        body = await self.get(path_or_next, params=params, headers=headers)
        return Page(items=list(body.get("value", [])), next_link=body.get("@odata.nextLink"))

    async def immutable_ids(self, ids: list[str]) -> dict[str, str]:
        """Outlook ids as immutable ids. $search ignores Prefer: IdType="ImmutableId" and returns REST ids,
        which change when a message moves (rule 34). Graph refuses a whole batch when one id in it is
        already immutable, so a refused batch is retried one id at a time, and an id is kept as it is only
        once Graph confirms it is immutable. Nothing is created: Graph only answers this question with POST."""
        if not ids:
            return {}
        out: dict[str, str] = {}

        async def translate(batch: list[str], source: str, target: str) -> dict[str, str]:
            """Every id of the batch translated, or a failure: a result missing an id (Graph answers
            200 with per-item errorDetails) must not leave that id to pass as it was."""
            body = await self.post("/me/translateExchangeIds", {
                "inputIds": batch, "sourceIdType": source, "targetIdType": target}) or {}
            got = {r["sourceId"]: r["targetId"] for r in body.get("value", []) if r.get("sourceId") and r.get("targetId")}
            if any(i not in got for i in batch):
                raise GraphError(502, "translationIncomplete", "")
            return got

        async def ask(batch: list[str]) -> None:
            out.update(await translate(batch, "restId", "restImmutableEntryId"))
        try:
            await ask(ids)
        except GraphError as e:
            if e.status != 400 or e.code != "InvalidArgument":
                raise
            for i in ids:
                try:
                    await ask([i])
                except GraphError as one:
                    if one.status != 400 or one.code != "InvalidArgument":
                        raise
                    # Refused alone as a REST id. It is kept only when Graph also takes it as an
                    # immutable id (measured: an immutable id translates back, a REST id is refused);
                    # anything else is the tool's error, never a move-sensitive id passed off.
                    await translate([i], "restImmutableEntryId", "restId")
                    out[i] = i
        return out

    # -- bytes --------------------------------------------------------------------------

    @asynccontextmanager
    async def stream(self, path: str, *, headers: dict[str, str] | None = None,
                     params: dict[str, str] | None = None) -> AsyncIterator[httpx.Response]:
        """GET that follows Graph's 302 to a pre-authenticated download URL, streamed."""
        url = self._url(path)
        resp = await self._client.send(
            self._client.build_request("GET", url, params=params, headers=self._headers(headers)), stream=True)
        try:
            if resp.status_code in (301, 302, 303, 307, 308):
                location = resp.headers.get("Location", "")
                await resp.aclose()
                if not on_file_host(location):
                    raise GraphError(502, "badRedirect", "")
                # The download URL is pre-authenticated; the bearer must not follow it.
                fwd = {k: v for k, v in (headers or {}).items() if k.lower() in ("range", "accept-encoding")}
                resp = await self._client.send(self._client.build_request("GET", location, headers=fwd), stream=True)
            if resp.status_code >= 400:
                await resp.aread()
                raise self._error(resp)
            yield resp
        finally:
            await resp.aclose()

    # Hosts Graph hands out pre-authenticated monitor URLs on (copy progress).
    MONITOR_HOSTS = (".sharepoint.com", ".onedrive.com", ".svc.ms")

    async def monitor(self, url: str) -> dict[str, Any]:
        """Read a copy monitor. These URLs are pre-authenticated and live on SharePoint or OneDrive
        hosts, so the request carries no bearer and goes only to those hosts."""
        host = urlsplit(url).hostname or ""
        if not url.startswith("https://") or not (host == GRAPH_HOST or host.endswith(self.MONITOR_HOSTS)):
            raise ValueError("refusing a monitor URL outside Microsoft's file hosts")
        headers = self._headers(None) if host == GRAPH_HOST else {"Accept": "application/json"}
        resp = await self._client.get(url, headers=headers)
        if resp.status_code in (302, 303):
            return {"status": "completed"}
        if resp.status_code >= 400:
            raise self._error(resp)
        return resp.json() if resp.content else {}

    async def upload_session(self, upload_url: str, chunks: AsyncIterator[bytes], size: int) -> dict[str, Any]:
        """PUT ranges to a pre-authenticated upload URL. Chunk sizes are the caller's job."""
        if not on_file_host(upload_url):
            raise ValueError("refusing an upload URL outside Microsoft's file hosts")
        offset = 0
        last: httpx.Response | None = None
        async for chunk in chunks:
            end = offset + len(chunk) - 1
            last = await self._client.put(upload_url, content=chunk, headers={
                "Content-Length": str(len(chunk)), "Content-Range": f"bytes {offset}-{end}/{size}"})
            if last.status_code >= 400:
                raise self._error(last)
            offset = end + 1
        if last is None or offset != size:
            raise GraphError(400, "incompleteUpload", "")
        return last.json() if last.content else {}
