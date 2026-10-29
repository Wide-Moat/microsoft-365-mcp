"""The drive kinds of the signed-link routes, end to end through the real Starlette app."""

import asyncio
import contextlib
import re
import time

import httpx
import pytest
from starlette.testclient import TestClient

from m365_mcp.links.routes import READ_CHUNK, REFUSED, links_app

FAR = int(time.time()) + 3600
BLOB = "https://contoso-my.sharepoint.com/blob/abc"


@pytest.fixture
def client(rt):
    return TestClient(links_app())


def issue(rt, *, op, kind, target, max_bytes=None, content_type=None, person="p1"):
    return rt.signer.issue(op=op, kind=kind, target=target, person=person, token="BEARER-XYZ", token_expires=FAR,
                           ttl=300 if op == "d" else 180, max_bytes=max_bytes, content_type=content_type)


# -- download ------------------------------------------------------------------------------------

def test_download_streams_the_bytes_through_a_redirect_without_the_bearer(rt, graph, client):
    graph.on("GET", "/me/drive/items/F1/content", None, status=302, headers={"Location": BLOB})
    graph.on("GET", re.escape(BLOB), b"hello drive", headers={"content-type": "text/plain", "content-length": "11"})
    tok = issue(rt, op="d", kind="drive", target={"i": "F1", "n": "a b.txt"})
    r = client.get(f"/d/{tok}")
    assert r.status_code == 200 and r.content == b"hello drive"
    assert "a%20b.txt" in r.headers["content-disposition"]
    first, second = [s for s in graph.seen if "mailboxSettings" not in s.path]
    assert first.headers["authorization"] == "Bearer BEARER-XYZ"
    assert "authorization" not in second.headers, "the pre-authenticated URL must not get the bearer token"


def test_download_of_another_drive_and_as_pdf_uses_the_right_path(rt, graph, client):
    graph.on("GET", "/drives/b!site/items/F2/content", None, status=302, headers={"Location": BLOB})
    graph.on("GET", re.escape(BLOB), b"%PDF")
    tok = issue(rt, op="d", kind="drive", target={"i": "F2", "d": "b!site", "f": "pdf", "n": "x.pdf"})
    assert client.get(f"/d/{tok}").content == b"%PDF"
    seen = graph.calls("GET", "/drives/b!site/items/F2/content")[0]
    assert seen.params == {"format": "pdf"}


def test_a_pdf_microsoft_cannot_make_says_so_rather_than_refused(rt, graph, client):
    # Measured: ?format=pdf of an empty .xlsx answers 302 to the conversion host, which answers 400
    # problem+json (Service_InvalidInput_InputBytesEmpty); QA read "refused the request ... none".
    graph.on("GET", "/me/drive/items/F1/content", None, status=302, headers={"Location": BLOB})
    graph.on("GET", re.escape(BLOB), {"title": "BadRequest", "status": 400}, status=400)
    tok = issue(rt, op="d", kind="drive", target={"i": "F1", "f": "pdf", "n": "empty.pdf"})
    r = client.get(f"/d/{tok}")
    assert r.status_code == 422
    assert "could not convert this file to PDF" in r.json()["error"] and "original" in r.json()["error"]


def test_range_requests_go_to_the_download_url_not_to_graph(rt, graph, client):
    graph.on("GET", "/me/drive/items/F1/content", None, status=302, headers={"Location": BLOB})
    graph.on("GET", re.escape(BLOB), b"ello", status=206, headers={"content-range": "bytes 1-4/11"})
    tok = issue(rt, op="d", kind="drive", target={"i": "F1"})
    r = client.get(f"/d/{tok}", headers={"Range": "bytes=1-4"})
    assert r.status_code == 206
    assert graph.seen[-1].headers["range"] == "bytes=1-4"


def test_download_of_a_missing_item_is_a_404_without_graph_detail(rt, graph, client):
    graph.on("GET", "/me/drive/items/F1/content", {"error": {"code": "itemNotFound", "message": "secret"}}, status=404)
    r = client.get(f"/d/{issue(rt, op='d', kind='drive', target={'i': 'F1'})}")
    assert r.status_code == 404 and "secret" not in r.text


# -- upload ---------------------------------------------------------------------------------------

def put_route(graph, path, item=None):
    item = item or {"id": "N1", "name": "report.pdf", "size": 100, "webUrl": "https://x.test/r"}
    graph.on("PUT", re.escape(path), item, status=201)


def test_a_small_upload_is_one_put_with_the_conflict_behaviour(rt, graph, client):
    put_route(graph, "/me/drive/items/P1:/report.pdf:/content")
    tok = issue(rt, op="u", kind="drive_upload", target={"p": "P1", "n": "report.pdf", "c": "fail"}, max_bytes=1000)
    r = client.put(f"/u/{tok}", content=b"x" * 100)
    assert r.status_code == 200 and r.json() == {"id": "N1", "name": "report.pdf", "size": 100, "web_url": "https://x.test/r"}
    puts = graph.calls("PUT", re.escape("/me/drive/items/P1:/report.pdf:/content"))
    assert len(puts) == 1 and puts[0].params == {"@microsoft.graph.conflictBehavior": "fail"}
    assert puts[0].content == b"x" * 100 and not graph.calls("POST", ".*createUploadSession")


@pytest.mark.parametrize("conflict", ["replace", "rename"])
def test_the_conflict_choice_reaches_graph(rt, graph, client, conflict):
    put_route(graph, "/drives/b!s/items/P1:/a.txt:/content")
    tok = issue(rt, op="u", kind="drive_upload", target={"p": "P1", "n": "a.txt", "c": conflict, "d": "b!s"}, max_bytes=100)
    assert client.put(f"/u/{tok}", content=b"abc").status_code == 200
    assert graph.calls("PUT", re.escape("/drives/b!s/items/P1:/a.txt:/content"))[0].params["@microsoft.graph.conflictBehavior"] == conflict


def test_a_file_name_with_spaces_and_unicode_is_encoded_in_the_path(rt, graph, client):
    put_route(graph, "/me/drive/items/P1:/Q3 Übersicht é.xlsx:/content")
    tok = issue(rt, op="u", kind="drive_upload", target={"p": "P1", "n": "Q3 Übersicht é.xlsx", "c": "fail"}, max_bytes=100)
    assert client.put(f"/u/{tok}", content=b"abc").status_code == 200


def test_a_10_mb_class_upload_uses_an_upload_session_in_aligned_chunks(rt, graph, client):
    size = 11 * 1024 * 1024
    graph.on("POST", re.escape("/me/drive/items/P1:/big.bin:/createUploadSession"), {"uploadUrl": "https://contoso-my.sharepoint.com/upload/sess"})
    graph.on("PUT", re.escape("https://contoso-my.sharepoint.com/upload/sess"), lambda s: httpx.Response(
        201 if s.headers["content-range"].endswith(f"-{size - 1}/{size}") else 202,
        json={"id": "BIG", "name": "big.bin", "size": size, "webUrl": "https://x.test/big"}))
    tok = issue(rt, op="u", kind="drive_upload", target={"p": "P1", "n": "big.bin", "c": "rename"}, max_bytes=20 * 1024 * 1024)
    r = client.put(f"/u/{tok}", content=bytes(size))
    assert r.status_code == 200 and r.json()["id"] == "BIG"
    sess = graph.calls("POST", re.escape("/me/drive/items/P1:/big.bin:/createUploadSession"))
    assert len(sess) == 1 and sess[0].json == {"item": {"@microsoft.graph.conflictBehavior": "rename", "name": "big.bin"}}
    assert not graph.calls("PUT", re.escape("/me/drive/items/P1:/big.bin:/content"))
    chunks = graph.calls("PUT", re.escape("https://contoso-my.sharepoint.com/upload/sess"))
    assert len(chunks) == 2
    pos = 0
    for i, ch in enumerate(chunks):
        n = len(ch.content)
        assert ch.headers["content-range"] == f"bytes {pos}-{pos + n - 1}/{size}"
        if i < len(chunks) - 1:
            assert n % 327680 == 0, "every chunk but the last is a multiple of 320 KiB"
        pos += n
    assert pos == size
    assert all("authorization" not in ch.headers for ch in chunks), "the upload URL is pre-authenticated"


def test_an_upload_over_max_bytes_is_413_and_nothing_reaches_graph(rt, graph, client):
    tok = issue(rt, op="u", kind="drive_upload", target={"p": "P1", "n": "a.bin", "c": "fail"}, max_bytes=1000)
    r = client.put(f"/u/{tok}", content=b"x" * 2000)
    assert r.status_code == 413
    assert [s for s in graph.seen if "mailboxSettings" not in s.path] == []


def test_a_graph_refusal_on_upload_is_reported_not_a_500(rt, graph, client):
    graph.on("PUT", re.escape("/me/drive/items/P1:/a.txt:/content"), {"error": {"code": "nameAlreadyExists"}}, status=409)
    tok = issue(rt, op="u", kind="drive_upload", target={"p": "P1", "n": "a.txt", "c": "fail"}, max_bytes=100)
    r = client.put(f"/u/{tok}", content=b"abc")
    assert r.status_code == 400 and "already exists" in r.json()["error"]


# -- refusals -------------------------------------------------------------------------------------

def test_a_tampered_link_is_403(rt, graph, client):
    tok = issue(rt, op="d", kind="drive", target={"i": "F1"})
    # The link is a short id; the signed value it stands for is still honoured on its own,
    # so tampering either one is refused.
    payload, sig = rt.signer._store.get(tok).split(".")
    flipped = payload[:-2] + ("AA" if payload[-2:] != "AA" else "BB")
    other_id = tok[:-2] + ("AA" if tok[-2:] != "AA" else "BB")
    for bad in (f"{flipped}.{sig}", f"{payload}.{sig[:-2]}xx", "garbage", payload, other_id):
        assert client.get(f"/d/{bad}").status_code == 403
    assert [s for s in graph.seen if "mailboxSettings" not in s.path] == []


def test_the_wrong_operation_is_403(rt, graph, client):
    down = issue(rt, op="d", kind="drive", target={"i": "F1"})
    up = issue(rt, op="u", kind="drive_upload", target={"p": "P1", "n": "a", "c": "fail"}, max_bytes=10)
    assert client.put(f"/u/{down}", content=b"x").status_code == 403
    assert client.get(f"/d/{up}").status_code == 403
    assert [s for s in graph.seen if "mailboxSettings" not in s.path] == []


def test_an_expired_link_is_403(rt, graph, client):
    old = rt.signer.issue(op="d", kind="drive", target={"i": "F1"}, person="p", token="t", token_expires=FAR, ttl=300,
                          now=time.time() - 1000)
    assert client.get(f"/d/{old}").status_code == 403


def test_a_compressed_answer_reaches_the_sandbox_as_the_real_file(rt, graph, client):
    # Measured: OneDrive gzips text files for a client that accepts gzip, and the link passed the
    # encoded bytes on, so curl without --compressed saved a gzip file. Identity is asked for, and an
    # encoded answer is decoded before it leaves.
    import gzip
    body = b"hello text " * 200
    graph.on("GET", "/me/drive/items/F1/content", None, status=302, headers={"Location": BLOB})
    graph.on("GET", re.escape(BLOB), gzip.compress(body), headers={"content-type": "text/plain", "content-encoding": "gzip"})
    tok = issue(rt, op="d", kind="drive", target={"i": "F1", "n": "t.txt"})
    with client.stream("GET", f"/d/{tok}") as r:
        raw = b"".join(r.iter_raw())
        assert r.status_code == 200 and "content-encoding" not in r.headers
    assert raw == body
    assert graph.calls("GET", re.escape(BLOB))[0].headers["accept-encoding"] == "identity"


def test_a_compressed_part_of_a_file_is_refused_rather_than_mislabelled(rt, graph, client):
    import gzip
    graph.on("GET", "/me/drive/items/F1/content", None, status=302, headers={"Location": BLOB})
    graph.on("GET", re.escape(BLOB), gzip.compress(b"ello"), status=206,
             headers={"content-encoding": "gzip", "content-range": "bytes 1-4/11"})
    tok = issue(rt, op="d", kind="drive", target={"i": "F1"})
    r = client.get(f"/d/{tok}", headers={"Range": "bytes=1-4"})
    assert r.status_code == 502 and "Range" in r.json()["error"]


def test_a_refused_link_says_to_ask_for_a_new_one_and_nothing_more(rt, graph, client):
    # An expired link answered 403 with an empty body; with short ids a restart
    # also refuses open links. One sentence for every refusal, naming no check.
    r = client.get("/d/unknownid0000000000000")
    assert r.status_code == 403 and "Ask for a new one" in r.text
    tok = issue(rt, op="d", kind="drive", target={"i": "F1"})
    assert client.put(f"/u/{tok}", content=b"x").text == r.text  # wrong operation: the same sentence


# -- an upload link moves one file, once --------------------------------------------------------

def test_an_upload_link_is_used_once(rt, graph, client):
    put_route(graph, "/me/drive/items/P1:/a.txt:/content")
    tok = issue(rt, op="u", kind="drive_upload", target={"p": "P1", "n": "a.txt", "c": "replace"}, max_bytes=100)
    assert client.put(f"/u/{tok}", content=b"first").status_code == 200
    again = client.put(f"/u/{tok}", content=b"second")
    assert again.status_code == 403 and again.text == REFUSED
    assert len(graph.calls("PUT", re.escape("/me/drive/items/P1:/a.txt:/content"))) == 1


def test_an_upload_the_link_refused_leaves_the_link_for_the_corrected_one(rt, graph, client):
    put_route(graph, "/me/drive/items/P1:/a.txt:/content")
    tok = issue(rt, op="u", kind="drive_upload", target={"p": "P1", "n": "a.txt", "c": "fail"}, max_bytes=10)
    assert client.put(f"/u/{tok}", content=b"x" * 11).status_code == 413
    assert client.put(f"/u/{tok}", content=b"x" * 10).status_code == 200


def test_a_second_upload_on_a_link_already_receiving_one_is_refused(rt, graph, client):
    put_route(graph, "/me/drive/items/P1:/a.txt:/content")
    tok = issue(rt, op="u", kind="drive_upload", target={"p": "P1", "n": "a.txt", "c": "fail"}, max_bytes=10)
    assert rt.signer.claim(tok)  # an upload in progress on this link
    r = client.put(f"/u/{tok}", content=b"abc")
    assert r.status_code == 409 and "already" in r.json()["error"]
    assert graph.writes() == []
    rt.signer.release(tok)
    assert client.put(f"/u/{tok}", content=b"abc").status_code == 200


def test_uploads_past_the_spool_budget_wait_and_are_not_saved(rt, graph, client, monkeypatch):
    from m365_mcp.links import routes
    put_route(graph, "/me/drive/items/P1:/a.txt:/content")
    monkeypatch.setattr(routes, "SPOOL", routes.SpoolBudget(100))
    tok = issue(rt, op="u", kind="drive_upload", target={"p": "P1", "n": "a.txt", "c": "fail"}, max_bytes=1000)
    r = client.put(f"/u/{tok}", content=b"x" * 101)
    assert r.status_code == 503 and int(r.headers["retry-after"]) > 0 and "Nothing was saved" in r.json()["error"]
    assert graph.writes() == [] and routes.SPOOL.used == 0
    assert client.put(f"/u/{tok}", content=b"x" * 100).status_code == 200
    assert routes.SPOOL.used == 0


def test_an_upload_the_disk_cannot_hold_gives_its_bytes_back_and_says_so(rt, graph, client, monkeypatch):
    import errno
    import tempfile

    from m365_mcp.links import routes
    put_route(graph, "/me/drive/items/P1:/a.txt:/content")
    monkeypatch.setattr(routes, "SPOOL", routes.SpoolBudget(1000))
    spools = []

    class Full(tempfile.SpooledTemporaryFile):
        def __init__(self, *a, **kw):
            super().__init__(*a, **kw)
            spools.append(self)

        def write(self, b):
            raise OSError(errno.ENOSPC, "No space left on device")
    monkeypatch.setattr(routes.tempfile, "SpooledTemporaryFile", Full)
    tok = issue(rt, op="u", kind="drive_upload", target={"p": "P1", "n": "a.txt", "c": "fail"}, max_bytes=1000)
    r = client.put(f"/u/{tok}", content=b"x" * 10)
    assert r.status_code == 507 and "Nothing was saved" in r.json()["error"]
    assert routes.SPOOL.used == 0 and spools and all(sp.closed for sp in spools)
    assert graph.writes() == []


async def test_an_upload_body_that_stalls_is_cut_off(monkeypatch):
    import asyncio

    from m365_mcp.links import routes
    monkeypatch.setattr(routes, "BODY_IDLE_TIMEOUT", 0.05)

    class Stalled:
        async def stream(self):
            yield b"abc"
            await asyncio.sleep(5)
            yield b"never"

    got = await routes._spool(Stalled(), 1000)
    assert got.status_code == 408 and routes.SPOOL.used == 0


# -- the upstream stream is released whatever the client does ----------------------------------

class _Counted(httpx.AsyncByteStream):
    def __init__(self, data: bytes) -> None:
        self.data, self.reads, self.closed = data, 0, False

    async def __aiter__(self):
        for i in range(0, len(self.data), 4):
            self.reads += 1
            yield self.data[i:i + 4]

    async def aclose(self) -> None:
        self.closed = True


def _upstream(graph, body: bytes) -> _Counted:
    stream = _Counted(body)
    graph.on("GET", "/me/drive/items/F1/content", lambda s: httpx.Response(
        200, stream=stream, headers={"content-type": "text/plain", "content-length": str(len(body))}))
    return stream


def test_head_answers_the_headers_without_pulling_the_file(rt, graph, client):
    stream = _upstream(graph, b"0123456789")
    r = client.head(f"/d/{issue(rt, op='d', kind='drive', target={'i': 'F1', 'n': 'a.txt'})}")
    assert r.status_code == 200 and r.headers["content-length"] == "10" and "a.txt" in r.headers["content-disposition"]
    assert stream.reads == 0 and stream.closed


def _scope(tok: str, spec: str) -> dict:
    # 2.3: Starlette watches for the disconnect in a task group and cancels the body when it comes, as
    # under uvicorn; 2.4: it learns of the client leaving only from a failed send.
    return {"type": "http", "asgi": {"version": "3.0", "spec_version": spec}, "http_version": "1.1",
            "method": "GET", "scheme": "https", "path": f"/d/{tok}", "raw_path": f"/d/{tok}".encode(),
            "root_path": "", "query_string": b"", "headers": [], "server": ("links", 443), "client": ("c", 1)}


@pytest.mark.parametrize("spec", ["2.3", "2.4"])
async def test_a_client_gone_before_the_first_chunk_still_releases_the_upstream(rt, graph, spec):
    stream = _upstream(graph, b"0123456789")
    tok = issue(rt, op="d", kind="drive", target={"i": "F1"})
    scope = _scope(tok, spec)

    async def receive():
        return {"type": "http.disconnect"}

    async def send(message):
        raise OSError("the client is gone")  # before the response could even start
    with contextlib.suppress(Exception):
        await links_app()(scope, receive, send)
    assert stream.reads == 0 and stream.closed


class _Hanging(httpx.AsyncByteStream):
    """One chunk, then a wait no client sits through; closing it yields to the loop once, as closing a
    real connection does, so a cancellation still pending interrupts a close that is not shielded."""

    def __init__(self) -> None:
        self.closed = False

    async def __aiter__(self):
        yield b"a" * READ_CHUNK  # a whole read: a shorter one waits in httpx for the rest
        await asyncio.sleep(30)
        yield b"never"

    async def aclose(self) -> None:
        await asyncio.sleep(0)
        self.closed = True


@pytest.mark.parametrize("spec", ["2.3", "2.4"])
async def test_a_client_gone_mid_stream_still_releases_the_upstream(rt, graph, spec):
    stream = _Hanging()
    graph.on("GET", "/me/drive/items/F1/content", lambda s: httpx.Response(
        200, stream=stream, headers={"content-type": "text/plain"}))
    tok = issue(rt, op="d", kind="drive", target={"i": "F1"})
    started = asyncio.Event()

    async def receive():
        await started.wait()
        return {"type": "http.disconnect"}

    async def send(message):
        if message["type"] == "http.response.body" and message.get("body"):
            started.set()
            if spec == "2.4":
                raise OSError("the client is gone")  # 2.4 finds out on the next send

    with contextlib.suppress(Exception):
        await asyncio.wait_for(links_app()(_scope(tok, spec), receive, send), timeout=5)
    assert started.is_set() and stream.closed


# -- throttling reaches the sandbox as throttling ---------------------------------------------

@pytest.mark.parametrize("status", [429, 503])
def test_a_throttled_download_is_passed_on_with_its_wait(rt, graph, client, status):
    graph.on("GET", "/me/drive/items/F1/content", {"error": {"code": "TooManyRequests"}}, status=status,
             headers={"Retry-After": "7"})
    r = client.get(f"/d/{issue(rt, op='d', kind='drive', target={'i': 'F1'})}")
    assert r.status_code == status and r.headers["retry-after"] == "7" and "7 seconds" in r.json()["error"]


@pytest.mark.parametrize("status", [429, 503])
def test_a_throttled_upload_is_passed_on_with_its_wait(rt, graph, client, status):
    graph.on("PUT", re.escape("/me/drive/items/P1:/a.txt:/content"), {"error": {"code": "TooManyRequests"}},
             status=status, headers={"Retry-After": "7"})
    tok = issue(rt, op="u", kind="drive_upload", target={"p": "P1", "n": "a.txt", "c": "fail"}, max_bytes=100)
    r = client.put(f"/u/{tok}", content=b"abc")
    assert r.status_code == status and r.headers["retry-after"] == "7" and "7 seconds" in r.json()["error"]
    assert client.put(f"/u/{tok}", content=b"abc").status_code == status, "a throttled upload leaves the link usable"
