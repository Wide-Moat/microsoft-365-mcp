import base64
import time

import httpx
import pytest
from starlette.testclient import TestClient

from m365_mcp.links.routes import ATTACH_CHUNK, links_app
from m365_mcp.tools.mail import NEW_PART_ID


@pytest.fixture
def client(rt):
    return TestClient(links_app())


@pytest.fixture(autouse=True)
def an_unsent_draft(graph):
    """d1 is a draft made here and not sent; a test about another state registers its own answer."""
    from m365_mcp.tools.mail import DRAFT_MARK
    graph.on("GET", "/me/messages/d1", {"id": "d1", "isDraft": True, "body": {"content": ""},
                                        "singleValueExtendedProperties": [{"id": DRAFT_MARK, "value": "1"}]})


def link(rt, op, kind, target, **kw):
    from helpers import make_token
    return rt.signer.issue(op=op, kind=kind, target=target, person="p1", token=make_token(),
                           token_expires=int(time.time()) + 3600, ttl=300 if op == "d" else 180, **kw)


def test_attachment_streams_bytes(rt, graph, client):
    graph.on("GET", "/me/messages/m1/attachments/a1/\\$value", b"%PDF-1.7 data")
    r = client.get(f"/d/{link(rt, 'd', 'attachment', {'m': 'm1', 'a': 'a1', 'n': 'q3.pdf'})}")
    assert r.status_code == 200 and r.content == b"%PDF-1.7 data"
    assert "q3.pdf" in r.headers["content-disposition"]


@pytest.mark.parametrize("value", ["", "garbage", "a.b", "e30.AAAA"])
def test_bad_links_are_403_not_500(rt, client, value):
    assert client.get(f"/d/{value}").status_code in (403, 404)
    assert client.put(f"/u/{value}", content=b"x").status_code in (403, 404, 405)


def test_download_link_cannot_upload(rt, client):
    v = link(rt, "d", "attachment", {"m": "m1", "a": "a1"})
    assert client.put(f"/u/{v}", content=b"x").status_code == 403


def test_expired_link_is_403(rt, client, graph):
    from helpers import make_token
    v = rt.signer.issue(op="d", kind="attachment", target={"m": "m", "a": "a"}, person="p", token=make_token(),
                        token_expires=int(time.time()) + 3600, ttl=300, now=time.time() - 1000)
    assert client.get(f"/d/{v}").status_code == 403
    assert graph.seen == []


def test_draft_body_upload_is_cleaned_and_keeps_the_thread(rt, graph, client):
    quote = '<html><body><div id="divRplyFwdMsg">From: Bo</div><p>old</p></body></html>'
    graph.on("GET", "/me/messages/d1", {"id": "d1", "isDraft": True, "body": {"content": quote}})
    graph.on("PATCH", "/me/messages/d1", {})
    graph.on("POST", "/me/messages/d1/attachments", {"id": "x"})
    png = base64.b64encode(b"\x89PNG" + b"\0" * 20).decode()
    html = f'<style>p{{color:red}}</style><p onclick="x()">Report</p><script>evil()</script><img alt="c" src="data:image/png;base64,{png}">'
    v = link(rt, "u", "draft_body", {"d": "d1"}, max_bytes=2 * 1024 * 1024, content_type="text/html")
    r = client.put(f"/u/{v}", content=html.encode(), headers={"Content-Type": "text/html; charset=utf-8"})
    assert r.status_code == 200, r.text
    body = graph.calls("PATCH", "/me/messages/d1")[0].json["body"]["content"]
    assert "evil" not in body and "onclick" not in body and "cid:" in body and "color:red" in body.replace(" ", "")
    assert body.endswith('<div id="divRplyFwdMsg">From: Bo</div><p>old</p></body></html>')
    assert f'id="{NEW_PART_ID}"' in body
    assert r.json()["html"]["sanitized"]


def test_draft_body_not_utf8_is_rejected(rt, client, graph):
    v = link(rt, "u", "draft_body", {"d": "d1"}, max_bytes=10_000, content_type="text/html")
    r = client.put(f"/u/{v}", content="<p>Grüße</p>".encode("latin-1"), headers={"Content-Type": "text/html"})
    assert r.status_code == 422 and "UTF-8" in r.json()["error"] and graph.writes() == []


def test_draft_body_wrong_content_type_is_415(rt, client):
    v = link(rt, "u", "draft_body", {"d": "d1"}, max_bytes=100, content_type="text/html")
    r = client.put(f"/u/{v}", content=b"x", headers={"Content-Type": "application/pdf"})
    # The refusal named nothing; the type the link takes is no one's data.
    assert r.status_code == 415 and "text/html" in r.json()["error"]


def test_draft_body_over_the_limit_is_413(rt, client, graph):
    v = link(rt, "u", "draft_body", {"d": "d1"}, max_bytes=100, content_type="text/html")
    r = client.put(f"/u/{v}", content=b"<p>" + b"a" * 200 + b"</p>", headers={"Content-Type": "text/html"})
    assert r.status_code == 413 and graph.writes() == []
    assert "at most 100 bytes" in r.json()["error"]  # The limit is named


def test_draft_body_with_relative_urls_is_rejected(rt, client, graph):
    v = link(rt, "u", "draft_body", {"d": "d1"}, max_bytes=10_000, content_type="text/html")
    r = client.put(f"/u/{v}", content=b'<img alt="x" src="/a.png">', headers={"Content-Type": "text/html"})
    assert r.status_code == 422 and "absolute" in r.json()["error"] and graph.writes() == []


def test_small_attachment_is_one_post(rt, graph, client):
    graph.on("POST", "/me/messages/d1/attachments", {"id": "a"})
    v = link(rt, "u", "draft_attachment", {"d": "d1"}, max_bytes=150 * 1024 * 1024)
    r = client.put(f"/u/{v}", content=b"hello", headers={"X-File-Name": "notes.txt", "Content-Type": "text/plain"})
    assert r.status_code == 200, r.text
    att = graph.calls("POST", "/me/messages/d1/attachments")[0].json
    assert att["name"] == "notes.txt" and base64.b64decode(att["contentBytes"]) == b"hello"


def test_large_attachment_uses_an_upload_session_in_ranges(rt, graph, client):
    size = 10 * 1024 * 1024
    graph.on("POST", "/me/messages/d1/attachments/createUploadSession", {"uploadUrl": "https://outlook.office.com/up/1"})
    ranges = []

    def put(seen):
        ranges.append((seen.headers["content-range"], len(seen.content)))
        return httpx.Response(200 if len(ranges) * ATTACH_CHUNK < size else 201, json={})
    graph.on("PUT", "https://outlook.office.com/up/1", put)
    v = link(rt, "u", "draft_attachment", {"d": "d1"}, max_bytes=150 * 1024 * 1024)
    r = client.put(f"/u/{v}", content=b"\1" * size, headers={"X-File-Name": "big.bin"})
    assert r.status_code == 200, r.text
    assert all(n <= 4 * 1024 * 1024 for _, n in ranges)
    assert ranges[0][0] == f"bytes 0-{ATTACH_CHUNK - 1}/{size}"
    assert ranges[-1][0].endswith(f"-{size - 1}/{size}")
    assert sum(n for _, n in ranges) == size
    assert not [s for s in graph.seen if s.path.startswith("https://outlook.office.com") and "authorization" in s.headers]


def test_an_attachment_over_the_link_limit_is_413(rt, client, graph):
    v = link(rt, "u", "draft_attachment", {"d": "d1"}, max_bytes=1000)
    r = client.put(f"/u/{v}", content=b"\1" * 1001, headers={"X-File-Name": "big.bin"})
    assert r.status_code == 413 and graph.writes() == []


async def test_draft_attachment_links_carry_the_150_mb_limit(graph, invoke, rt):
    graph.on("POST", "/me/messages", {"id": "d1"})
    graph.on("GET", "/me/messages/d1", {"id": "d1", "isDraft": True, "body": {"content": ""}})
    r = await invoke("save_draft", subject="S", attachments_upload=1)
    assert not r.error, r.text
    value = r.data["attachment_uploads"][0].rsplit("/u/", 1)[1].rstrip("'")
    assert rt.signer.verify(value, op="u").max_bytes == 150 * 1024 * 1024


def test_an_attachment_past_the_message_size_limit_is_refused(rt, client, graph):
    graph.on("POST", "/me/messages/d1/attachments", {"error": {"code": "ErrorMessageSizeExceeded"}}, status=413)
    v = link(rt, "u", "draft_attachment", {"d": "d1"}, max_bytes=150 * 1024 * 1024)
    r = client.put(f"/u/{v}", content=b"x" * 100, headers={"X-File-Name": "a.pdf"})
    assert r.status_code == 400 and "larger than Microsoft 365 accepts" in r.json()["error"]


def test_attachment_without_a_file_name_is_rejected(rt, client, graph):
    v = link(rt, "u", "draft_attachment", {"d": "d1"}, max_bytes=1000)
    assert client.put(f"/u/{v}", content=b"x").status_code == 422
    assert graph.writes() == []


def test_note_title_is_escaped_not_json_encoded(rt, graph, client):
    import httpx as hx
    seen = {}

    def create(s):
        seen["body"] = s.content.decode()
        return hx.Response(201, json={"id": "p1"})
    graph.on("POST", "/me/onenote/sections/S/pages", create)
    v = link(rt, "u", "note_body", {"s": "S", "title": "</title><img src=https://x>Grüße ü"}, max_bytes=10_000,
             content_type="text/html")
    r = client.put(f"/u/{v}", content=b"<p>x</p>", headers={"Content-Type": "text/html"})
    assert r.status_code == 200, r.text
    assert "<img src=https://x>" not in seen["body"] and "&lt;/title&gt;" in seen["body"] and "Grüße ü" in seen["body"]


def test_an_uploaded_page_body_keeps_the_page_title(rt, graph, client):
    # Measured: a body PATCH soon after creation empties the listed title unless the title goes along.
    from datetime import UTC, datetime
    graph.on("GET", "/me/onenote/pages/p1/content", b"<html><head><title>Plan</title></head><body></body></html>")
    graph.on("GET", "/me/onenote/pages/p1", {"id": "p1", "title": "", "createdDateTime": datetime.now(UTC).isoformat()})
    graph.on("PATCH", "/me/onenote/pages/p1/content", None, status=204)
    v = link(rt, "u", "note_body", {"page": "p1", "a": "append"}, max_bytes=10_000, content_type="text/html")
    r = client.put(f"/u/{v}", content=b"<p>x</p>", headers={"Content-Type": "text/html"})
    assert r.status_code == 200, r.text
    assert graph.writes()[0].json == [{"target": "title", "action": "replace", "content": "Plan"},
                                      {"target": "body", "action": "append", "content": "<p>x</p>"}]


def test_an_upstream_answer_without_an_upload_url_is_502_not_500(rt, graph, client):
    graph.on("POST", "/me/messages/d1/attachments/createUploadSession", {})
    v = link(rt, "u", "draft_attachment", {"d": "d1"}, max_bytes=150 * 1024 * 1024)
    r = client.put(f"/u/{v}", content=b"\1" * (4 * 1024 * 1024), headers={"X-File-Name": "big.bin"})
    assert r.status_code == 502


def test_an_uploaded_body_replaces_the_drafts_own_part_and_keeps_a_quoted_one(rt, graph, client):
    from m365_mcp.tools.mail import DRAFT_MARK
    part = f"{NEW_PART_ID}-abc123abc123"
    body = (f'<html><body><div id="{part}"><p>first</p></div><hr><div id="divRplyFwdMsg">From: Bo</div>'
            f'<div><div id="{NEW_PART_ID}"><p>QUOTED LETTER</p></div></div></body></html>')
    graph.on("GET", "/me/messages/d1", {"id": "d1", "isDraft": True, "body": {"content": body},
                                        "singleValueExtendedProperties": [{"id": DRAFT_MARK, "value": part}]})
    graph.on("PATCH", "/me/messages/d1", {})
    v = link(rt, "u", "draft_body", {"d": "d1"}, max_bytes=10_000, content_type="text/html")
    r = client.put(f"/u/{v}", content=b"<p>second</p>", headers={"Content-Type": "text/html"})
    assert r.status_code == 200, r.text
    out = graph.calls("PATCH", "/me/messages/d1")[0].json["body"]["content"]
    assert "first" not in out and "second" in out and "QUOTED LETTER" in out


def test_a_body_cannot_be_uploaded_to_a_draft_already_sent(rt, graph, client):
    from m365_mcp.tools.mail import DRAFT_MARK
    graph.on("GET", "/me/messages/d1", {"id": "d1", "isDraft": True, "body": {"content": ""},
                                        "singleValueExtendedProperties": [{"id": DRAFT_MARK, "value": "sent"}]})
    v = link(rt, "u", "draft_body", {"d": "d1"}, max_bytes=10_000, content_type="text/html")
    r = client.put(f"/u/{v}", content=b"<p>late</p>", headers={"Content-Type": "text/html"})
    assert r.status_code == 422 and "already sent" in r.text and graph.writes() == []


@pytest.mark.parametrize("name", ["Reunión_presupuesto.docx", "family \U0001f468\u200d\U0001f469\u200d\U0001f467.png", "qa7-\u043e\u0442\u0447\u0451\u0442.txt", "emoji 📎.txt"])
def test_a_utf8_file_name_is_kept_as_sent(rt, graph, client, name):
    # curl sends the header's UTF-8 bytes; read as Latin-1 they became
    # "ReuniÃ³n_presupuesto.docx", in the draft and at the recipient.
    graph.on("POST", "/me/messages/d1/attachments", {"id": "a"})
    v = link(rt, "u", "draft_attachment", {"d": "d1"}, max_bytes=150 * 1024 * 1024)
    r = client.put(f"/u/{v}", content=b"hello", headers=[(b"X-File-Name", name.encode())])
    assert r.status_code == 200, r.text
    assert graph.calls("POST", "/me/messages/d1/attachments")[0].json["name"] == name


def test_a_percent_encoded_file_name_is_decoded(rt, graph, client):
    # The way round for a client that cannot send raw UTF-8 in a header: RFC 5987's UTF-8''<percent>.
    graph.on("POST", "/me/messages/d1/attachments", {"id": "a"})
    v = link(rt, "u", "draft_attachment", {"d": "d1"}, max_bytes=150 * 1024 * 1024)
    r = client.put(f"/u/{v}", content=b"hello", headers={"X-File-Name": "UTF-8''Reuni%C3%B3n_pct.txt"})
    assert r.status_code == 200, r.text
    assert graph.calls("POST", "/me/messages/d1/attachments")[0].json["name"] == "Reunión_pct.txt"



def test_header_bytes_that_are_not_utf8_stay_latin1():
    from m365_mcp.links.routes import _header_text
    assert _header_text(b"caf\xe9.txt".decode("latin-1")) == "caf\xe9.txt"  # not valid UTF-8
    assert _header_text("Reunión.docx".encode().decode("latin-1")) == "Reunión.docx"


def test_a_bidi_override_is_dropped_from_a_file_name(rt, graph, client):
    graph.on("POST", "/me/messages/d1/attachments", {"id": "a"})
    v = link(rt, "u", "draft_attachment", {"d": "d1"}, max_bytes=150 * 1024 * 1024)
    r = client.put(f"/u/{v}", content=b"x", headers=[(b"X-File-Name", "invoice\u202efdp.exe".encode())])
    assert r.status_code == 200, r.text
    assert graph.calls("POST", "/me/messages/d1/attachments")[0].json["name"] == "invoicefdp.exe"


@pytest.mark.parametrize("name, sent, want", [("scan.pdf", "application/octet-stream", "application/pdf"),
                                               ("banner.png", None, "image/png"),
                                               ("data.bin", "application/octet-stream", "application/octet-stream"),
                                               ("a.pdf", "text/plain", "text/plain")])
def test_an_attachment_without_a_real_type_gets_one_from_its_name(rt, graph, client, name, sent, want):
    # Every attachment arrived as application/octet-stream.
    graph.on("POST", "/me/messages/d1/attachments", {"id": "a"})
    v = link(rt, "u", "draft_attachment", {"d": "d1"}, max_bytes=150 * 1024 * 1024)
    headers = {"X-File-Name": name, **({"Content-Type": sent} if sent else {})}
    r = client.put(f"/u/{v}", content=b"x", headers=headers)
    assert r.status_code == 200, r.text
    assert graph.calls("POST", "/me/messages/d1/attachments")[0].json["contentType"] == want


@pytest.mark.parametrize("state", ["sent_mark", "not_a_draft"])
def test_an_attachment_cannot_be_uploaded_to_a_draft_already_sent(rt, graph, client, state):
    from m365_mcp.tools.mail import DRAFT_MARK
    graph.on("GET", "/me/messages/d1", {"id": "d1", "isDraft": state != "not_a_draft", "singleValueExtendedProperties": [
        {"id": DRAFT_MARK, "value": "sent" if state == "sent_mark" else "1"}]})
    graph.on("POST", "/me/messages/d1/attachments", {"id": "a"})
    v = link(rt, "u", "draft_attachment", {"d": "d1"}, max_bytes=1000)
    r = client.put(f"/u/{v}", content=b"late", headers={"X-File-Name": "late.txt"})
    assert r.status_code == 422 and ("already sent" in r.text or "not a draft" in r.text)
    assert graph.writes() == []


async def test_sending_a_draft_revokes_its_upload_links(graph, invoke, rt):
    from m365_mcp.tools.mail import DRAFT_MARK
    graph.on("POST", "/me/messages", {"id": "d1"})
    graph.on("PATCH", "/me/messages/d1", {})
    graph.on("POST", "/me/messages/d1/send", None, status=202)
    graph.on("GET", "/me/messages/d1", {"id": "d1", "isDraft": True, "subject": "S", "body": {"content": ""},
                                        "toRecipients": [{"emailAddress": {"address": "ann@x.test"}}],
                                        "singleValueExtendedProperties": [{"id": DRAFT_MARK, "value": "1"}]})
    r = await invoke("save_draft", subject="S", to=["ann@x.test"], body_upload=True, attachments_upload=2)
    assert not r.error, r.text
    values = [cmd.rsplit("/u/", 1)[1].rstrip("'") for cmd in [r.data["body_upload"], *r.data["attachment_uploads"]]]
    other = link(rt, "u", "draft_attachment", {"d": "d2"}, max_bytes=1000)
    assert not (await invoke("send_draft", id="d1", user_confirmed=True)).error
    from m365_mcp.links.tokens import LinkRefused
    for v in values:
        with pytest.raises(LinkRefused):
            rt.signer.verify(v, op="u")
    assert rt.signer.verify(other, op="u").target == {"d": "d2"}, "another draft's link is untouched"


def _draft_with_links(graph):
    from m365_mcp.tools.mail import DRAFT_MARK
    graph.on("POST", "/me/messages", {"id": "d1"})
    graph.on("GET", "/me/messages/d1", {"id": "d1", "isDraft": True, "subject": "S", "body": {"content": ""},
                                        "toRecipients": [{"emailAddress": {"address": "ann@x.test"}}],
                                        "singleValueExtendedProperties": [{"id": DRAFT_MARK, "value": "1"}]})


async def _links(invoke) -> list[str]:
    r = await invoke("save_draft", subject="S", to=["ann@x.test"], body_upload=True, attachments_upload=2)
    assert not r.error, r.text
    return [cmd.rsplit("/u/", 1)[1].rstrip("'") for cmd in [r.data["body_upload"], *r.data["attachment_uploads"]]]


async def test_send_draft_refuses_while_an_upload_to_the_draft_is_running(graph, invoke, rt):
    _draft_with_links(graph)
    graph.on("PATCH", "/me/messages/d1", {})
    graph.on("POST", "/me/messages/d1/send", None, status=202)
    values = await _links(invoke)
    assert rt.signer.claim(values[1])  # a long attachment upload, past its link check
    graph.seen.clear()
    r = await invoke("send_draft", id="d1", user_confirmed=True)
    assert r.error and "upload to this draft is still running" in r.text and "send_draft again" in r.text
    assert graph.writes() == []
    rt.signer.consume(values[1], time.time() + 180)  # it finished: now the draft sends
    assert not (await invoke("send_draft", id="d1", user_confirmed=True)).error


async def test_an_upload_starting_while_the_draft_is_marked_sent_stops_the_send(graph, invoke, rt):
    _draft_with_links(graph)
    values = await _links(invoke)

    def mark(seen):
        # The upload claims its link while send_draft's mark is on its way, and read the draft before it.
        if "sent" in str(seen.json):
            rt.signer.claim(values[1])
        return httpx.Response(200, json={})
    graph.on("PATCH", "/me/messages/d1", mark)
    graph.on("POST", "/me/messages/d1/send", None, status=202)
    r = await invoke("send_draft", id="d1", user_confirmed=True)
    assert r.error and "still running" in r.text
    assert graph.calls("POST", "/me/messages/d1/send") == []
    marks = [w.json for w in graph.writes() if w.method == "PATCH"]
    assert len(marks) == 2 and "sent" not in str(marks[-1]), "the sent mark is taken back"


async def test_a_send_that_fails_leaves_the_upload_links_usable(graph, invoke, rt):
    _draft_with_links(graph)
    graph.on("PATCH", "/me/messages/d1", {})
    graph.on("POST", "/me/messages/d1/send", {"error": {"code": "ErrorInvalidRecipients", "message": "x"}}, status=400)
    values = await _links(invoke)
    r = await invoke("send_draft", id="d1", user_confirmed=True)
    assert r.error
    for v in values:
        assert rt.signer.verify(v, op="u").target == {"d": "d1"}
