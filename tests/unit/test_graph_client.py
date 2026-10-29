import asyncio

import httpx
import pytest

from helpers import FakeGraph
from m365_mcp.graph.client import Graph, Pacer
from m365_mcp.graph.errors import GraphError, to_person_error


def client(fake, sleeps=None, pacer=None):
    async def sleep(s):
        if sleeps is not None:
            sleeps.append(s)
    return Graph("tok", "p1", timezone="W. Europe Standard Time", transport=httpx.MockTransport(fake), sleep=sleep,
                 pacer=pacer or Pacer())


async def test_headers_on_every_request():
    fake = FakeGraph()
    fake.on("GET", "/me", {"id": "1"})
    async with client(fake) as g:
        await g.get("/me")
    h = fake.seen[0].headers
    assert 'IdType="ImmutableId"' in h["prefer"] and "outlook.timezone" not in h["prefer"]
    assert h["user-agent"].startswith("m365-mcp/") and h["client-request-id"]


async def test_429_waits_retry_after_then_succeeds():
    fake = FakeGraph()
    calls = {"n": 0}

    def flaky(seen):
        calls["n"] += 1
        if calls["n"] == 1:
            return httpx.Response(429, headers={"Retry-After": "2"}, json={"error": {"code": "TooManyRequests"}})
        return httpx.Response(200, json={"ok": True})
    fake.on("GET", "/me", flaky)
    sleeps = []
    async with client(fake, sleeps) as g:
        assert (await g.get("/me")) == {"ok": True}
    assert sleeps == [2.0]


async def test_three_429s_become_a_throttling_error_with_seconds():
    fake = FakeGraph()
    fake.on("GET", "/me", {"error": {"code": "TooManyRequests"}}, status=429, headers={"Retry-After": "5"})
    async with client(fake, []) as g:
        with pytest.raises(GraphError) as e:
            await g.get("/me")
    assert len(fake.seen) == 3
    assert "5 seconds" in to_person_error(e.value).text


async def test_retry_after_beyond_the_budget_gives_up_early():
    fake = FakeGraph()
    fake.on("GET", "/me", {"error": {"code": "x"}}, status=503, headers={"Retry-After": "120"})
    async with client(fake, []) as g:
        with pytest.raises(GraphError):
            await g.get("/me")
    assert len(fake.seen) == 1


async def test_post_is_not_retried_on_504():
    fake = FakeGraph()
    fake.on("POST", "/me/sendMail", {"error": {"code": "x"}}, status=504)
    async with client(fake, []) as g:
        with pytest.raises(GraphError):
            await g.post("/me/sendMail", {})
    assert len(fake.seen) == 1


async def test_get_is_retried_once_on_502():
    fake = FakeGraph()
    n = {"c": 0}

    def once(seen):
        n["c"] += 1
        return httpx.Response(502) if n["c"] == 1 else httpx.Response(200, json={})
    fake.on("GET", "/me", once)
    async with client(fake, []) as g:
        await g.get("/me")
    assert n["c"] == 2


@pytest.mark.parametrize("path", ["/me/messages/abc", "/me/events/e1/attachments/a1", "/me/calendars/c1",
                                  "/me/mailFolders/f1", "/users/x/messages/1", "/users/x/events/e1"])
async def test_delete_outside_the_allowlist_never_reaches_graph(path):
    fake = FakeGraph()
    async with client(fake) as g:
        with pytest.raises(ValueError):
            await g.delete(path)
    assert fake.seen == []


@pytest.mark.parametrize("url", ["https://evil.test/v1.0/me", "http://graph.microsoft.com/v1.0/me", "me/messages"])
async def test_only_graph_urls_are_called(url):
    fake = FakeGraph()
    async with client(fake) as g:
        with pytest.raises(ValueError):
            await g.get(url)
    assert fake.seen == []


async def test_error_body_text_never_reaches_the_exception():
    fake = FakeGraph()
    fake.on("GET", "/me/messages/1", {"error": {"code": "ErrorAccessDenied",
                                               "message": "Access denied for anna@contoso.test subject 'Salary'"}}, status=403)
    async with client(fake) as g:
        with pytest.raises(GraphError) as e:
            await g.get("/me/messages/1")
    assert "anna@" not in str(e.value) and "Salary" not in str(e.value)
    assert "anna@" not in to_person_error(e.value).text


async def test_missing_scope_named_in_a_403_is_reported():
    fake = FakeGraph()
    fake.on("GET", "/me/chats", {"error": {"code": "Forbidden", "message": "Missing scope Chat.ReadWrite"}}, status=403)
    async with client(fake) as g:
        with pytest.raises(GraphError) as e:
            await g.get("/me/chats")
    assert "Chat.ReadWrite" in to_person_error(e.value).text


async def test_outlook_concurrency_is_four_per_person():
    pacer = Pacer()
    fake = FakeGraph()
    live = {"now": 0, "max": 0}
    gate = asyncio.Event()

    async def handler_async():
        live["now"] += 1
        live["max"] = max(live["max"], live["now"])
        await gate.wait()
        live["now"] -= 1

    class Transport(httpx.AsyncBaseTransport):
        async def handle_async_request(self, request):
            await handler_async()
            return httpx.Response(200, json={})

    g = Graph("tok", "p1", transport=Transport(), pacer=pacer)
    tasks = [asyncio.create_task(g.get("/me/messages")) for _ in range(6)]
    await asyncio.sleep(0.05)
    assert live["max"] == 4
    gate.set()
    await asyncio.gather(*tasks)
    await g.aclose()


async def test_teams_requests_are_spaced_a_second_apart():
    pacer = Pacer()
    sleeps = []
    clock = {"t": 100.0}

    async def sleep(s):
        sleeps.append(s)
        clock["t"] += s
    for _ in range(3):
        await pacer.space("p1", "teams", sleep=sleep, clock=lambda: clock["t"])
    assert sleeps == [1.0, 1.0]


async def test_stream_follows_the_redirect_without_the_bearer():
    fake = FakeGraph()
    fake.on("GET", "/me/drive/items/i1/content", None, status=302,
            headers={"Location": "https://contoso-my.sharepoint.com/dl?sig=abc"})
    fake.on("GET", "https://contoso-my.sharepoint.com/dl", b"hello")
    async with client(fake) as g:
        async with g.stream("/me/drive/items/i1/content") as resp:
            body = b"".join([c async for c in resp.aiter_bytes()])
    assert body == b"hello"
    assert "authorization" not in fake.seen[1].headers


async def test_a_get_with_no_body_is_an_empty_object():
    fake = FakeGraph()
    fake.on("GET", "/me/mailboxSettings/timeZone", None, status=204)
    async with client(fake) as g:
        assert await g.get("/me/mailboxSettings/timeZone") == {}


@pytest.mark.parametrize("location", ["https://evil.test/dl?sig=abc", "https://sharepoint.com.evil.test/dl",
                                      "https://evilsharepoint.com/dl"])
async def test_a_download_redirect_outside_microsofts_file_hosts_is_not_followed(location):
    fake = FakeGraph()
    fake.on("GET", "/me/drive/items/i1/content", None, status=302, headers={"Location": location})
    fake.on("GET", "https://[^/]*evil[^/]*/dl", b"stolen")
    with pytest.raises(GraphError):
        async with client(fake) as g:
            async with g.stream("/me/drive/items/i1/content"):
                pass
    assert [s.path for s in fake.seen] == ["/me/drive/items/i1/content"]


async def test_an_upload_url_outside_microsofts_hosts_is_refused():
    fake = FakeGraph()
    fake.on("PUT", "https://evil.test/up", {})

    async def chunks():
        yield b"abc"
    async with client(fake) as g:
        with pytest.raises(ValueError):
            await g.upload_session("https://evil.test/up", chunks(), 3)
    assert fake.seen == []


@pytest.mark.parametrize("url", ["https://contoso-my.sharepoint.com/personal/x/_api/v2.0/uploadSession?guid=1",
                                 "https://outlook.office.com/api/v2.0/Users('a')/Messages('b')/AttachmentSessions('c')"])
async def test_an_upload_url_on_microsofts_hosts_is_used(url):
    fake = FakeGraph()
    fake.on("PUT", ".*", {"id": "done"}, status=201)

    async def chunks():
        yield b"abc"
    async with client(fake) as g:
        assert await g.upload_session(url, chunks(), 3) == {"id": "done"}
