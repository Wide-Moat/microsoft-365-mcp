"""A fake Microsoft Graph and a configured server, for every test.

The fake answers only the routes a test registers and records every request, so a test can
assert exactly what reached Graph — including that nothing did.
"""

from __future__ import annotations

import base64
import json
import re
import time
from dataclasses import dataclass, field
from typing import Any

import httpx
import pytest

from m365_mcp.config import Settings
from m365_mcp.core.context import Runtime, configure
from m365_mcp.core.cursor import Cursors
from m365_mcp.core.guard import Guard
from m365_mcp.identity import current_areas, current_token
from m365_mcp.links.store import LinkStore
from m365_mcp.links.tokens import LinkSigner

from helpers import ENC, LINKS, SIGN, FakeGraph, make_token  # noqa: F401


async def _no_sleep(_s: float) -> None:
    return None


@pytest.fixture
def graph() -> FakeGraph:
    g = FakeGraph()
    g.on("GET", "/me/mailboxSettings/timeZone", {"value": "UTC"})
    # A small organisation at x.test where every address asked about is a user, nobody is in the
    # People index and there are no personal contacts. Tests about groups, names or contacts
    # register their own routes, which take precedence.
    g.on("GET", "/organization", {"value": [{"verifiedDomains": [{"name": "x.test"}]}]})
    g.on("GET", "/users", _everyone_is_a_user)
    g.on("GET", "/me/people", {"value": []})
    g.on("GET", "/me/contacts", {"value": []})
    g.on("GET", "/me/contactFolders", {"value": []})
    g.on("GET", "/me/messages/[^/]+/attachments", {"value": []})
    return g


def _everyone_is_a_user(seen) -> httpx.Response:
    m = re.search(r"mail eq '((?:[^']|'')*)'", seen.params.get("$filter", ""))
    if not m:
        return httpx.Response(200, json={"value": []})
    addr = m.group(1).replace("''", "'")
    return httpx.Response(200, json={"value": [{"displayName": "", "mail": addr, "userPrincipalName": addr}]})


@pytest.fixture
def settings() -> Settings:
    from m365_mcp.registry import AREAS
    return Settings(areas=frozenset(a for a, on in AREAS.items() if on) | {"bookings"},
                    switches=frozenset({"teams_channels", "transcripts", "rooms", "chat_admin"}),
                    sign_keys=(SIGN,), enc_keys=(ENC,), links_base_url=LINKS)


@pytest.fixture
def guard() -> Guard:
    return Guard()


@pytest.fixture
def rt(graph: FakeGraph, settings: Settings, guard: Guard) -> Runtime:
    r = Runtime(settings=settings, signer=LinkSigner((SIGN,), (ENC,), LinkStore()), cursors=Cursors(SIGN), guard=guard,
                transport=httpx.MockTransport(graph), sleep=_no_sleep)
    configure(r)
    return r


@pytest.fixture
def token(rt: Runtime) -> str:
    t = make_token()
    tok = current_token.set(t)
    areas = current_areas.set(None)
    yield t
    current_token.reset(tok)
    current_areas.reset(areas)


@pytest.fixture
def server(rt: Runtime, settings: Settings):
    from m365_mcp.server import M365Server
    return M365Server(settings)


@dataclass
class Result:
    error: bool
    text: str
    data: Any


@pytest.fixture
def invoke(server, token):
    """Call a tool the way the MCP handler does, and return the error flag, text and structured data."""
    from mcp.types import CallToolRequestParams

    async def _invoke(_tool: str, **args: Any) -> Result:
        params = CallToolRequestParams(name=_tool, arguments=args)
        res = await server._handle_call_tool(None, params)
        text = "".join(getattr(c, "text", "") for c in res.content)
        data = res.structured_content
        if isinstance(data, dict) and set(data) == {"result"}:
            data = data["result"]
        if not res.is_error:
            # Every tool declares structured output: a success without it is the copy Claude Code shows gone.
            assert res.structured_content is not None, f"{_tool}: a result with no structured content"
            # A client shows the model one of the two copies, Claude Code the
            # structured one. They must be the same result, link for link, or that client loses the
            # command the other one carries. Checked on every call every test makes.
            assert json.loads(text) == data, f"{_tool}: the text and the structured content differ"
        return Result(bool(res.is_error), text, data)

    return _invoke
