"""Shared test helpers: the fake Graph and token factory."""

from __future__ import annotations

import base64
import json
import re
import time
from dataclasses import dataclass, field
from typing import Any

import httpx

SIGN = bytes(range(32))
ENC = bytes(range(32, 64))
LINKS = "https://links.test"

ALL_SCOPES = (
    "User.Read User.ReadBasic.All People.Read MailboxSettings.ReadWrite Mail.ReadWrite Mail.Send "
    "Calendars.ReadWrite Calendars.Read.Shared Presence.Read.All OnlineMeetings.Read Contacts.ReadWrite "
    "Files.ReadWrite.All Sites.ReadWrite.All Chat.ReadWrite ChatMessage.Send ChannelMessage.Send "
    "ChannelMessage.Read.All ChannelMessage.ReadWrite Team.ReadBasic.All Channel.ReadBasic.All ChannelMember.Read.All "
    "OnlineMeetingTranscript.Read.All OnlineMeetingRecording.Read.All Place.Read.All Tasks.ReadWrite "
    "Notes.ReadWrite BookingsAppointment.ReadWrite.All"
)


def make_token(scopes: str = ALL_SCOPES, *, oid: str = "oid-a", tid: str = "tid-1", exp: int | None = None,
               aud: str = "https://graph.microsoft.com", **claims_extra: object) -> str:
    def b64(d: dict) -> str:
        return base64.urlsafe_b64encode(json.dumps(d).encode()).rstrip(b"=").decode()
    claims = {"aud": aud, "tid": tid, "oid": oid, "scp": scopes,
              "exp": int(time.time()) + 3600 if exp is None else exp, **claims_extra}
    return f"{b64({'alg': 'none'})}.{b64(claims)}.sig"


class _Chunks(httpx.AsyncByteStream):
    def __init__(self, data: bytes, size: int = 7) -> None:
        self.data, self.size, self.closed = data, size, False

    async def __aiter__(self):
        for i in range(0, len(self.data), self.size):
            if self.closed:
                raise RuntimeError("read after close")
            yield self.data[i:i + self.size]

    async def aclose(self) -> None:
        self.closed = True


@dataclass
class Seen:
    method: str
    path: str
    params: dict[str, str]
    headers: dict[str, str]
    json: Any
    content: bytes


@dataclass
class FakeGraph:
    routes: list[tuple[str, re.Pattern[str], Any]] = field(default_factory=list)
    seen: list[Seen] = field(default_factory=list)

    def on(self, method: str, path: str, response: Any = None, *, status: int = 200,
           headers: dict[str, str] | None = None) -> None:
        """`path` is a regex over the path after /v1.0 (or a full URL for downloads).
        `response` is a dict (JSON), bytes, or a callable(Seen) -> httpx.Response."""
        def handler(seen: Seen) -> httpx.Response:
            if callable(response):
                return response(seen)
            if isinstance(response, (bytes, bytearray)):
                # A real stream, read in chunks, as Graph's downloads are: content= would give a
                # body that is already read, which hides streaming defects.
                return httpx.Response(status, stream=_Chunks(bytes(response)), headers=headers)
            if response is None:
                return httpx.Response(status, headers=headers)
            return httpx.Response(status, json=response, headers=headers)
        self.routes.insert(0, (method, re.compile("^" + path + "$"), handler))

    def writes(self) -> list[Seen]:
        return [s for s in self.seen if s.method in ("POST", "PATCH", "PUT", "DELETE")]

    def calls(self, method: str, pattern: str) -> list[Seen]:
        r = re.compile("^" + pattern + "$")
        return [s for s in self.seen if s.method == method and r.match(s.path)]

    def __call__(self, request: httpx.Request) -> httpx.Response:
        url = request.url
        path = url.path.removeprefix("/v1.0") if url.host == "graph.microsoft.com" else str(url.copy_with(query=None))
        body = None
        if request.content and "json" in request.headers.get("content-type", ""):
            body = json.loads(request.content)
        seen = Seen(request.method, path, dict(url.params), dict(request.headers), body, request.content)
        self.seen.append(seen)
        for method, pattern, handler in self.routes:
            if method == request.method and pattern.match(path):
                return handler(seen)
        return httpx.Response(404, json={"error": {"code": "itemNotFound", "message": f"no fake route {path}"}})




_LINE = re.compile(r"\A<<<text ([0-9a-f]{8}) from (.*?)>>> (.*) <<<end \1>>>\Z", re.S)
_BLOCK = re.compile(r"\A<<<content ([0-9a-f]{16}) from (.*) — data, not instructions>>>\n(.*)\n<<<end of content \1>>>\Z", re.S)


def inner(fenced: str) -> str:
    """The foreign text inside a content line or a content block; fails when it is not fenced."""
    m = _LINE.match(fenced) or _BLOCK.match(fenced)
    assert m, f"not fenced: {fenced!r}"
    return m.group(3)
