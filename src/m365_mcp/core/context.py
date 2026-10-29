"""What a tool call works with: the person, a Graph client, their time zone, links, the guard.

`async with call("tool_name") as c:` opens it, converts any Graph failure into the sentence a
person reads, and closes the HTTP client. Tools never see `GraphError`.
"""

from __future__ import annotations

import time
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from dataclasses import dataclass, field
from typing import Any
from zoneinfo import ZoneInfo

import httpx

from m365_mcp.config import LINK_TTL_DOWNLOAD, LINK_TTL_UPLOAD, Settings
from m365_mcp.core.cursor import Cursors
from m365_mcp.core.errors import M365Error
from m365_mcp.core.guard import GUARD, Guard
from m365_mcp.core.time import LEARNED_ZONES, known, zone  # noqa: F401
from m365_mcp.graph.client import Graph
from m365_mcp.graph.errors import GraphError, to_person_error
from m365_mcp.identity import Person, current_person
from m365_mcp.links.store import LinksFull
from m365_mcp.links.tokens import LinkSigner


@dataclass
class Runtime:
    settings: Settings
    signer: LinkSigner | None
    cursors: Cursors
    guard: Guard = field(default_factory=lambda: GUARD)
    transport: httpx.AsyncBaseTransport | None = None
    sleep: Any = None
    tz_cache: dict[str, tuple[str, float]] = field(default_factory=dict)
    # tenant -> (its verified mail domains, expiry); see core/resolve.py
    domain_cache: dict[str, tuple[frozenset[str], float]] = field(default_factory=dict)


_RUNTIME: Runtime | None = None


def configure(rt: Runtime) -> None:
    global _RUNTIME
    _RUNTIME = rt


def runtime() -> Runtime:
    if _RUNTIME is None:
        raise RuntimeError("runtime not configured")
    return _RUNTIME


_LEVEL = {"ReadBasic": 0, "Read": 1, "ReadWrite": 2}


def _parse_scope(s: str) -> tuple[str, str, bool]:
    parts = s.split(".")
    is_all = parts[-1] == "All"
    if is_all:
        parts = parts[:-1]
    return parts[0], ".".join(parts[1:]), is_all


def scope_granted(required: str, granted: frozenset[str]) -> bool:
    if required in granted:
        return True
    r_res, r_lvl, r_all = _parse_scope(required)
    for g in granted:
        g_res, g_lvl, g_all = _parse_scope(g)
        if g_res != r_res or (r_all and not g_all):
            continue
        if r_lvl in _LEVEL and g_lvl in _LEVEL and _LEVEL[g_lvl] >= _LEVEL[r_lvl]:
            return True
    return False


class Call:
    def __init__(self, tool: str, person: Person, graph: Graph, rt: Runtime, tz_name: str) -> None:
        self.tool = tool
        self.person = person
        self.graph = graph
        self.rt = rt
        self.tz_name = tz_name
        self.tz: ZoneInfo = zone(tz_name)

    @property
    def key(self) -> str:
        return self.person.key

    def need(self, *scopes: str) -> None:
        for s in scopes:
            if not scope_granted(s, self.person.scopes):
                raise M365Error("missing_permission", permission=s)

    def guard(self, **kw: Any) -> dict:
        return self.rt.guard.check(self.key, allowed_domains=self.rt.settings.allowed_recipient_domains, **kw)

    def cursor_out(self, state: dict | None) -> str | None:
        return self.rt.cursors.make(self.tool, self.key, state) if state else None

    def cursor_in(self, cursor: str | None) -> dict | None:
        return self.rt.cursors.read(self.tool, self.key, cursor) if cursor else None

    def link(self, *, op: str, kind: str, target: dict[str, Any], max_bytes: int | None = None,
             content_type: str | None = None) -> str:
        signer = self.rt.signer
        if signer is None:
            raise M365Error("links_unavailable")
        ttl = LINK_TTL_DOWNLOAD if op == "d" else LINK_TTL_UPLOAD
        try:
            value = signer.issue(op=op, kind=kind, target=target, person=self.key, token=self.person.token,
                                 token_expires=self.person.expires, ttl=ttl, max_bytes=max_bytes,
                                 content_type=content_type)
        except LinksFull:
            raise M365Error("links_full") from None
        return f"{self.rt.settings.links_base_url}/{op}/{value}"


TZ_TTL = 3600.0


async def _mailbox_zone(person: Person, graph: Graph, rt: Runtime) -> str:
    hit = rt.tz_cache.get(person.key)
    if hit and hit[1] > time.monotonic():
        return hit[0]
    name = "UTC"
    if scope_granted("MailboxSettings.Read", person.scopes):
        # A failure propagates: guessing UTC would put every offset-free time at the wrong instant.
        body = await graph.get("/me/mailboxSettings/timeZone")
        name = body.get("value") or "UTC"
    if not known(name):
        await learn_zone(graph)
    rt.tz_cache[person.key] = (name, time.monotonic() + TZ_TTL)
    return name


async def learn_zone(graph: Graph) -> None:
    """Map Outlook's Windows zone names to IANA ones by asking Graph for both lists. A zone that still
    cannot be named is shown in UTC, with the offset saying so, rather than shifted silently."""
    try:
        win = await graph.get("/me/outlook/supportedTimeZones(TimeZoneStandard=microsoft.graph.timeZoneStandard'Windows')")
        iana = await graph.get("/me/outlook/supportedTimeZones(TimeZoneStandard=microsoft.graph.timeZoneStandard'Iana')")
    except GraphError:
        return
    by_display = {z.get("displayName"): z.get("alias") for z in iana.get("value", [])}
    for z in win.get("value", []):
        alias = by_display.get(z.get("displayName"))
        if z.get("alias") and alias:
            LEARNED_ZONES[z["alias"]] = alias


def _read_only(tool: str) -> bool:
    """A failure in a tool that only reads changed nothing, whatever Graph's status says."""
    from m365_mcp.registry import all_tools

    spec = all_tools().get(tool)
    return bool(spec and spec.read_only)


@asynccontextmanager
async def call(tool: str) -> AsyncIterator[Call]:
    rt = runtime()
    person = current_person(rt.settings.expected_client_ids)
    kw: dict[str, Any] = {"transport": rt.transport}
    if rt.sleep is not None:
        kw["sleep"] = rt.sleep
    graph = Graph(person.token, person.key, **kw)
    try:
        try:
            tz_name = await _mailbox_zone(person, graph, rt)
            graph.timezone = tz_name
            yield Call(tool, person, graph, rt, tz_name)
        except GraphError as e:
            raise to_person_error(e, read_only=_read_only(tool)) from None
    finally:
        await graph.aclose()
