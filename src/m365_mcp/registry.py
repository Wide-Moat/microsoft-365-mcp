"""Tools declare their area and behaviour once; the server publishes only enabled areas.

The behaviour flags are the source of the MCP annotations, and a contract test checks the
two never disagree with the code: a tool that notifies other people must call the guard,
a read-only tool must not write.

A parameter can need an area or a switch of its own (`gates`): one question stays one tool, and
the depth that needs another permission disappears from the schema while it is off.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any

AREAS: dict[str, bool] = {
    # area -> on by default
    "account": True,
    "mail": True,
    "calendar": True,
    "contacts": True,
    "files": True,
    "lists": True,
    "excel": True,
    "teams": True,
    "recordings": True,
    "tasks": True,
    "notes": True,
    "bookings": False,
}

ALWAYS_ON = frozenset({"account"})

ALLOWED_VERBS = (
    "list_", "get_", "save_", "delete_", "send_", "cancel_", "respond_", "share_", "update_",
    "download_", "upload_", "find_",
)


@dataclass(frozen=True)
class ToolSpec:
    fn: Callable[..., Any]
    name: str
    title: str
    area: str
    scopes: tuple[str, ...]
    read_only: bool
    notifies_others: bool
    irreversible: bool
    idempotent: bool
    returns_list: bool
    switch: str | None = None  # an extra switch inside an area, e.g. "teams_channels"
    description: str = field(default="")
    # parameter -> ("area" | "switch", name): the parameter exists only while that is on
    gates: dict[str, tuple[str, str]] = field(default_factory=dict)


_TOOLS: dict[str, ToolSpec] = {}


def tool(
    *,
    area: str,
    title: str,
    scopes: list[str],
    read_only: bool = False,
    notifies_others: bool = False,
    irreversible: bool = False,
    idempotent: bool = False,
    returns_list: bool = False,
    switch: str | None = None,
    gates: dict[str, tuple[str, str]] | None = None,
) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    if area not in AREAS:
        raise ValueError(f"unknown area {area!r}")
    for kind, name in (gates or {}).values():
        if kind not in ("area", "switch") or (kind == "area" and name not in AREAS):
            raise ValueError(f"unknown gate {kind} {name!r}")

    def register(fn: Callable[..., Any]) -> Callable[..., Any]:
        name = fn.__name__
        if name in _TOOLS:
            raise ValueError(f"tool {name!r} declared twice")
        doc = (fn.__doc__ or "").strip()
        _TOOLS[name] = ToolSpec(
            fn=fn, name=name, title=title, area=area, scopes=tuple(scopes), read_only=read_only,
            notifies_others=notifies_others, irreversible=irreversible, idempotent=idempotent,
            returns_list=returns_list, switch=switch, description=" ".join(doc.split()), gates=dict(gates or {}),
        )
        return fn

    return register


def all_tools() -> dict[str, ToolSpec]:
    # Importing the area modules is what registers their tools.
    from m365_mcp import tools  # noqa: F401

    return dict(_TOOLS)


def enabled(areas: frozenset[str], switches: frozenset[str] = frozenset()) -> list[ToolSpec]:
    on = set(areas) | ALWAYS_ON
    return [
        s for s in all_tools().values()
        if s.area in on and (s.switch is None or s.switch in switches)
    ]
