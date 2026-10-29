"""Shared pieces of projections: people as one string, recipients lists, compact dicts."""

from __future__ import annotations

from typing import Any

from m365_mcp.core.errors import invalid


def who(addr_obj: dict[str, Any] | None) -> str | None:
    """Graph `recipient`/`emailAddress` → "Name <address>"."""
    if not addr_obj:
        return None
    ea = addr_obj.get("emailAddress", addr_obj)
    name, address = (ea.get("name") or "").strip(), (ea.get("address") or "").strip()
    if name and address and name.casefold() != address.casefold():
        return f"{name} <{address}>"
    return address or name or None


def whos(items: list[dict[str, Any]] | None, limit: int | None = None) -> list[str]:
    out = [w for w in (who(i) for i in (items or [])) if w]
    return out[:limit] if limit else out


def compact(d: dict[str, Any]) -> dict[str, Any]:
    """Drop empty values and false booleans: absent means no."""
    return {k: v for k, v in d.items() if not (v is None or v is False or (isinstance(v, (str, list, dict)) and not v))}


def page_result(items: list[dict[str, Any]], cursor: str | None, **extra: Any) -> dict[str, Any]:
    out: dict[str, Any] = {"items": items, "count": len(items), "has_more": bool(cursor)}
    if cursor:
        out["cursor"] = cursor
    out.update({k: v for k, v in extra.items() if v is not None})
    if cursor and out.get("limit_clamped_to"):
        # Said only where there is a cursor to follow: a tool that does not page must not promise one.
        out["limit_note"] = f"At most {out['limit_clamped_to']} come at a time; cursor gives the next page."
    return out


def clamp_limit(limit: int, default: int = 10, cap: int = 25) -> tuple[int, dict[str, Any]]:
    if limit is None:
        return default, {}
    if limit < 1:
        # Below the range is a mistake to name, not a value to correct silently (above it is clamped and said).
        raise invalid(f"limit is 1 to {cap}.")
    if limit > cap:
        return cap, {"limit_clamped_to": cap}
    return limit, {}
