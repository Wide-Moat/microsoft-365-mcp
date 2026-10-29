"""Times as people say them, in the person's mailbox time zone."""

from __future__ import annotations

import re
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from m365_mcp.core.errors import invalid

# Outlook reports Windows zone names; the common ones mapped to IANA. Unknown → UTC, said so.
WINDOWS_ZONES = {
    "UTC": "UTC", "GMT Standard Time": "Europe/London", "W. Europe Standard Time": "Europe/Berlin",
    "Romance Standard Time": "Europe/Paris", "Central Europe Standard Time": "Europe/Budapest",
    "E. Europe Standard Time": "Europe/Chisinau", "FLE Standard Time": "Europe/Kiev",
    "Russian Standard Time": "Europe/Moscow", "Eastern Standard Time": "America/New_York",
    "Central Standard Time": "America/Chicago", "Mountain Standard Time": "America/Denver",
    "Pacific Standard Time": "America/Los_Angeles", "GTB Standard Time": "Europe/Bucharest",
    "Greenwich Standard Time": "Atlantic/Reykjavik", "India Standard Time": "Asia/Kolkata",
    "China Standard Time": "Asia/Shanghai", "Tokyo Standard Time": "Asia/Tokyo",
    "AUS Eastern Standard Time": "Australia/Sydney", "Arabian Standard Time": "Asia/Dubai",
    "Turkey Standard Time": "Europe/Istanbul", "Israel Standard Time": "Asia/Jerusalem",
    "E. South America Standard Time": "America/Sao_Paulo", "SA Pacific Standard Time": "America/Bogota",
}

_LOCAL = re.compile(r"^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?\s*(Z|[+-]\d{2}:?\d{2})?$")


# Learned from Graph at run time (supportedTimeZones, Windows and IANA lists joined by display name).
LEARNED_ZONES: dict[str, str] = {}


def known(name: str | None) -> bool:
    if not name:
        return False
    iana = LEARNED_ZONES.get(name) or WINDOWS_ZONES.get(name, name)
    try:
        ZoneInfo(iana)
        return True
    except (ZoneInfoNotFoundError, ValueError):
        return False


def zone(name: str | None) -> ZoneInfo:
    if not name:
        return ZoneInfo("UTC")
    iana = LEARNED_ZONES.get(name) or WINDOWS_ZONES.get(name, name)
    try:
        return ZoneInfo(iana)
    except (ZoneInfoNotFoundError, ValueError):
        return ZoneInfo("UTC")


def parse(value: str, tz: ZoneInfo, *, field: str = "time") -> datetime:
    m = _LOCAL.match(value.strip())
    if not m:
        raise invalid(f"{field} must look like 2026-10-23 15:00 (local time) or carry an offset like +01:00.")
    y, mo, d, hh, mm, ss, off = m.groups()
    try:
        dt = datetime(int(y), int(mo), int(d), int(hh or 0), int(mm or 0), int(ss or 0))
    except ValueError:
        raise invalid(f"{field} is not a real date or time.") from None
    if off:
        return datetime.fromisoformat(dt.isoformat() + ("+00:00" if off == "Z" else off if ":" in off else off[:3] + ":" + off[3:]))
    return dt.replace(tzinfo=tz)


def graph_dt(dt: datetime) -> dict:
    """Graph's dateTimeTimeZone, always sent in UTC so nothing depends on Windows zone names."""
    utc = dt.astimezone(ZoneInfo("UTC"))
    return {"dateTime": utc.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": "UTC"}


def from_graph(value: dict | str | None, tz: ZoneInfo) -> str | None:
    """Graph time (dateTimeTimeZone or ISO string) → local string with offset."""
    if not value:
        return None
    if isinstance(value, dict):
        raw, src = value.get("dateTime"), value.get("timeZone") or "UTC"
        if not raw:
            return None
        base = datetime.fromisoformat(raw[:19])
        dt = base.replace(tzinfo=zone(src))
    else:
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=ZoneInfo("UTC"))
    return dt.astimezone(tz).isoformat(timespec="minutes")


def window(start: str, end: str, tz: ZoneInfo, *, max_days: int) -> tuple[datetime, datetime]:
    s, e = parse(start, tz, field="start"), parse(end, tz, field="end")
    if e <= s:
        raise invalid("end must be after start.")
    if e - s > timedelta(days=max_days):
        raise invalid(f"The window is at most {max_days} days. Narrow it or make several calls.")
    return s, e
