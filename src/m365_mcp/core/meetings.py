"""Teams meetings behind calendar events: finding the meeting, and reading a transcript as text."""

from __future__ import annotations

import re
from typing import Any

from m365_mcp.core.context import Call
from m365_mcp.graph.errors import GraphError

TEXT_TRANSCRIPT = "application/vnd.microsoft.graph.transcript+text"
_VOICE = re.compile(r"<v\s+([^>]+)>(.*?)(?:</v>)?$")
_CUE_ID = re.compile(r"[0-9a-fA-F-]{8,}/\d+-\d+")
_TAG = re.compile(r"</?[a-z][^>]*>")


def event_join_url(event: dict[str, Any]) -> str | None:
    """The Teams join link of a calendar event, or None when it is not an online meeting."""
    om = event.get("onlineMeeting") or {}
    return om.get("joinUrl") or (event.get("onlineMeetingUrl") if event.get("isOnlineMeeting") else None) or None


async def meeting_for(c: Call, join_url: str) -> dict[str, Any] | None:
    """The online meeting that a join link belongs to; None when Graph has no record of it."""
    url = join_url.replace("'", "''")
    body = await c.graph.get("/me/onlineMeetings", params={"$filter": f"JoinWebUrl eq '{url}'"})
    found = body.get("value") or []
    return found[0] if found else None


def vtt_to_text(vtt: str) -> str:
    """Speaker-attributed VTT reduced to `speaker: text` lines; consecutive lines of one speaker merge."""
    lines: list[tuple[str | None, str]] = []
    for raw in vtt.replace("\r", "").split("\n"):
        line = raw.strip()
        if not line or line.startswith(("WEBVTT", "NOTE", "STYLE")) or "-->" in line:
            continue
        if _CUE_ID.fullmatch(line):
            continue
        m = _VOICE.match(line)
        speaker, text = (m.group(1).strip(), m.group(2)) if m else (None, line)
        text = _TAG.sub("", text).strip()
        if not text:
            continue
        if lines and lines[-1][0] == speaker:
            lines[-1] = (speaker, lines[-1][1] + " " + text)
        else:
            lines.append((speaker, text))
    return "\n".join(f"{s}: {t}" if s else t for s, t in lines)


async def transcript_text(c: Call, meeting_id: str, transcript_id: str) -> tuple[str, bool]:
    """The transcript as text and whether speakers are named. Where the organisation forbids speaker
    attribution the plain-text form is read instead."""
    path = f"/me/onlineMeetings/{meeting_id}/transcripts/{transcript_id}/content"
    try:
        resp = await c.graph.request("GET", path, headers={"Accept": "text/vtt"})
        return vtt_to_text(resp.text), True
    except GraphError as e:
        if e.code != "SpeakerAttributionNotAllowed":
            raise
    resp = await c.graph.request("GET", path, headers={"Accept": TEXT_TRANSCRIPT})
    return vtt_to_text(resp.text), False
