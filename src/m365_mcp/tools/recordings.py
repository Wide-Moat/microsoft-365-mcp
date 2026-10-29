"""Recordings and transcripts of Teams meetings, as download links the agent's sandbox can use.

The recordings area has no tool of its own: it is the `include_recording` depth of get_event.
"""

from __future__ import annotations

import re
from datetime import datetime
from typing import Any

from m365_mcp.core.context import Call
from m365_mcp.core.errors import invalid
from m365_mcp.core.meetings import event_join_url, meeting_for
from m365_mcp.core.project import compact
from m365_mcp.core.time import from_graph
from m365_mcp.graph.errors import GraphError, to_person_error

SIZE_NOTE = "A recording can reach about 400 MB per hour; the link streams and supports Range."


def _file_name(subject: str, ext: str) -> str:
    base = re.sub(r"[^\w.\- ]", "_", subject).strip() or "meeting"
    return f"{base[:80]}.{ext}"


def _curl(name: str, url: str) -> str:
    return f"curl -sS --fail-with-body -o '{name.replace(chr(39), '')}' '{url}'"


def _minutes(item: dict[str, Any]) -> int | None:
    a, b = item.get("createdDateTime"), item.get("endDateTime")
    if not (a and b):
        return None
    try:
        delta = datetime.fromisoformat(b.replace("Z", "+00:00")) - datetime.fromisoformat(a.replace("Z", "+00:00"))
    except ValueError:
        return None
    return max(int(delta.total_seconds() // 60), 0) or None


def transcript_refusal(e: GraphError, *, beside_recording: bool = False) -> str:
    """Why transcripts could not be listed. The generic "does not let you do this here" read as if the
    recording itself were refused (round 3), so a bare refusal is named for what it was."""
    err = to_person_error(e)
    text = err.text
    if err.key == "forbidden":
        text = ("Microsoft 365 refused to list this meeting's transcripts: transcription may not be enabled for it, "
                "or only the organiser may read them.")
    # Beside a recording, any transcript refusal read as if the recording were refused too.
    return text + " The recording is not affected." if beside_recording else text


def _expired(e: GraphError) -> bool:
    return e.status == 410 or "expire" in (e.code or "").lower()


async def recording_view(c: Call, e: dict[str, Any]) -> dict[str, Any]:
    """The recordings and transcripts of the Teams meeting behind one calendar event, as curl commands
    that download them into the sandbox; the video never enters the chat. This is the `recordings`
    area: get_event calls it for include_recording, which exists only while the area is on."""
    c.need("OnlineMeetings.Read")
    join_url = event_join_url(e)
    if not join_url:
        raise invalid("That event is not a Teams meeting, so it has no recording or transcript.")
    subject, organizer_is_me = e.get("subject") or "", e.get("isOrganizer") is True
    meeting = await meeting_for(c, join_url)
    if not meeting:
        return {"recording_available": False, "recording_note": "Teams has no record of this meeting, so there is no "
                "recording. Recordings exist only for meetings created from a calendar invitation."}
    mid = meeting["id"]
    subject = subject or meeting.get("subject") or ""
    out: dict[str, Any] = {}
    notes: list[str] = []
    recordings, transcripts = await _recordings(c, mid, subject, out, notes), await _transcripts(c, mid, subject, notes)
    if recordings:
        out["recordings"] = recordings
        notes.append(SIZE_NOTE)
        if not organizer_is_me:
            notes.append("Only the organiser can download recordings; if a link fails, watch the meeting "
                         "recording in Teams from the meeting link.")
    if transcripts:
        out["transcripts"] = transcripts
    out["recording_available"] = bool(recordings or transcripts)
    if not out["recording_available"] and not notes:
        notes.append("There is no recording or transcript yet. They appear some minutes after the meeting "
                     "ends, and only if it was recorded or transcribed.")
    if notes:
        out["recording_note"] = " ".join(notes)
    return out


async def _recordings(c: Call, mid: str, subject: str, out: dict[str, Any], notes: list[str]) -> list[dict[str, Any]]:
    c.need("OnlineMeetingRecording.Read.All")
    try:
        items = (await c.graph.get(f"/me/onlineMeetings/{mid}/recordings")).get("value") or []
    except GraphError as e:
        if e.status == 403 and not e.missing_scope:
            notes.append("This recording can be downloaded only by the meeting's organiser. Watch it in Teams "
                         "from the meeting link.")
            return []
        if _expired(e):
            notes.append("The recording has expired and was removed under the organisation's retention policy.")
            return []
        raise
    found = []
    for r in items:
        name = _file_name(subject, "mp4")
        url = c.link(op="d", kind="recording", target={"m": mid, "r": r["id"], "n": name})
        found.append(compact({"started": from_graph(r.get("createdDateTime"), c.tz), "minutes": _minutes(r),
                              "download": _curl(name, url)}))
    return found


async def _transcripts(c: Call, mid: str, subject: str, notes: list[str]) -> list[dict[str, Any]]:
    c.need("OnlineMeetingTranscript.Read.All")
    try:
        items = (await c.graph.get(f"/me/onlineMeetings/{mid}/transcripts")).get("value") or []
    except GraphError as e:
        if e.status == 401 or (e.status == 403 and e.missing_scope):
            raise
        if _expired(e):
            notes.append("The transcript has expired and was removed under the organisation's retention policy.")
        else:
            notes.append(transcript_refusal(e, beside_recording=True))
        return []
    found = []
    for t in items:
        name = _file_name(subject, "vtt")
        url = c.link(op="d", kind="transcript", target={"m": mid, "t": t["id"], "n": name, "a": "text/vtt"})
        found.append(compact({"started": from_graph(t.get("createdDateTime"), c.tz), "download": _curl(name, url)}))
    return found
