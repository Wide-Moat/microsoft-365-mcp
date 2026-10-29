import re

import httpx
import pytest

from m365_mcp.identity import current_person
from helpers import inner

JOIN = "https://teams.test/l/meetup-join/abc"
ME = {"emailAddress": {"name": "Me", "address": "me@x.test"}}
REC = {"id": "r1", "createdDateTime": "2026-10-23T12:00:00Z", "endDateTime": "2026-10-23T13:30:00Z"}
TRN = {"id": "t1", "createdDateTime": "2026-10-23T12:00:00Z"}


def event(**kw):
    e = {"id": "e1", "subject": "Planning <review>", "isOnlineMeeting": True, "isOrganizer": True, "organizer": ME,
         "onlineMeeting": {"joinUrl": JOIN}}
    e.update(kw)
    return e


def routes(graph, ev=None, recordings=None, transcripts=None, meeting=True):
    graph.on("GET", "/me/events/e1", ev or event())
    graph.on("GET", "/me/onlineMeetings", {"value": [{"id": "mt1", "subject": "Planning"}] if meeting else []})
    graph.on("GET", "/me/onlineMeetings/mt1/recordings", {"value": [REC] if recordings is None else recordings})
    graph.on("GET", "/me/onlineMeetings/mt1/transcripts", {"value": [TRN] if transcripts is None else transcripts})


def link_of(command):
    m = re.search(r"'(https://links\.test/d/[^']+)'", command)
    assert m, command
    return m.group(1).rsplit("/", 1)[1]


async def test_recording_and_transcript_come_back_as_signed_download_commands(graph, invoke, rt):
    routes(graph)
    r = await invoke("get_event", id="e1", include_recording=True)
    assert not r.error, r.text
    assert graph.calls("GET", "/me/onlineMeetings")[0].params["$filter"] == f"JoinWebUrl eq '{JOIN}'"
    rec, = r.data["recordings"]
    # The commands are in the structured result the agent reads, as in the text:
    # Claude Code shows the structured copy, where round 3 had left a placeholder.
    rec_cmd = rec["download"]
    trn_cmd = r.data["transcripts"][0]["download"]
    assert rec_cmd.startswith("curl ") and trn_cmd.startswith("curl ") and "text result" not in r.text
    assert rec["minutes"] == 90 and rec["started"] and rec_cmd.startswith("curl -sS --fail-with-body -o 'Planning _review_.mp4'")
    link = rt.signer.verify(link_of(rec_cmd), op="d")
    assert link.kind == "recording" and link.target == {"m": "mt1", "r": "r1", "n": "Planning _review_.mp4"}
    assert link.person == current_person().key
    trn, = r.data["transcripts"]
    tl = rt.signer.verify(link_of(trn_cmd), op="d")
    assert tl.kind == "transcript" and tl.target == {"m": "mt1", "t": "t1", "n": "Planning _review_.vtt", "a": "text/vtt"}
    assert r.data["recording_available"] is True and "400 MB per hour" in r.data["recording_note"]
    assert r.data["online_meeting"] == JOIN and inner(r.data["subject"]) == "Planning <review>"
    assert len(graph.calls("GET", "/me/events/e1")) == 1  # one read serves the event and the meeting
    assert graph.writes() == []


async def test_without_include_recording_the_meeting_is_not_looked_up(graph, invoke):
    routes(graph)
    r = await invoke("get_event", id="e1")
    assert not r.error and "recordings" not in r.data and "recording_available" not in r.data
    assert graph.calls("GET", "/me/onlineMeetings") == []


async def test_recording_and_transcript_text_together(graph, invoke):
    routes(graph)
    graph.on("GET", "/me/onlineMeetings/mt1/transcripts/t1/content", b"WEBVTT\n\n00:00.000 --> 00:01.000\n<v Anna>Hi</v>\n")
    r = await invoke("get_event", id="e1", include_recording=True, include_transcript=True)
    assert not r.error, r.text
    assert r.data["recordings"] and r.data["transcripts"] and "Anna: Hi" in r.data["transcript"]


async def test_an_event_that_is_not_a_teams_meeting_is_refused(graph, invoke):
    graph.on("GET", "/me/events/e1", {"id": "e1", "subject": "Lunch", "isOnlineMeeting": False})
    r = await invoke("get_event", id="e1", include_recording=True)
    assert r.error and "not a Teams meeting" in r.text
    assert graph.calls("GET", "/me/onlineMeetings") == []


async def test_no_recording_yet_is_a_sentence_not_an_error(graph, invoke):
    routes(graph, recordings=[], transcripts=[])
    r = await invoke("get_event", id="e1", include_recording=True)
    assert not r.error, r.text
    assert r.data["recording_available"] is False and "no recording or transcript yet" in r.data["recording_note"]
    assert "not found" not in r.text.lower()


async def test_only_a_transcript_is_still_available(graph, invoke):
    routes(graph, recordings=[])
    r = await invoke("get_event", id="e1", include_recording=True)
    assert r.data["recording_available"] is True and "recordings" not in r.data and len(r.data["transcripts"]) == 1


async def test_a_participant_without_download_rights_is_told_and_given_the_meeting_link(graph, invoke):
    routes(graph, ev=event(isOrganizer=False))
    graph.on("GET", "/me/onlineMeetings/mt1/recordings", {"error": {"code": "Forbidden", "message": "denied"}}, status=403)
    r = await invoke("get_event", id="e1", include_recording=True)
    assert not r.error, r.text
    assert "only by the meeting's organiser" in r.data["recording_note"] and r.data["online_meeting"] == JOIN
    assert "recordings" not in r.data and len(r.data["transcripts"]) == 1


async def test_a_participant_who_is_listed_a_recording_is_warned_about_download(graph, invoke):
    routes(graph, ev=event(isOrganizer=False))
    r = await invoke("get_event", id="e1", include_recording=True)
    assert r.data["recordings"] and "Only the organiser can download" in r.data["recording_note"]


async def test_an_expired_recording_says_it_expired(graph, invoke):
    routes(graph)
    graph.on("GET", "/me/onlineMeetings/mt1/recordings", {"error": {"code": "Gone", "message": "x"}}, status=410)
    r = await invoke("get_event", id="e1", include_recording=True)
    assert not r.error and "has expired" in r.data["recording_note"] and "recordings" not in r.data


async def test_a_missing_permission_is_named(graph, invoke):
    routes(graph)
    graph.on("GET", "/me/onlineMeetings/mt1/recordings", {"error": {
        "code": "Forbidden", "message": "Missing scope OnlineMeetingRecording.Read.All"}}, status=403)
    r = await invoke("get_event", id="e1", include_recording=True)
    assert r.error and "OnlineMeetingRecording.Read.All" in r.text and "administrator" in r.text


async def test_a_meeting_teams_does_not_know_is_reported(graph, invoke):
    routes(graph, meeting=False)
    r = await invoke("get_event", id="e1", include_recording=True)
    assert not r.error and r.data["recording_available"] is False and "no record" in r.data["recording_note"]


async def test_transcripts_turned_off_by_policy_do_not_hide_the_recording(graph, invoke):
    routes(graph)
    graph.on("GET", "/me/onlineMeetings/mt1/transcripts",
             {"error": {"code": "GraphAccessToTranscriptsDisabled", "message": "x"}}, status=403)
    r = await invoke("get_event", id="e1", include_recording=True)
    assert not r.error and r.data["recordings"] and "turned off access to meeting transcripts" in r.data["recording_note"]


async def test_transcripts_turned_off_as_the_live_tenant_says_it_do_not_hide_the_recording(graph, invoke):
    # The shape a Business Basic tenant returned on 2026-10-23: the reason is in innerError, not code.
    routes(graph)
    graph.on("GET", "/me/onlineMeetings/mt1/transcripts",
             {"error": {"code": "Forbidden", "message": "Graph API access to transcripts is disabled for this tenant.",
                        "innerError": {"code": "GraphAccessToTranscriptsDisabled"}}}, status=403)
    r = await invoke("get_event", id="e1", include_recording=True)
    assert not r.error and r.data["recordings"]
    assert "turned off access to meeting transcripts" in r.data["recording_note"]
    assert "lack access" not in r.data["recording_note"]
    assert graph.writes() == []


async def test_without_link_signing_no_command_is_made(graph, invoke, rt):
    routes(graph)
    rt.signer = None
    r = await invoke("get_event", id="e1", include_recording=True)
    assert r.error and "File transfer is not configured" in r.text


async def test_a_refused_transcript_list_is_not_the_generic_refusal_beside_a_recording(graph, invoke):
    # Round 3 (C12): recording_available was true while the note said "does not let you do this here",
    # which reads as the recording being refused. The note now says which part was refused.
    routes(graph)
    graph.on("GET", "/me/onlineMeetings/mt1/transcripts", {"error": {"code": "Forbidden"}}, status=403)
    r = await invoke("get_event", id="e1", include_recording=True)
    note = r.data["recording_note"]
    assert r.data["recording_available"] is True and r.data["recordings"]
    assert "does not let you" not in note and "lack access" not in note
    assert "transcript" in note and "recording is not affected" in note


async def test_a_refused_transcript_read_says_it_is_the_transcript(graph, invoke):
    routes(graph)
    graph.on("GET", "/me/onlineMeetings/mt1/transcripts", {"error": {"code": "Forbidden"}}, status=403)
    r = await invoke("get_event", id="e1", include_transcript=True)
    assert "does not let you" not in r.data["transcript_note"] and "transcript" in r.data["transcript_note"]


async def test_any_transcript_refusal_beside_a_recording_says_the_recording_is_fine(graph, invoke):
    routes(graph)
    graph.on("GET", "/me/onlineMeetings/mt1/transcripts",
             {"error": {"code": "GraphAccessToTranscriptsDisabled", "message": "x"}}, status=403)
    r = await invoke("get_event", id="e1", include_recording=True)
    assert "The recording is not affected." in r.data["recording_note"]
