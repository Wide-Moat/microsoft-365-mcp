#!/usr/bin/env python3
"""Turn live-findings.json (written by `pytest -m live`) into docs/graph-live-findings.md.

    python scripts/live_findings_to_md.py [--report live-findings.json] [--out docs/graph-live-findings.md]

Every finding the suite knows about is listed; one with no entry in the report reads "not yet measured".
Rows already in the document are kept when the report lacks them, so a partial run (one probe
re-measured) does not blank the rest. With no report and no document this writes the skeleton.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

NOT_MEASURED = "not yet measured"

# name -> (what is unknown, what the code does about it today)
CATALOGUE: dict[str, tuple[str, str]] = {
    "chat_send_scope": (
        "Which of ChatMessage.Send and ChannelMessage.Send Graph enforces for a chat post",
        "send_message asks for ChatMessage.Send for chats and people, ChannelMessage.Send for channels (tools/teams.py). Which one Graph enforces cannot matter for this app: consent covers both and no token lacking one can be had."),
    "chat_soft_delete_endpoint": (
        "Whether softDelete / undoSoftDelete work on a chat message, and under which path",
        "update_message posts to /users/{me}/chats/{chat}/messages/{id}/softDelete and undoSoftDelete."),
    "chat_soft_delete_tool": (
        "What update_message(action=delete) does on a real chat",
        "Reports Graph's refusal as an error; nothing else depends on it."),
    "chat_soft_delete_restore_tool": (
        "Whether update_message(action=restore) brings the deleted message back",
        "Same tool, action restore."),
    "chat_threaded_replies": (
        "Whether chat messages have threaded replies",
        "send_message refuses reply_to outside channels and says chats have no threads."),
    "teams_html_tags": (
        "Which HTML tags Teams keeps in a message (Microsoft does not publish them)",
        "prepare_outgoing(teams=True) applies an allowlist of Teams tags (core/html.py)."),
    "search_with_filter": (
        "Whether $search combines with $filter or $orderby on /me/messages",
        "list_messages never combines them: with query, sender or recipient it sends $search alone and applies "
        "unread/flagged on the returned page; otherwise $filter with $orderby."),
    "calendarview_filter_orderby": (
        "Whether /me/calendarView accepts $filter together with $orderby",
        "list_events always sends $orderby=start/dateTime and adds a $filter only for query."),
    "pdf_conversion_small_docx": (
        "That ?format=pdf converts a small Word document",
        "download_file(format=pdf) streams Graph's conversion; a conversion refused with 4xx reaches the link as 422 JSON "
        "saying so (measured: an empty workbook), other failures as 404/502 JSON."),
    "pdf_conversion_3mb": (
        "The source size limit of ?format=pdf (the page says 2 MB)",
        "download_file(format=pdf) does not check size; a conversion Graph refuses reaches the link as 422 JSON."),
    "copy_monitor_host": (
        "The host of the copy monitor URL Graph returns",
        "Graph.monitor follows only graph.microsoft.com, *.sharepoint.com, *.onedrive.com and *.svc.ms, "
        "without a bearer token."),
    "deleted_contact_location": (
        "Where a deleted personal contact goes",
        "delete_contact needs user_confirmed and says the contact goes to Deleted Items, restorable in Outlook."),
    "deleted_todo_task": (
        "Whether a deleted To Do task can be recovered",
        "delete_task needs user_confirmed and names status=completed as the reversible alternative."),
    "deleted_onenote_page": (
        "Whether a deleted OneNote page can be recovered",
        "delete_note_page needs user_confirmed and says OneNote may not keep the page."),
    "deleted_list_item": (
        "Whether a deleted SharePoint list item goes to the site recycle bin",
        "delete_list_item needs user_confirmed and says it may not be possible to undo; the recycle bin is not read (v1.0 has none)."),
    "planner_stale_etag_status": (
        "The status Planner answers to a write with a stale etag (409 or 412)",
        "save_task and delete_task send If-Match from a fresh read; a conflict is reported as changed since read."),
    "schedule_busy_after_accept": (
        "How soon an accepted meeting shows as a busy block in getSchedule",
        "find_meeting_times (without a duration) reads getSchedule as is; nothing waits for it."),
    "people_search_lag_contact": (
        "Whether a new contact is found by /me/people within two minutes",
        "list_people(source=contacts or all) reads /me/contacts itself, so a new contact is found at once whatever the index does."),
    "recording_participant_download": (
        "Whether a participant can download a shared recording through the drive item",
        "get_event(include_recording=true) tries the meeting API and says only the organiser may download. "
        "No probe: it needs a Teams meeting recorded by someone else, which the test tenant does not have."),
    "transcripts_graph_access": (
        "Whether a tenant lets Graph read transcripts of a meeting it recorded and transcribed",
        "get_event(include_recording or include_transcript) lists recordings either way; a tenant that turns Graph "
        "access to transcripts off is named in words (GraphAccessToTranscriptsDisabled, as code or innerError)."),
    "file_search_index_lag": (
        "How long Microsoft Search takes to find a newly uploaded file",
        "list_files(query) reads search as is; where search does not answer it matches names in the one folder and says so."),
    "messages_exclude_deleted_items": (
        "Whether /me/messages honours `parentFolderId ne <Deleted Items id>` beside $orderby, and whether $search takes a $filter",
        "list_messages without a folder filters Deleted Items out of a plain listing; with query, sender or recipient the page is trimmed instead, because $search answers 400 SearchWithFilter."),
    "attachment_size_vs_real_bytes": (
        "What Graph's `size` of a mail attachment measures",
        "get_message and save_draft report it as size_incl_overhead, never as the file's size: 55 bytes answer 317, 5,000 answer 5,270, 200,000 answer 200,278 (about 260 to 280 bytes of the message's own wrapping on top, on a draft and on the received copy alike)."),
    "todo_paging": (
        "To Do's page size, and whether $top is capped",
        "list_tasks reads each To Do list with $top=100 and follows @odata.nextLink up to its page bound, saying so when cut."),
    "online_meeting_link": (
        "Whether save_event(online_meeting=true) gets a Teams join link, and when",
        "save_event waits up to 10 s for the link; a calendar that offers no Teams meetings saves a plain event, "
        "and the result says so in online_meeting_note."),
    "file_search_in_a_folder": (
        "Whether Microsoft Search answers list_files(query) for the test tenant",
        "When search answers 5xx, list_files compares the names in the one folder asked and says so in note; "
        "location=shared is refused with that reason."),
    "site_search": (
        "Whether /sites?search= finds the person's team sites",
        "list_files(kind=site) and site names also look at followed sites and each joined team's root site."),
    "list_items_expand_select": (
        "How the columns of SharePoint list items are selected inside $expand",
        "list_list_items expands fields($select=...); fields(select=...) is what it sent before, and no row could be read."),
    "onenote_paging": (
        "Whether OneNote sends a next link once $top is given, and whether it honours $skip",
        "list_notes asks for one item more than the page and continues by $skip in its own cursor."),
    "onenote_title_after_fast_patch": (
        "What a body PATCH soon after a page's creation does to its listed title",
        "Every body change sends the page's own title too, read from its content (Graph.patch_note_page)."),
    "find_meeting_times_partial": (
        "Whether findMeetingTimes offers a time when a required attendee is busy, at minimumAttendeePercentage 50 or 0",
        "find_meeting_times(required_all=false) adds the times when at least half are free from getSchedule."),
    "copy_name_clash_monitor": (
        "What the copy monitor reports when the target already holds the name",
        "update_file(copy) counts a copy its monitor reports failed as failed, with the reason for nameAlreadyExists."),
}


def cell(value: object) -> str:
    out = value if isinstance(value, str) else json.dumps(value, ensure_ascii=False, sort_keys=True)
    out = out.replace("|", "\\|").replace("\n", " ")
    return out if len(out) <= 600 else out[:597] + "..."


def existing_rows(doc: Path) -> dict[str, tuple[str, str]]:
    """name -> (observed, date) already in the document, to keep what this run did not re-measure."""
    rows: dict[str, tuple[str, str]] = {}
    if not doc.exists():
        return rows
    for line in doc.read_text().splitlines():
        m = re.match(r"^\| `([a-z0-9_]+)` \| (.*?) \| ([0-9-]*) \| .* \|$", line)
        if m and m.group(2) != NOT_MEASURED:
            rows[m.group(1)] = (m.group(2), m.group(3))
    return rows


def entry_parts(entry: object) -> tuple[object, str]:
    if isinstance(entry, dict) and "observed" in entry:
        return entry["observed"], str(entry.get("date", ""))
    return entry, ""


def render(report: dict, previous: dict[str, tuple[str, str]]) -> str:
    lines = [
        "# Graph live findings", "",
        "What the test tenant answered where Microsoft's documentation is silent. Written by",
        "`scripts/live_findings_to_md.py` from `live-findings.json`; the probes are in",
        "`tests/live/`, the procedure in [live-testing.md](live-testing.md). Do not edit by hand.", "",
        "| Finding | Observed | Date | What is unknown, and what the code does about it |", "|---|---|---|---|",
    ]
    for name, (title, code) in CATALOGUE.items():
        if name in report:
            observed, when = entry_parts(report[name])
            obs = cell(observed)
        elif name in previous:
            obs, when = previous[name]
        else:
            obs, when = NOT_MEASURED, ""
        lines.append(f"| `{name}` | {obs} | {when} | **{cell(title)}.** {cell(code)} |")
    for name in sorted(set(report) - set(CATALOGUE)):  # written by a probe the catalogue does not describe yet
        observed, when = entry_parts(report[name])
        lines.append(f"| `{name}` | {cell(observed)} | {when} | (not described: add it to CATALOGUE) |")
    return "\n".join(lines) + "\n"


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--report", default="live-findings.json")
    p.add_argument("--out", default="docs/graph-live-findings.md")
    args = p.parse_args(argv)
    report_path, out = Path(args.report), Path(args.out)
    report: dict = json.loads(report_path.read_text()) if report_path.exists() else {}
    if not report and out.exists():
        print(f"{report_path} is missing or empty; {out} left as it is", file=sys.stderr)
        return 0
    out.write_text(render(report, existing_rows(out)))
    print(f"wrote {out}: {sum(1 for n in CATALOGUE if n in report)} of {len(CATALOGUE)} findings measured in this report")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
