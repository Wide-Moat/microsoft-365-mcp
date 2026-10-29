"""Scenarios on a real Microsoft 365 test tenant, two people, real effects.

Every name carries the run id (`m365-live-<8 hex>`), and the session finalizer removes what the run
made. Some scenarios are chains (the reply needs the first letter), so a test whose predecessor did
not run says so and skips; the predecessor's failure is the one to read.
"""

from __future__ import annotations

import asyncio
import base64
import hashlib
import json
import os
import time
import uuid
from datetime import datetime, timedelta

import pytest

import builders
from m365_mcp.core.time import zone

pytestmark = pytest.mark.live


def unfenced_rows(block: str) -> list:
    """The rows inside a get_range content block: one JSON array per line between the markers."""
    lines = block.split("\n")
    assert lines[0].startswith("<<<content ") and lines[-1].startswith("<<<end of content "), block
    return [json.loads(line) for line in lines[1:-1]]

MB = 1024 * 1024
ARRIVE = 60  # seconds a sent letter may take to show up in a search


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def need(live, key: str, because: str):
    if key not in live.state:
        pytest.skip(f"depends on {because}, which did not complete")
    return live.state[key]


async def local_zone(live, user: str):
    return zone((await live.ok(user, "get_connection"))["time_zone"])


def next_working_days(tz, count: int = 5):
    day, days = datetime.now(tz).date() + timedelta(days=1), []
    while len(days) < count:
        if day.weekday() < 5:
            days.append(day)
        day += timedelta(days=1)
    return days


async def find_message(live, user: str, needle: str, *, subject_starts: str = "", timeout: float = ARRIVE):
    async def check():
        data = await live.ok(user, "list_messages", query=needle, limit=25)
        hits = [m for m in data["items"] if m["subject"].startswith(subject_starts)]
        return hits[0] if hits else None
    return await live.poll(check, timeout=timeout, what=f"a letter matching {needle!r} for user {user}")


async def full_html(live, user: str, message_id: str) -> str:
    r = await live.raw(user, "GET", f"/me/messages/{message_id}", params={"$select": "body"},
                       prefer='outlook.body-content-type="html"')
    assert r.status_code == 200, live.error_of(r)
    return r.json()["body"]["content"]


# -- mail ---------------------------------------------------------------------------------------

async def test_mail_a_to_b_with_table_inline_image_and_10mb_attachment(live):
    b = await live.address("b")
    marker = f"ORIGINALMARK{uuid.uuid4().hex[:10]}"
    subject = f"{live.run} letter"
    pixel = base64.b64encode(builders.png()).decode()
    html = (f"<p>{marker}</p>"
            '<table style="border:1px solid #888;border-collapse:collapse"><tr>'
            '<th style="background:#eef;padding:6px">Item</th><th style="background:#eef;padding:6px">Qty</th></tr>'
            '<tr><td style="padding:6px">Widgets</td><td style="padding:6px">3</td></tr></table>'
            f'<p><img src="data:image/png;base64,{pixel}" alt="a blue square" width="24" height="24"></p>')

    draft = await live.ok("a", "save_draft", to=[b], subject=subject, body=html, body_format="html",
                          attachments_upload=1)
    assert "data:image" not in draft["preview"]
    assert any(a.get("inline") and a["name"].endswith(".png") for a in draft["attachments"]), draft["attachments"]

    name, payload = f"{live.run}-10mb.bin", builders.random_bytes(10 * MB)
    put = await live.upload(live.link_in(draft["attachment_uploads"][0]), payload,
                            {"Content-Type": "application/octet-stream", "X-File-Name": name})
    assert put.status_code == 200, put.text[:300]
    assert put.json()["size"] == 10 * MB

    sent = await live.ok("a", "send_draft", id=draft["draft_id"], user_confirmed=True)
    assert sent["sent"] is True

    arrived = await find_message(live, "b", live.run, subject_starts=subject)
    got = await live.ok("b", "get_message", id=arrived["id"], include_attachments=True)
    assert marker in got["body"]
    big = next((a for a in got["attachments"] if a["name"] == name), None)
    assert big, got["attachments"]
    res = await live.fetch(live.link_in(big["download"]))
    assert res.status_code == 200, res.text[:300]
    assert len(res.content) == 10 * MB and sha(res.content) == sha(payload)
    inline = [a for a in got["attachments"] if a.get("inline")]
    assert inline and inline[0]["type"] == "image/png", got["attachments"]
    assert marker in await full_html(live, "b", arrived["id"])

    live.state["marker"] = marker
    live.state["b_message"] = arrived["id"]
    live.state["subject"] = subject


async def test_mail_reply_keeps_the_quoted_letter(live):
    source = need(live, "b_message", "the first letter")
    marker = live.state["marker"]
    text = f"REPLYMARK{uuid.uuid4().hex[:10]}"

    draft = await live.ok("b", "save_draft", mode="reply", source_id=source, body=text)
    html = await full_html(live, "b", draft["draft_id"])
    assert text in html and marker in html, "the draft lost either the new text or the quoted letter"
    assert html.index(text) < html.index(marker), "the new text must sit above the quoted history"

    await live.ok("b", "send_draft", id=draft["draft_id"], user_confirmed=True)
    reply = await find_message(live, "a", live.run, subject_starts="RE:")
    assert text in (await live.ok("a", "get_message", id=reply["id"]))["body"]
    assert marker in await full_html(live, "a", reply["id"]), "the original must be present in the full body"


async def test_send_without_confirmation_sends_nothing(live):
    b = await live.address("b")
    token = f"REFUSEDMARK{uuid.uuid4().hex[:10]}"
    draft = await live.ok("a", "save_draft", to=[b], subject=f"{live.run} {token}", body="must not arrive")

    for args in ({}, {"user_confirmed": False}):
        error, text, _ = await live.invoke("a", "send_draft", id=draft["draft_id"], **args)
        assert error and "user_confirmed" in text, text

    still = await live.raw("a", "GET", f"/me/messages/{draft['draft_id']}", params={"$select": "isDraft"})
    assert still.json()["isDraft"] is True
    for _ in range(4):  # 20 s in all
        await asyncio.sleep(5)
        seen = await live.ok("b", "list_messages", query=token, limit=5)
        assert seen["items"] == [], "a letter arrived although the send was refused"


# -- calendar -------------------------------------------------------------------------------------

async def test_calendar_invite_answer_and_cancel(live):
    b = await live.address("b")
    tz = await local_zone(live, "a")
    days = next_working_days(tz)

    found = await live.ok("a", "find_meeting_times", people=[b], duration_minutes=30,
                          start=f"{days[0]} 00:00", end=f"{days[-1] + timedelta(days=1)} 00:00")
    assert found["count"] >= 1, found.get("note")
    slot = found["items"][0]["start"]
    # Working hours: the person's, or the documented default; never a night or a weekend slot.
    hours = found["working_hours"]
    for item in found["items"]:
        at = datetime.fromisoformat(item["start"]).astimezone(tz)
        assert at.weekday() < 5 or hours["source"] == "Outlook", item
        assert hours["start"] <= at.strftime("%H:%M") < hours["end"], (item, hours)

    subject = f"{live.run} meeting"
    made = await live.ok("a", "save_event", subject=subject, start=slot, duration_minutes=30, attendees=[b],
                         online_meeting=True, user_confirmed=True)
    assert made["saved"] == "created" and made["attendee_count"] == 1
    event = made["id"]

    # Measured: a calendar that offers no Teams meetings saves a plain event; the result must say so
    # rather than look like a meeting whose link is late.
    note = made.get("online_meeting_note", "")
    live.finding("online_meeting_link", made.get("online_meeting") and "link in the save result" or note)
    if "plain event" not in note:
        async def joinable():
            return (await live.ok("a", "get_event", id=event)).get("online_meeting")
        assert made.get("online_meeting") or await live.poll(joinable, timeout=90, what="the Teams join link of the event")

    slot_day = datetime.fromisoformat(slot).date()
    window = {"start": f"{slot_day - timedelta(days=1)} 00:00", "end": f"{slot_day + timedelta(days=2)} 00:00"}

    async def invited():
        data = await live.ok("b", "list_events", query=live.run, **window)
        return data["items"][0] if data["items"] else None
    theirs = await live.poll(invited, timeout=ARRIVE, what="the invitation in B's calendar")
    assert theirs["subject"] == subject and theirs["my_response"] == "not answered"

    quiet = await live.ok("b", "respond_to_event", id=theirs["id"], response="tentative", send_response=False)
    assert quiet["answered"] == "tentative" and quiet["organizer_told"] is False
    await live.ok("b", "respond_to_event", id=theirs["id"], response="accept")

    async def accepted():
        data = await live.ok("a", "get_event", id=event, include_attendees=True)
        return any(a.get("response") == "accepted" and b.casefold() in a["who"].casefold()
                   for a in data.get("attendees", []))
    await live.poll(accepted, timeout=ARRIVE, what="B's acceptance on the organiser's copy")

    schedule = await live.ok("a", "find_meeting_times", people=[b], start=f"{slot_day} 00:00",
                             end=f"{slot_day + timedelta(days=1)} 00:00")
    row = schedule["items"][0]
    assert set(row["availability"]) <= set("01234") and row["availability"], row
    assert row.get("working_hours") and isinstance(row.get("busy", []), list)
    started = time.monotonic()

    async def busy():
        r = (await live.ok("a", "find_meeting_times", people=[b], start=f"{slot_day} 00:00",
                           end=f"{slot_day + timedelta(days=1)} 00:00"))["items"][0]
        return any(blk["start"] == slot for blk in r.get("busy", []))
    try:
        await live.poll(busy, timeout=60, what="the accepted meeting as a busy block")
        live.finding("schedule_busy_after_accept", f"busy block visible {int(time.monotonic() - started)} s after accept")
    except AssertionError:
        live.finding("schedule_busy_after_accept", "no busy block for the accepted meeting within 60 s")

    cancelled = await live.ok("a", "cancel_event", id=event, comment="live test", user_confirmed=True)
    assert cancelled["cancelled"] is True and cancelled["notified"] == 1

    async def gone():
        data = await live.ok("b", "list_events", query=live.run, **window)
        return all(e.get("cancelled") for e in data["items"])
    await live.poll(gone, timeout=ARRIVE, what="the meeting cancelled or gone for B")


# -- files ----------------------------------------------------------------------------------------

async def test_files_upload_list_download_share_delete(live):
    folder = await live.files_folder()
    name, payload = f"{live.run}-6mb.bin", builders.random_bytes(6 * MB)  # over 4 MB: an upload session
    item = await live.put_file(folder, name, payload)

    listing = await live.ok("a", "list_files", folder=folder)
    assert name in [i["name"] for i in listing["items"]]
    # Through Microsoft Search, which indexes a new file some time after the upload, or, where search
    # does not answer, by name in the folder at once (and said so). A text file must turn up; whether
    # the binary one does by then is recorded, not asserted (measured: search left a .bin out for 10 min).
    text_name = f"{live.run}-searchable.txt"
    await live.put_file(folder, text_name, b"searchable")
    started, last = time.monotonic(), {}

    async def found():
        last.update(await live.ok("a", "list_files", folder=folder, query=live.run))
        return [i for i in last["items"] if i["name"] == text_name]
    await live.poll(found, timeout=600, every=15, what="the new text file in list_files(query)")
    names = [i["name"] for i in last["items"]]
    live.finding("file_search_index_lag", {
        "new .txt found after s": round(time.monotonic() - started),
        "6 MB .bin found by then": name in names,
    })
    live.finding("file_search_in_a_folder", last.get("note") or "Microsoft Search answered")

    plan = await live.ok("a", "download_file", id=item["id"])
    res = await live.fetch(live.link_in(plan["download"]))
    assert res.status_code == 200 and sha(res.content) == sha(payload)

    error, text, _ = await live.invoke("a", "share_file", id=item["id"], action="link")
    assert error and "user_confirmed" in text
    shared = await live.ok("a", "share_file", id=item["id"], action="link", user_confirmed=True)
    assert shared["reach"] == "organization" and shared["link"].startswith("https://"), shared

    text = (f"{live.run} text line\n" * 4000).encode()  # text is what OneDrive gzips (measured)
    small = await live.put_file(folder, f"{live.run}.txt", text)
    got = await live.fetch(live.link_in((await live.ok("a", "download_file", id=small["id"]))["download"]))
    assert got.status_code == 200 and "content-encoding" not in got.headers and got.content == text
    assert int(got.headers.get("content-length", len(text))) == len(text)

    deleted = await live.ok("a", "update_file", action="delete", ids=[item["id"]])
    assert deleted["done"] == 1 and "recycle bin" in deleted["note"]
    after = await live.ok("a", "list_files", folder=folder)
    assert name not in [i["name"] for i in after["items"]]


async def test_files_pdf_conversion_of_a_small_document(live):
    folder = await live.files_folder()
    item = await live.put_file(folder, f"{live.run}-small.docx", builders.docx(f"{live.run} conversion"))
    plan = await live.ok("a", "download_file", id=item["id"], format="pdf")
    res = await live.fetch(live.link_in(plan["download"]))
    if res.status_code != 200:
        live.finding("pdf_conversion_small_docx", f"refused: {res.status_code} {res.text[:200]}")
        pytest.skip(f"Graph refused the PDF conversion of a small docx: {res.status_code}")
    assert res.content.startswith(b"%PDF")
    live.finding("pdf_conversion_small_docx", f"converted, {len(res.content)} bytes of PDF")


async def test_the_persons_team_sites_are_listed_and_found_by_name(live):
    sites = (await live.ok("a", "list_files", kind="site", limit=25))["items"]
    if not sites:
        pytest.skip("user A follows no site and is in no team")
    word = sites[0]["name"].split()[0]
    found = await live.ok("a", "list_files", kind="site", query=word)
    assert sites[0]["id"] in [s["id"] for s in found["items"]], found
    live.finding("site_search", found.get("note") or "Microsoft Search answered")


# -- excel ----------------------------------------------------------------------------------------

async def test_excel_write_and_read_back(live):
    folder = await live.files_folder()
    name = f"{live.run}.xlsx"
    await live.put_file(folder, name, builders.xlsx(live.run))
    path = f"{folder}/{name}"

    async def readable():
        error, _, data = await live.invoke("a", "get_range", file=path)
        return data if not error else None
    book = await live.poll(readable, timeout=45, what="the uploaded workbook to open")
    assert "Sheet1" in [s["name"] for s in book["sheets"]]

    values = [["h1", "h2"], [1, 2]]
    wrote = await live.ok("a", "save_range", file=path, sheet="Sheet1", address="A1:B2", values=values)
    assert wrote["written_cells"] == 4
    read = await live.ok("a", "get_range", file=path, sheet="Sheet1", address="A1:B2")
    assert unfenced_rows(read["rows"]) == values

    # A workbook holding a table: its structure still reads (rows/$count answers 400; measured).
    made = await live.raw("a", "POST", f"/me/drive/root:/{path}:/workbook/tables/add",
                          json={"address": "Sheet1!A1:B2", "hasHeaders": True})
    assert made.status_code in (200, 201), live.error_of(made)
    book = await live.ok("a", "get_range", file=path)
    [table] = book["tables"]
    assert table["columns"] == ["h1", "h2"] and table["rows"] == 1, table


# -- teams ----------------------------------------------------------------------------------------

async def test_teams_message_arrives_and_soft_delete_is_recorded(live):
    b = await live.address("b")
    marker = f"{live.run} hello {uuid.uuid4().hex[:8]}"
    sent = await live.ok("a", "send_message", to=[b], body=marker, user_confirmed=True)
    chat, message = sent["chat"], sent["message_id"]
    assert chat and message

    async def seen_by_b():
        chats = await live.ok("b", "list_conversations", kind="chat", limit=25)
        if chat not in [c["id"] for c in chats["items"]]:
            return None
        got = await live.ok("b", "get_conversation", chat=chat, messages=10)
        return any(marker in (m.get("text") or "") for m in got["messages"])
    await live.poll(seen_by_b, timeout=ARRIVE, what="the Teams message in B's chat")

    deleted = await live.invoke("a", "update_message", id=message, chat=chat, action="delete")
    live.finding("chat_soft_delete_tool", "ok" if not deleted.error else deleted.text[:300])
    if deleted.error:
        pytest.skip(f"soft delete refused: {deleted.text[:200]}")

    async def hidden():
        got = await live.ok("a", "get_conversation", chat=chat, messages=10)
        return not any(marker in (m.get("text") or "") for m in got["messages"])
    await live.poll(hidden, timeout=30, what="the deleted message to disappear")
    restored = await live.invoke("a", "update_message", id=message, chat=chat, action="restore")
    live.finding("chat_soft_delete_restore_tool", "ok" if not restored.error else restored.text[:300])
    assert not restored.error, restored.text

    async def back():
        got = await live.ok("a", "get_conversation", chat=chat, messages=10)
        return any(marker in (m.get("text") or "") for m in got["messages"])
    await live.poll(back, timeout=30, what="the restored message")

    png = base64.b64encode(builders.png()).decode()
    pic = await live.ok("a", "send_message", to=[chat], body_format="html", user_confirmed=True,
                        body=f'<p>{live.run} picture</p><img alt="dot" src="data:image/png;base64,{png}">')
    live.later_tool("a", "update_message", id=pic["message_id"], chat=chat, action="delete")

    async def pictured():
        got = await live.ok("b", "get_conversation", chat=chat, messages=10)
        return any(f"{live.run} picture" in (m.get("text") or "") and m.get("inline_images") for m in got["messages"])
    await live.poll(pictured, timeout=ARRIVE, what="the message with a hosted image in B's chat")
    await live.ok("a", "update_message", id=pic["message_id"], chat=chat, action="delete")


async def test_a_deleted_channel_post_is_restored(live):
    teams = (await live.ok("a", "list_conversations", kind="channel", limit=25))["items"]
    if not teams:
        pytest.skip("user A is in no team")
    team = teams[0]["team"]
    channel = next((c for c in teams[0]["channels"] if c == "General"), teams[0]["channels"][0])
    marker = f"{live.run} channel {uuid.uuid4().hex[:8]}"
    sent = await live.ok("a", "send_message", to=[f"{team}/{channel}"], body=marker, user_confirmed=True)
    post = sent["message_id"]
    live.later_tool("a", "update_message", id=post, channel=f"{team}/{channel}", action="delete")

    async def listed(present: bool):
        got = await live.ok("a", "get_conversation", team=team, channel=channel, messages=20)
        return any(marker in (m.get("text") or "") for m in got["messages"]) == present
    await live.poll(lambda: listed(True), timeout=30, what="the channel post")
    await live.ok("a", "update_message", id=post, channel=f"{team}/{channel}", action="delete")
    await live.poll(lambda: listed(False), timeout=30, what="the deleted post to disappear")
    await live.ok("a", "update_message", id=post, channel=f"{team}/{channel}", action="restore")
    await live.poll(lambda: listed(True), timeout=30, what="the restored post")
    page = await live.ok("a", "get_conversation", team=team, channel=channel, messages=1)
    assert "has_more" in page and (page["has_more"] is False or page.get("cursor") or page.get("next_before"))
    await live.ok("a", "update_message", id=post, channel=f"{team}/{channel}", action="delete")


# -- tasks ----------------------------------------------------------------------------------------

async def test_todo_task_lifecycle(live):
    lists = await live.ok("a", "list_tasks", source="todo", overview=True)
    default = next(i["name"] for i in lists["items"] if i.get("default"))
    title = f"{live.run} task"
    saved = await live.ok("a", "save_task", source="todo", title=title)
    task = saved["id"]
    live.later_tool("a", "delete_task", id=task, source="todo", list=default, user_confirmed=True)

    found = await live.ok("a", "list_tasks", source="todo", query=live.run)
    assert [t["title"] for t in found["items"]] == [title]
    await live.ok("a", "save_task", source="todo", id=task, status="completed")
    done = await live.ok("a", "list_tasks", source="todo", query=live.run, status="completed")
    assert done["items"] and done["items"][0]["status"] == "completed"

    error, text, _ = await live.invoke("a", "delete_task", id=task, source="todo", list=default)
    assert error
    await live.ok("a", "delete_task", id=task, source="todo", list=default, user_confirmed=True)
    gone = await live.ok("a", "list_tasks", source="todo", query=live.run, status="all")
    assert gone["items"] == []


async def test_planner_task_lifecycle(live):
    plans = (await live.ok("a", "list_tasks", source="planner", overview=True))["items"]
    if not plans:
        pytest.skip("no Planner plan in the test tenant")
    plan = plans[0]
    title = f"{live.run} planner"
    extra = {"bucket": plan["buckets"][0]} if plan.get("buckets") else {}
    task = (await live.ok("a", "save_task", source="planner", plan=plan["name"], title=title, **extra))["id"]
    live.later_tool("a", "delete_task", id=task, source="planner", user_confirmed=True)

    found = await live.ok("a", "list_tasks", source="planner", plan=plan["name"], assigned_to_me=False,
                          status="all", query=live.run)
    assert [t["title"] for t in found["items"]] == [title]
    await live.ok("a", "save_task", source="planner", id=task, title=f"{title} changed", status="in_progress")
    again = await live.ok("a", "list_tasks", source="planner", plan=plan["name"], assigned_to_me=False,
                          status="all", query=live.run)
    assert again["items"][0]["title"] == f"{title} changed" and again["items"][0]["status"] == "in_progress"

    stale = await live.raw("a", "PATCH", f"/planner/tasks/{task}", json={"title": "stale"},
                           headers={"If-Match": 'W/"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBAWCc="'})
    live.finding("planner_stale_etag_status", live.error_of(stale))
    assert stale.status_code in (409, 412), live.error_of(stale)

    await live.ok("a", "delete_task", id=task, source="planner", user_confirmed=True)
    gone = await live.ok("a", "list_tasks", source="planner", plan=plan["name"], assigned_to_me=False,
                         status="all", query=live.run)
    assert gone["items"] == []


# -- notes ----------------------------------------------------------------------------------------

async def test_onenote_page_create_append_read_delete(live):
    notebook, section = await live.notes_target()
    first, second = f"FIRSTPART{uuid.uuid4().hex[:8]}", f"SECONDPART{uuid.uuid4().hex[:8]}"
    title = f"{live.run} page"
    created = await live.ok("a", "save_note_page", title=title, body=first, section=section, notebook=notebook)
    page = created["id"]
    live.later_tool("a", "delete_note_page", id=page, user_confirmed=True)

    await live.ok("a", "save_note_page", id=page, action="append", body=second)

    async def both():
        error, _, data = await live.invoke("a", "get_note_page", id=page)
        return data if not error and first in data["text"] and second in data["text"] else None
    read = await live.poll(both, timeout=60, what="both parts on the page")
    assert read["title"] == title

    error, _, _ = await live.invoke("a", "delete_note_page", id=page)
    assert error
    await live.ok("a", "delete_note_page", id=page, user_confirmed=True)
    pages = await live.ok("a", "list_notes", section=section, notebook=notebook, query=live.run)
    assert page not in [p["id"] for p in pages["items"]]


# -- contacts -------------------------------------------------------------------------------------

async def test_contact_save_find_delete(live):
    name = f"Livetest {live.tag}"
    saved = await live.ok("a", "save_contact", name=name, emails=[f"{live.tag}@example.com"], company="Live Test")
    contact = saved["id"]
    live.later_tool("a", "delete_contact", id=contact, user_confirmed=True)

    async def listed():
        data = await live.ok("a", "list_people", query=live.tag, source="contacts")
        return [p for p in data["items"] if live.tag in p.get("name", "")]
    try:
        hits = await live.poll(listed, timeout=120, every=6, what="the new contact in list_people")
    except AssertionError:
        live.finding("people_search_lag_contact", "not found by list_people within 120 s")
        raise
    assert hits[0]["source"] == "contacts"

    error, _, _ = await live.invoke("a", "delete_contact", id=contact)
    assert error
    await live.ok("a", "delete_contact", id=contact, user_confirmed=True)
    # Measured: Graph keeps a deleted contact, in Deleted Items, under the same id (200, not 404).
    gone = (await live.raw("a", "GET", f"/me/contacts/{contact}")).json()
    deleted_items = (await live.raw("a", "GET", "/me/mailFolders/deleteditems")).json()["id"]
    assert gone["parentFolderId"] == deleted_items


async def test_contact_phone_is_replaced_not_added(live):
    saved = await live.ok("a", "save_contact", name=f"Livetest {live.tag}", phones=["+47 111 22 333"])
    live.later_tool("a", "delete_contact", id=saved["id"], user_confirmed=True)
    assert saved["phones"] == ["+47 111 22 333"]
    changed = await live.ok("a", "save_contact", id=saved["id"], phones=["+47 999 88 777"], remove_phones=["+4711122333"])
    assert changed["phones"] == ["+47 999 88 777"], changed
    held = (await live.raw("a", "GET", f"/me/contacts/{saved['id']}")).json()
    assert held["businessPhones"] == ["+47 999 88 777"]


async def test_deleting_twice_says_it_was_already_deleted_and_a_draft_flags_outside_recipients(live):
    # Nothing is sent: both are drafts.
    b = await live.address("b")
    draft = await live.ok("a", "save_draft", to=[b, f"{live.tag}@example.com"], subject=f"{live.run} outside", body="x")
    assert draft["outside_organisation"] == [f"{live.tag}@example.com"], draft
    first = await live.ok("a", "update_messages", ids=[draft["draft_id"]], action="delete")
    assert first["done"] == 1 and "already_there" not in first
    second = await live.ok("a", "update_messages", ids=[draft["draft_id"]], action="delete")
    assert second["done"] == 0 and second["already_there"] == 1, second


# -- automatic reply ------------------------------------------------------------------------------

async def test_automatic_reply_scheduled_then_off(live):
    tz = await local_zone(live, "a")
    tomorrow = datetime.now(tz).date() + timedelta(days=1)
    args = {"status": "scheduled", "internal_message": f"{live.run} away", "external_message": f"{live.run} away",
            "audience": "none", "start": f"{tomorrow} 08:00", "end": f"{tomorrow} 18:00"}

    # Whatever the person had before is put back, whatever happens below.
    before = (await live.raw("a", "GET", "/me/mailboxSettings/automaticRepliesSetting")).json()
    before.pop("@odata.context", None)
    live.later("restore the automatic reply of a",
               lambda: live.raw("a", "PATCH", "/me/mailboxSettings", json={"automaticRepliesSetting": before}))

    error, text, _ = await live.invoke("a", "save_automatic_reply", **args)
    assert error and "user_confirmed" in text

    await live.ok("a", "save_automatic_reply", user_confirmed=True, **args)

    async def scheduled():
        return (await live.ok("a", "get_connection")).get("automatic_reply") == "scheduled"
    await live.poll(scheduled, timeout=30, what="get_connection to show the scheduled reply")

    await live.ok("a", "save_automatic_reply", status="off")

    async def off():
        return (await live.ok("a", "get_connection")).get("automatic_reply") == "off"
    await live.poll(off, timeout=30, what="get_connection to show the reply off")


# -- recordings -----------------------------------------------------------------------------------

@pytest.mark.skipif(not os.environ.get("M365_LIVE_RECORDED_EVENT"),
                    reason="M365_LIVE_RECORDED_EVENT (an event id in A's calendar with a recording) is not set")
async def test_meeting_recording_and_transcript_links(live):
    got = await live.ok("a", "get_event", id=os.environ["M365_LIVE_RECORDED_EVENT"], include_recording=True)
    assert got["recording_available"] is True, got.get("recording_note")
    live.finding("transcripts_graph_access",
                 f"{len(got['transcripts'])} transcript(s) listed" if got.get("transcripts")
                 else got.get("recording_note") or "no transcript and no note")
    assert got.get("recordings") or got.get("transcripts")
    for t in got.get("transcripts", [])[:1]:
        res = await live.fetch(live.link_in(t["download"]))
        assert res.status_code == 200 and res.content.lstrip().startswith(b"WEBVTT"), res.content[:80]
    for r in got.get("recordings", [])[:1]:
        res = await live.fetch(live.link_in(r["download"]), headers={"Range": "bytes=0-1023"})
        assert res.status_code in (200, 206) and len(res.content) <= 2048
