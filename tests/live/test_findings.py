"""Probes that MEASURE what Microsoft's documentation leaves unknown, and record it.

Each test writes one or more lines to `live-findings.json` (name -> {observed, date}); `python
scripts/live_findings_to_md.py` turns that into docs/graph-live-findings.md. A probe fails only when
it cannot measure (the request never reached Graph, or setup broke), not because Graph answered
something unexpected: the answer is the result.
"""

from __future__ import annotations

import asyncio
import re
from datetime import datetime, timedelta
from urllib.parse import urlsplit

import httpx
import pytest

import builders
from livekit import claims
from m365_mcp.core.time import zone
from m365_mcp.tools.teams import _member

pytestmark = pytest.mark.live

MB = 1024 * 1024


def observed_status(resp, live) -> str:
    return "ok" if resp.status_code < 300 else live.error_of(resp)


async def chat_probe(live) -> tuple[str, str]:
    """(chat id, A's user id): the one-to-one chat of A and B, created if it does not exist."""
    if "probe_chat" not in live.state:
        ids = {}
        for user in ("a", "b"):
            me = await live.raw(user, "GET", "/me", params={"$select": "id"})
            assert me.status_code == 200, live.error_of(me)
            ids[user] = me.json()["id"]
        made = await live.raw("a", "POST", "/chats", json={
            "chatType": "oneOnOne", "members": [_member(ids["a"]), _member(ids["b"])]})
        assert made.status_code in (200, 201), live.error_of(made)
        live.state["probe_chat"] = (made.json()["id"], ids["a"])
    return live.state["probe_chat"]


async def post_chat(live, html: str) -> object:
    chat, me = await chat_probe(live)
    resp = await live.raw("a", "POST", f"/chats/{chat}/messages", json={"body": {"contentType": "html", "content": html}})

    async def tidy() -> None:
        if resp.status_code == 201:
            await live.raw("a", "POST", f"/users/{me}/chats/{chat}/messages/{resp.json()['id']}/softDelete")
    live.later("probe chat message", tidy)
    return resp


# -- Teams ----------------------------------------------------------------------------------------

async def test_chat_send_permission(live):
    """ChatMessage.Send vs ChannelMessage.Send: what the token holds, whether a chat post succeeds, and
    whether a token lacking one can be had at all. A refresh asking for fewer scopes is the only way to
    get one without a new sign-in; what it returns is the answer."""
    scopes = set(claims(live.tokens["a"]["access_token"]).get("scp", "").split())
    resp = await post_chat(live, f"<p>{live.run} permission probe</p>")
    narrowed = await live.refresh_with("a", ["User.Read", "Chat.ReadWrite"])
    held = set(claims(narrowed).get("scp", "").split())
    live.finding("chat_send_scope", {
        "token holds ChatMessage.Send": "ChatMessage.Send" in scopes,
        "token holds ChannelMessage.Send": "ChannelMessage.Send" in scopes,
        "chat post": observed_status(resp, live),
        "refresh asking only User.Read Chat.ReadWrite still holds ChatMessage.Send": "ChatMessage.Send" in held,
        "refresh asking only User.Read Chat.ReadWrite still holds ChannelMessage.Send": "ChannelMessage.Send" in held,
    })
    assert resp.status_code < 500


async def test_chat_soft_delete_endpoints_and_threaded_replies(live):
    chat, me = await chat_probe(live)
    sent = await post_chat(live, f"<p>{live.run} soft delete probe</p>")
    assert sent.status_code == 201, live.error_of(sent)
    message = sent.json()["id"]

    results = {}
    for label, path in (("users/{me}", f"/users/{me}/chats/{chat}/messages/{message}"),
                        ("me", f"/me/chats/{chat}/messages/{message}")):
        r = await live.raw("a", "POST", f"{path}/softDelete")
        results[f"softDelete via {label}"] = observed_status(r, live)
        if r.status_code < 300:
            u = await live.raw("a", "POST", f"{path}/undoSoftDelete")
            results[f"undoSoftDelete via {label}"] = observed_status(u, live)
            break
    live.finding("chat_soft_delete_endpoint", results)

    replied = await live.raw("a", "POST", f"/chats/{chat}/messages/{message}/replies",
                             json={"body": {"content": f"{live.run} reply probe"}})
    live.finding("chat_threaded_replies", observed_status(replied, live))


async def test_teams_html_tags_that_survive(live):
    tags = ["p", "b", "strong", "i", "em", "u", "s", "h1", "h2", "ul", "ol", "li", "table", "thead", "tbody", "tr",
            "th", "td", "pre", "code", "blockquote", "a", "br", "hr", "span", "div", "font"]
    html = (f"<div>{live.run} tags<h1>h</h1><h2>h</h2><p><b>b</b><strong>s</strong><i>i</i><em>e</em><u>u</u><s>s</s>"
            '<span style="color:red">span</span><font color="red">font</font><br><a href="https://example.com">a</a></p>'
            "<ul><li>l</li></ul><ol><li>l</li></ol><blockquote>q</blockquote><pre>pre</pre><code>code</code><hr>"
            "<table><thead><tr><th>h</th></tr></thead><tbody><tr><td>c</td></tr></tbody></table></div>")
    sent = await post_chat(live, html)
    if sent.status_code != 201:
        live.finding("teams_html_tags", {"post refused": live.error_of(sent)})
        pytest.skip("the tag sample was refused as a whole; recorded")
    chat, _ = await chat_probe(live)
    got = await live.raw("a", "GET", f"/chats/{chat}/messages/{sent.json()['id']}")
    kept = sorted(set(re.findall(r"<([a-zA-Z][a-zA-Z0-9]*)", got.json()["body"]["content"])))
    live.finding("teams_html_tags", {"sent": tags, "kept": kept, "dropped": sorted(set(tags) - set(kept))})
    assert kept


# -- mail -----------------------------------------------------------------------------------------

async def test_search_combined_with_filter_orderby(live):
    base = {"$search": '"m365-live"', "$top": "5", "$select": "id,subject"}
    cases = {
        "$search alone": {},
        "$search + $filter on receivedDateTime": {"$filter": "receivedDateTime ge 2000-01-01T00:00:00Z"},
        "$search + $filter on isRead": {"$filter": "isRead eq false"},
        "$search + $orderby": {"$orderby": "receivedDateTime desc"},
    }
    out = {}
    for label, extra in cases.items():
        r = await live.raw("a", "GET", "/me/messages", params={**base, **extra})
        out[label] = observed_status(r, live) if r.status_code != 200 else f"ok, {len(r.json().get('value', []))} items"
    live.finding("search_with_filter", out)
    assert out["$search alone"].startswith("ok"), out


# -- calendar -------------------------------------------------------------------------------------

async def test_calendar_view_filter_and_orderby(live):
    tz = zone((await live.ok("a", "get_connection"))["time_zone"])
    start = datetime.now(tz).date()
    window = {"startDateTime": f"{start}T00:00:00Z", "endDateTime": f"{start + timedelta(days=14)}T00:00:00Z",
              "$top": "5", "$select": "id,subject"}
    cases = {
        "$filter only": {"$filter": "contains(subject,'x')"},
        "$orderby only": {"$orderby": "start/dateTime"},
        "$filter + $orderby": {"$filter": "contains(subject,'x')", "$orderby": "start/dateTime"},
    }
    out = {}
    for label, extra in cases.items():
        r = await live.raw("a", "GET", "/me/calendarView", params={**window, **extra})
        out[label] = observed_status(r, live)
    live.finding("calendarview_filter_orderby", out)
    assert set(out) == set(cases)


# -- files ----------------------------------------------------------------------------------------

async def test_pdf_conversion_of_a_3mb_document(live):
    folder = await live.files_folder()
    data = builders.docx(f"{live.run} large", padding_bytes=3 * MB)
    item = await live.put_file(folder, f"{live.run}-3mb.docx", data)
    plan = await live.ok("a", "download_file", id=item["id"], format="pdf")
    res = await live.fetch(live.link_in(plan["download"]))
    size = f"{len(data) / MB:.1f} MB docx"
    if res.status_code == 200 and res.content.startswith(b"%PDF"):
        live.finding("pdf_conversion_3mb", f"converted: {size} -> {len(res.content)} bytes of PDF")
    else:
        live.finding("pdf_conversion_3mb", f"refused: {size}: {res.status_code} {res.text[:200]}")


async def test_copy_monitor_host(live):
    folder = await live.files_folder()
    item = await live.put_file(folder, f"{live.run}-copy-source.txt", b"copy me")
    meta = await live.raw("a", "GET", f"/me/drive/items/{item['id']}", params={"$select": "parentReference"})
    parent = meta.json()["parentReference"]
    r = await live.raw("a", "POST", f"/me/drive/items/{item['id']}/copy", json={
        "parentReference": {"driveId": parent["driveId"], "id": parent["id"]}, "name": f"{live.run}-copy.txt"})
    assert r.status_code == 202, live.error_of(r)
    host = urlsplit(r.headers["location"]).hostname or ""
    # The first label names the tenant; the finding is about the suffix, and the document is public.
    shown = host if host == "graph.microsoft.com" else re.sub(r"^[^.]+?(-my)?\.", r"<tenant>\1.", host)
    live.finding("copy_monitor_host", {"host": shown, "suffix accepted by Graph.MONITOR_HOSTS":
                                       host.endswith((".sharepoint.com", ".onedrive.com", ".svc.ms")) or host == "graph.microsoft.com"})


# -- what deletion keeps ----------------------------------------------------------------------------

async def deleted_items_count(live) -> int:
    r = await live.raw("a", "GET", "/me/mailFolders/deleteditems", params={"$select": "totalItemCount"})
    assert r.status_code == 200, live.error_of(r)
    return r.json()["totalItemCount"]


async def test_where_a_deleted_contact_goes(live):
    before = await deleted_items_count(live)
    name = f"Deleted {live.tag}"
    contact = (await live.ok("a", "save_contact", name=name, emails=[f"{live.tag}@example.com"]))["id"]
    await live.ok("a", "delete_contact", id=contact, user_confirmed=True)
    await asyncio.sleep(10)
    after = await deleted_items_count(live)
    direct = await live.raw("a", "GET", f"/me/contacts/{contact}")
    bin_ = await live.raw("a", "GET", "/me/contactFolders/deleteditems/contacts",
                          params={"$top": "50", "$select": "displayName"})
    in_bin = None
    if bin_.status_code == 200:
        in_bin = any(c.get("displayName") == name for c in bin_.json().get("value", []))
    live.finding("deleted_contact_location", {
        "GET the contact afterwards": direct.status_code,
        "Deleted Items folder item count before/after": [before, after],
        "contactFolders/deleteditems listing": observed_status(bin_, live) if bin_.status_code != 200 else f"ok, contact present: {in_bin}",
    })


async def test_whether_a_deleted_todo_task_is_recoverable(live):
    lists = (await live.raw("a", "GET", "/me/todo/lists")).json()["value"]
    default = next(x for x in lists if x.get("wellknownListName") == "defaultList")
    title = f"{live.run} deleted task"
    task = (await live.ok("a", "save_task", source="todo", title=title))["id"]
    await live.ok("a", "delete_task", id=task, source="todo", list=default["displayName"], user_confirmed=True)
    got = await live.raw("a", "GET", f"/me/todo/lists/{default['id']}/tasks/{task}")
    listing = await live.raw("a", "GET", f"/me/todo/lists/{default['id']}/tasks",
                             params={"$filter": f"title eq '{title}'"})
    live.finding("deleted_todo_task", {
        "GET the task afterwards": got.status_code,
        "listed afterwards": [t["title"] for t in listing.json().get("value", [])] if listing.status_code == 200 else observed_status(listing, live),
    })


async def test_todo_paging(live):
    """To Do's page size: the default, and whether $top is capped."""
    made = await live.raw("a", "POST", "/me/todo/lists", json={"displayName": f"{live.run} paging"})
    assert made.status_code == 201, live.error_of(made)
    lid = made.json()["id"]
    live.later("paging task list", lambda: live.raw("a", "DELETE", f"/me/todo/lists/{lid}"))
    gate = asyncio.Semaphore(4)

    async def one(i: int) -> int:
        async with gate:
            return (await live.raw("a", "POST", f"/me/todo/lists/{lid}/tasks", json={"title": f"paging probe {i}"})).status_code
    codes = await asyncio.gather(*(one(i) for i in range(60)))
    assert codes.count(201) == 60, codes
    observed = {}
    for label, params in (("default", {}), ("$top=5", {"$top": 5}), ("$top=1000", {"$top": 1000})):
        page = (await live.raw("a", "GET", f"/me/todo/lists/{lid}/tasks", params=params)).json()
        observed[label] = {"first page": len(page.get("value", [])), "next link": "@odata.nextLink" in page}
    live.finding("todo_paging", observed)


async def test_whether_a_deleted_onenote_page_is_recoverable(live):
    notebook, section = await live.notes_target()
    title = f"{live.run} deleted page"
    page = (await live.ok("a", "save_note_page", title=title, body="gone soon", section=section, notebook=notebook))["id"]
    await live.ok("a", "delete_note_page", id=page, user_confirmed=True)
    got = await live.raw("a", "GET", f"/me/onenote/pages/{page}")
    content = await live.raw("a", "GET", f"/me/onenote/pages/{page}/content")
    listed = await live.ok("a", "list_notes", section=section, notebook=notebook, query=title)
    live.finding("deleted_onenote_page", {
        "GET the page afterwards": got.status_code,
        "GET its content afterwards": content.status_code,
        "listed in its section afterwards": any(p["id"] == page for p in listed["items"]),
    })


async def test_whether_a_deleted_list_item_is_recoverable(live):
    sites = (await live.ok("a", "list_files", kind="site", include_lists=True, limit=25))["items"]
    for site in sites:
        for lst in site.get("lists", []):
            made = await live.invoke("a", "save_list_item", site=site["id"], list=lst["id"], fields=[f"Title={live.run}"])
            if made.error:
                continue
            row = made.data["id"]
            await live.ok("a", "delete_list_item", site=site["id"], list=lst["id"], id=row, user_confirmed=True)
            got = await live.raw("a", "GET", f"/sites/{site['id']}/lists/{lst['id']}/items/{row}")
            # v1.0 has no recycle bin for a site (400); beta reads it. Only this probe uses beta.
            binned = await live.raw("a", "GET", f"https://graph.microsoft.com/beta/sites/{site['id']}/recycleBin/items")
            live.finding("deleted_list_item", {
                "GET the item afterwards": got.status_code,
                "in the site recycle bin (beta /recycleBin/items)": (
                    any(i.get("title") == live.run for i in binned.json().get("value", []))
                    if binned.status_code == 200 else observed_status(binned, live)),
            })
            return
    pytest.skip("no SharePoint list with a writable Title column reachable by user A")


# -- found by QA, round 2 -------------------------------------------------------------------------

async def test_list_item_fields_expand_syntax(live):
    """How Graph wants the columns of list items selected inside $expand."""
    sites = (await live.ok("a", "list_files", kind="site", include_lists=True, limit=25))["items"]
    for site in sites:
        for lst in site.get("lists", []):
            base = f"/sites/{site['id']}/lists/{lst['id']}/items"
            observed = {}
            for label, expand in (("fields(select=Title)", "fields(select=Title)"),
                                  ("fields($select=Title)", "fields($select=Title)")):
                resp = await live.raw("a", "GET", base, params={"$top": "1", "$expand": expand})
                observed[label] = observed_status(resp, live)
            live.finding("list_items_expand_select", observed)
            return
    pytest.skip("no SharePoint list reachable by user A")


async def test_onenote_paging(live):
    """Whether OneNote sends a next link once $top is given, and whether it honours $skip."""
    notebook, section = await live.notes_target()
    for i in range(3):
        await live.ok("a", "save_note_page", title=f"{live.run} paging {i}", body="x", section=section, notebook=notebook)
    secs = (await live.raw("a", "GET", "/me/onenote/sections", params={"$select": "id,displayName"})).json()["value"]
    sid = next(s["id"] for s in secs if s["displayName"] == section)
    base = f"/me/onenote/sections/{sid}/pages"
    order = {"$select": "id", "$orderby": "lastModifiedDateTime desc"}
    first = (await live.raw("a", "GET", base, params={**order, "$top": "1"})).json()
    second = (await live.raw("a", "GET", base, params={**order, "$top": "1", "$skip": "1"})).json()
    live.finding("onenote_paging", {
        "$top=1: items": len(first.get("value", [])), "$top=1: next link": "@odata.nextLink" in first,
        "$skip=1 gives another page": bool(second.get("value")) and second["value"][0]["id"] != first["value"][0]["id"],
    })


async def test_onenote_title_after_a_fast_body_change(live):
    """What a body PATCH right after creation does to the page's listed title, with and without a title
    op. It does not happen every time, so each way is tried three times and counted."""
    notebook, section = await live.notes_target()
    secs = (await live.raw("a", "GET", "/me/onenote/sections", params={"$select": "id,displayName"})).json()["value"]
    sid = next(s["id"] for s in secs if s["displayName"] == section)
    observed: dict[str, dict[str, int]] = {}
    for label, with_title in (("body only", False), ("body and title", True)):
        counts = observed.setdefault(label, {"title kept": 0, "title emptied": 0})
        for attempt in range(3):
            title = f"{live.run} title {int(with_title)}{attempt}"
            doc = f"<!DOCTYPE html><html><head><title>{title}</title></head><body><p>x</p></body></html>"
            made = await live.raw("a", "POST", base_path(sid), content=doc.encode(), headers={"Content-Type": "text/html"})
            assert made.status_code == 201, live.error_of(made)
            page = made.json()["id"]
            live.later("probe page", lambda page=page: live.raw("a", "DELETE", f"/me/onenote/pages/{page}"))
            ops = [{"target": "body", "action": "append", "content": "<p>y</p>"}]
            if with_title:
                ops.insert(0, {"target": "title", "action": "replace", "content": title})
            for _ in range(6):
                resp = await live.raw("a", "PATCH", f"/me/onenote/pages/{page}/content", json=ops)
                if resp.status_code != 404:
                    break
                await asyncio.sleep(2)
            await asyncio.sleep(8)
            got = (await live.raw("a", "GET", f"/me/onenote/pages/{page}", params={"$select": "title"})).json()
            counts["title kept" if got.get("title") == title else "title emptied"] += 1
    live.finding("onenote_title_after_fast_patch", observed)


def base_path(section_id: str) -> str:
    return f"/me/onenote/sections/{section_id}/pages"


async def test_find_meeting_times_with_half_the_people_busy(live):
    """Whether findMeetingTimes offers a time when one of two required attendees is busy, at 50 and 0 percent."""
    day = (datetime.now() + timedelta(days=21)).date().isoformat()
    made = await live.raw("b", "POST", "/me/events", json={
        "subject": f"{live.run} busy", "showAs": "busy",
        "start": {"dateTime": f"{day}T10:00:00", "timeZone": "UTC"}, "end": {"dateTime": f"{day}T12:00:00", "timeZone": "UTC"}})
    assert made.status_code == 201, live.error_of(made)
    eid = made.json()["id"]
    live.later("busy event", lambda: live.raw("b", "DELETE", f"/me/events/{eid}"))
    b = await live.address("b")

    async def busy_seen() -> bool:
        body = (await live.raw("a", "POST", "/me/calendar/getSchedule", json={
            "schedules": [b], "startTime": {"dateTime": f"{day}T10:00:00", "timeZone": "UTC"},
            "endTime": {"dateTime": f"{day}T11:00:00", "timeZone": "UTC"}, "availabilityViewInterval": 60})).json()
        return "2" in (body.get("value") or [{}])[0].get("availabilityView", "")
    await live.poll(busy_seen, timeout=120, what="the busy block in free/busy")
    observed = {}
    for pct in (50, 0):
        body = (await live.raw("a", "POST", "/me/findMeetingTimes", json={
            "attendees": [{"type": "required", "emailAddress": {"address": b}}],
            "timeConstraint": {"activityDomain": "unrestricted", "timeslots": [{
                "start": {"dateTime": f"{day}T09:00:00", "timeZone": "UTC"},
                "end": {"dateTime": f"{day}T13:00:00", "timeZone": "UTC"}}]},
            "meetingDuration": "PT60M", "maxCandidates": 25, "isOrganizerOptional": True,
            "minimumAttendeePercentage": pct})).json()
        starts = [s["meetingTimeSlot"]["start"]["dateTime"][11:16] for s in body.get("meetingTimeSuggestions", [])]
        observed[f"minimumAttendeePercentage={pct}"] = {
            "offered while busy (10:00 or 11:00)": any(t in ("10:00", "11:00") for t in starts), "offered": starts}
    live.finding("find_meeting_times_partial", observed)


async def test_copy_onto_a_name_that_exists(live):
    """What copy answers when the target folder already holds the name: with a new name given, and without."""
    folder = await live.files_folder()
    item = await live.put_file(folder, f"{live.run} copy.txt", b"copy me")
    here = (await live.raw("a", "GET", f"/me/drive/root:/{folder}", params={"$select": "id,parentReference"})).json()
    sub = await live.raw("a", "POST", f"/me/drive/items/{here['id']}/children", json={"name": "sub", "folder": {}})
    assert sub.status_code == 201, live.error_of(sub)
    drive, target = here["parentReference"]["driveId"], sub.json()["id"]
    observed = {}
    for label, extra in (("same name", {}), ("same name again", {}), ("new name given twice", {"name": "renamed.txt"}),
                         ("new name given twice, again", {"name": "renamed.txt"})):
        resp = await live.raw("a", "POST", f"/me/drive/items/{item['id']}/copy",
                              json={"parentReference": {"driveId": drive, "id": target}, **extra})
        entry: dict[str, object] = {"answer": resp.status_code}
        monitor = resp.headers.get("Location")
        for _ in range(15):
            if not monitor:
                break
            await asyncio.sleep(2)
            async with httpx.AsyncClient(timeout=30) as client:
                got = (await client.get(monitor, headers={"Accept": "application/json"})).json()
            if got.get("status") in ("completed", "failed"):
                entry["monitor"] = {"status": got["status"], "code": (got.get("error") or {}).get("code")}
                break
        observed[label] = entry
    live.finding("copy_name_clash_monitor", observed)
