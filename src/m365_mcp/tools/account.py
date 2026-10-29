"""Who the person is and what this connection can do; finding colleagues and contacts."""

from __future__ import annotations

from dataclasses import replace
from typing import Annotated, Any, Literal

from pydantic import Field

from m365_mcp.core.context import call, scope_granted
from m365_mcp.core.errors import invalid
from m365_mcp.core.project import clamp_limit, compact, page_result
from m365_mcp.core.resolve import Contacts, Who, contacts, people
from m365_mcp.core.time import from_graph
from m365_mcp.registry import AREAS, tool

MAX_PEOPLE = 100  # matches listed across pages; the ranking is read again from the top for each page

# The permissions each area needs to work at all; reported by get_connection.
AREA_SCOPES: dict[str, list[str]] = {
    "mail": ["Mail.ReadWrite", "Mail.Send"],
    "calendar": ["Calendars.ReadWrite"],
    "contacts": ["Contacts.ReadWrite"],
    "files": ["Files.ReadWrite.All"],
    "lists": ["Sites.ReadWrite.All"],
    "excel": ["Files.ReadWrite"],
    "teams": ["Chat.ReadWrite"],
    "recordings": ["OnlineMeetingTranscript.Read.All"],
    "tasks": ["Tasks.ReadWrite"],
    "notes": ["Notes.ReadWrite"],
    "bookings": ["BookingsAppointment.ReadWrite.All"],
}


@tool(area="account", title="Microsoft 365 connection", scopes=["User.Read"], read_only=True, idempotent=True)
async def get_connection() -> dict[str, Any]:
    """Who is signed in to Microsoft 365, their time zone and working hours, their automatic reply,
    and which areas this connection can use. Call it first when another tool fails for a permission,
    or when the person asks what the assistant can do in Microsoft 365. Areas whose permission was
    not granted are listed with the missing permission, so the person knows to ask an administrator."""
    async with call("get_connection") as c:
        from m365_mcp.core.context import runtime

        me = await c.graph.get("/me", params={"$select": "displayName,mail,userPrincipalName,jobTitle,department"})
        out: dict[str, Any] = compact({
            "name": me.get("displayName"),
            "address": me.get("mail") or me.get("userPrincipalName"),
            "job_title": me.get("jobTitle"),
            "department": me.get("department"),
            "time_zone": c.tz_name,
        })
        if scope_granted("MailboxSettings.Read", c.person.scopes):
            s = await c.graph.get("/me/mailboxSettings", params={"$select": "workingHours,automaticRepliesSetting"})
            wh = s.get("workingHours") or {}
            # Unset is said, as find_meeting_times says it: the model then knows the default is in use.
            out["working_hours"] = compact({
                "days": wh.get("daysOfWeek"), "start": (wh.get("startTime") or "")[:5],
                "end": (wh.get("endTime") or "")[:5],
            }) if wh.get("daysOfWeek") else "not set in Outlook"
            ar = s.get("automaticRepliesSetting") or {}
            out["automatic_reply"] = {"disabled": "off", "alwaysEnabled": "on", "scheduled": "scheduled"}.get(
                ar.get("status", "disabled"), "off")
            if ar.get("status") == "scheduled":
                # The period, in the person's own zone: "scheduled" alone cannot remind them when it runs.
                period = compact({"from": from_graph(ar.get("scheduledStartDateTime"), c.tz),
                                  "until": from_graph(ar.get("scheduledEndDateTime"), c.tz)})
                if period:
                    out["automatic_reply_period"] = period
        settings = runtime().settings
        on = set(settings.areas)
        areas: dict[str, str] = {}
        for area in AREAS:
            if area == "account":
                continue
            if area not in on:
                areas[area] = "off"
                continue
            missing = [s for s in AREA_SCOPES.get(area, []) if not scope_granted(s, c.person.scopes)]
            areas[area] = f"missing permission {', '.join(missing)}" if missing else "on"
        out["areas"] = areas
        return out


@tool(area="account", title="Find people", scopes=["People.Read"], read_only=True, idempotent=True, returns_list=True)
async def list_people(
    query: Annotated[str, Field(description="A name, part of a name, or an address.")],
    source: Annotated[Literal["all", "contacts", "directory"], Field(
        description="contacts: only the person's own Outlook contacts; directory: only the organisation.")] = "all",
    include_org: Annotated[bool, Field(description="Add office, manager and phone; one extra request per person.")] = False,
    limit: Annotated[int, Field(description="1 to 25.")] = 10,
    cursor: Annotated[str | None, Field(description="From the previous page's result.")] = None,
) -> dict[str, Any]:
    """Find colleagues and personal contacts by name or address, ranked by how much the person works
    with them. Use it before writing to or inviting someone whose address is not known, and show the
    candidates when several match. Each result says whether it came from the person's contacts or
    the organisation's directory."""
    if not query.strip():
        raise invalid("Give a name or an address to look for.")
    n, extra = clamp_limit(limit)
    async with call("list_people") as c:
        state = c.cursor_in(cursor)
        asked = [query.strip(), source, include_org]
        if state is not None and state.get("args") != asked:
            raise invalid("This cursor continues another search for people. Repeat the call without the cursor "
                          "to start this one.")
        skip = int((state or {}).get("skip", 0))
        n = max(0, min(n, MAX_PEOPLE - skip))  # the last page ends at the cap, it is not refused
        # Contacts are read from the person's contacts folder itself; the People index learns a new one
        # late or never. Directory people and groups come from People, then the directory.
        own = await contacts(c, query) if source != "directory" else Contacts()
        # The page and one more, from the top: the ranking is Microsoft's and has no cursor of its own.
        # The source filter drops some of what is read, so more is read until the page and one more
        # remain, or Microsoft has no more.
        want = skip + n + 1
        read = want
        complete = own.complete
        while True:
            raw = await people(c, query, limit=read, directory=source != "contacts")
            ranked = raw if source == "all" else [w for w in raw if (w.source == "contacts") == (source == "contacts")]
            # Groups People knows are taken before the source filter, which would drop them for contacts.
            groups = {w.address.casefold() for w in raw if w.is_group}
            # A contact holding an address that People knows as a group is that group.
            unique: dict[str, Who] = {}  # one entry per address, the first kept: a person saved in two folders
            for w in own:
                unique.setdefault(w.address.casefold(), replace(w, is_group=True) if w.address.casefold() in groups else w)
            mine = list(unique.values())
            seen = {w.address.casefold() for w in mine}
            # Counted after the merge: a contact People also ranks is one match, not two.
            found = mine + [w for w in ranked if w.address.casefold() not in seen]
            if len(found) >= want or len(raw) < read:
                break
            if read >= 4 * MAX_PEOPLE:
                extra["truncated"] = True  # read as far as allowed; more may match
                break
            read = min(read * 2, 4 * MAX_PEOPLE)
        items = []
        for w in found[skip:skip + n]:
            item = compact({"name": w.name, "address": w.address, "job_title": w.job_title,
                            "department": w.department, "source": w.source, "group": w.is_group,
                            "contact_id": w.contact_id if w.source == "contacts" else None})
            if include_org and not w.is_group and w.source != "contacts":
                try:
                    u = await c.graph.get(f"/users/{w.address}", params={
                        "$select": "officeLocation,businessPhones,mobilePhone"})
                    item.update(compact({"office": u.get("officeLocation"),
                                         "phone": (u.get("businessPhones") or [None])[0] or u.get("mobilePhone")}))
                    mgr = await c.graph.get(f"/users/{w.address}/manager", params={"$select": "displayName"})
                    item["manager"] = mgr.get("displayName")
                except Exception:  # noqa: BLE001 - org details are best effort; the person is still found
                    pass
            items.append(item)
        more = None
        if len(found) > skip + n:
            if skip + n < MAX_PEOPLE:
                more = c.cursor_out({"args": asked, "skip": skip + n})
            else:
                extra["truncated"] = True  # the cap: the most relevant MAX_PEOPLE have been listed
        if not complete:
            extra["truncated"] = True
            extra["contacts_note"] = ("Not every contact could be searched; one past those read may be missing. "
                                     "Use the person's address instead.")
        if not items and not skip and not extra.get("truncated"):  # an incomplete search cannot say "nobody"
            where = {"all": "the person's contacts and the organisation's directory", "contacts": "the person's contacts only",
                     "directory": "the organisation's directory only"}[source]
            extra.setdefault("note", f"Nobody matched in {where}. Try part of the name, or the address.")
        return page_result(items, more, **extra)
