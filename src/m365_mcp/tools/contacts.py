"""Personal Outlook contacts. Reading and finding people is list_people; this area changes them."""

from __future__ import annotations

import re
from typing import Annotated, Any

from pydantic import Field

from m365_mcp.core.confirm import require
from m365_mcp.core.context import Call, call
from m365_mcp.core.errors import invalid
from m365_mcp.core.resolve import one
from m365_mcp.graph.errors import GraphError
from m365_mcp.registry import tool

MAX_FOLDER_REQUESTS = 20  # pages of contact folders read; past that a folder name is not resolved
MAX_EMAILS = 3  # Outlook keeps three email addresses on a contact and refuses a fourth
_ADDRESS = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s.]+$")

NOT_A_CONTACT = ("Only personal contacts are changed here, and no personal contact has this id: colleagues from "
                 "the directory are not contacts. Find the contact's contact_id with list_people source=contacts. "
                 "Nothing was changed.")
UNPLACED = ("This contact could not be placed in any of the contact folders read here, so it cannot be told from a "
            "deleted one. Change it in Outlook. Nothing was changed.")
DELETED = ("This contact was deleted and sits in Deleted Items. The person can restore it there in Outlook, and "
           "then it can be changed. Nothing was changed.")


def _phone_key(p: str) -> str:
    """Two spellings of one number (+47 123 45, +4712345) are the same number."""
    return re.sub(r"\W", "", p).casefold()


def _phones_of(contact: dict[str, Any]) -> list[str]:
    return [*(contact.get("businessPhones") or []), *(contact.get("homePhones") or []),
            *([contact["mobilePhone"]] if contact.get("mobilePhone") else [])]


def _merge_emails(existing: list[dict[str, Any]], new: list[str], name: str) -> list[dict[str, Any]]:
    seen = {(e.get("address") or "").casefold() for e in existing}
    added = []
    for a in new:
        if a.strip() and a.strip().casefold() not in seen:
            seen.add(a.strip().casefold())  # one address given twice is added once
            added.append({"address": a.strip(), "name": name})
    return [*existing, *added]


async def _contact_folders(c: Call) -> tuple[list[dict[str, Any]], bool]:
    """Every contact folder with its path, the default one first, and whether that was all of them.
    Contacts permissions only: the mail folders are another area's (Deleted Items is one of them)."""
    default = await c.graph.get("/me/contactFolders/contacts", params={"$select": "id,displayName"})
    found: list[dict[str, Any]] = [{**default, "path": default.get("displayName") or "Contacts"}]
    seen = {default.get("id")}
    # The folders beside the default one, and those inside it: Graph's documentation puts the first
    # listing under the default folder, so each id is kept once whichever listing names it.
    queue: list[tuple[str, str]] = [("/me/contactFolders", ""),
                                    (f"/me/contactFolders/{default.get('id')}/childFolders", found[0]["path"])]
    if default.get("id") is None:
        queue = queue[:1]
    requests = 0
    while queue:
        if requests >= MAX_FOLDER_REQUESTS:
            return found, False
        path, prefix = queue.pop(0)
        nxt: str | None = path
        while nxt:
            requests += 1
            page = await c.graph.page(nxt, params={"$select": "id,displayName", "$top": "100"} if nxt == path else None)
            for f in page.items:
                if f.get("id") in seen:
                    continue
                seen.add(f.get("id"))
                full = f"{prefix}/{f['displayName']}" if prefix else f["displayName"]
                found.append({**f, "path": full})
                queue.append((f"/me/contactFolders/{f['id']}/childFolders", full))
            nxt = page.next_link
            if nxt and requests >= MAX_FOLDER_REQUESTS:
                return found, False
    return found, True


async def _personal(c: Call, id: str, select: str) -> dict[str, Any]:
    """The personal contact behind `id`, refused when it is someone else or already deleted.

    A directory person has no id under /me/contacts, and an address is no id at all. A deleted
    contact keeps its id and answers from Deleted Items, where PATCH and a second DELETE both
    succeed (measured) — the second taking it out of Deleted Items, beyond the person's reach. It is
    recognised as a contact in no contact folder."""
    if "@" in id:
        raise invalid(NOT_A_CONTACT)
    try:
        current = await c.graph.get(f"/me/contacts/{id}", params={"$select": f"{select},parentFolderId"})
    except GraphError as e:
        if e.status in (400, 404):
            raise invalid(NOT_A_CONTACT) from None
        raise
    folders, whole = await _contact_folders(c)
    parent = current.get("parentFolderId")
    if not parent or parent not in {f.get("id") for f in folders}:
        if not whole or not parent:
            # In none of the folders read, and not all were: it may sit in Deleted Items, where a
            # second DELETE takes it beyond reach. Only a contact placed in a contact folder is changed.
            raise invalid(UNPLACED)
        raise invalid(DELETED)
    return current


_HELD = "displayName,emailAddresses,businessPhones,homePhones,mobilePhone"


async def _saved(c: Call, id: str | None, answer: dict[str, Any] | None, not_removed: list[str]) -> dict[str, Any]:
    """What the contact holds now, as Microsoft holds it: the write's own answer, or a read after it when
    the answer does not carry the lists. Built from the request instead, a result would say what was
    asked and not what was kept (Outlook drops, reorders and rewrites values)."""
    held = answer or {}
    id = id or held.get("id")
    if id and not all(k in held for k in ("emailAddresses", "businessPhones")):
        held = await c.graph.get(f"/me/contacts/{id}", params={"$select": _HELD})
    out: dict[str, Any] = {
        "saved": True, "id": id, "name": held.get("displayName"),
        "emails": [e["address"] for e in held.get("emailAddresses") or [] if e.get("address")],
        "phones": _phones_of(held)}
    if not_removed:
        out["not_removed"] = not_removed  # named: a value that matched nothing removed nothing
    return out


@tool(area="contacts", title="Create or change a contact", scopes=["Contacts.ReadWrite"])
async def save_contact(
    name: Annotated[str, Field(description="Full name; required for a new contact.")] = "",
    id: Annotated[str, Field(description="A contact_id from list_people, to change it; omit to create.")] = "",
    emails: Annotated[list[str] | None, Field(description="To add; a contact holds at most three.")] = None,
    phones: Annotated[list[str] | None, Field(description="To add.")] = None,
    remove_emails: Annotated[list[str] | None, Field(description="Addresses to take off the contact; with emails this replaces one.")] = None,
    remove_phones: Annotated[list[str] | None, Field(description="Numbers to take off the contact; with phones this replaces one.")] = None,
    company: Annotated[str, Field(description="Employer.")] = "",
    job_title: Annotated[str, Field(description="Role at the company.")] = "",
    notes: str | None = None,
    folder: Annotated[str | None, Field(description="A contact folder by name, for a new contact; omit for Contacts.")] = None,
) -> dict[str, Any]:
    """Create a personal Outlook contact or change one. Emails and phones are added to those it holds;
    to change a number or address, give the new one in phones or emails and the old one in remove_phones
    or remove_emails. The result lists the emails and phones the contact holds now. Colleagues from the
    directory are not contacts and cannot be changed here."""
    if not id and not (name and name.strip()):
        raise invalid("A new contact needs a name.")
    if id and folder:
        raise invalid("folder places a new contact; an existing one is moved in Outlook. Nothing was changed.")
    if not id and (remove_emails or remove_phones):
        raise invalid("remove_emails and remove_phones take values off an existing contact; give its id. "
                      "Nothing was changed.")
    bad = [e for e in emails or [] if e.strip() and not _ADDRESS.match(e.strip())]
    if bad:
        raise invalid(f"{len(bad)} of the emails is not an email address (name@domain). Nothing was changed.")
    async with call("save_contact") as c:
        c.need("Contacts.ReadWrite")
        current: dict[str, Any] = {}
        if id:
            current = await _personal(c, id, "displayName,emailAddresses,businessPhones,homePhones,mobilePhone")
        fields: dict[str, Any] = {}
        if name and name.strip():
            parts = name.split()
            fields.update({"displayName": name.strip(), "givenName": " ".join(parts[:-1]) or parts[0],
                           "surname": parts[-1] if len(parts) > 1 else ""})
        held_emails = current.get("emailAddresses") or []
        not_removed: list[str] = []
        if remove_emails:
            gone = {e.strip().casefold() for e in remove_emails if e.strip()}
            not_removed += [e for e in remove_emails if e.strip() and e.strip().casefold() not in
                            {(h.get("address") or "").casefold() for h in held_emails}]
            held_emails = [h for h in held_emails if (h.get("address") or "").casefold() not in gone]
        if emails or remove_emails:
            merged = _merge_emails(held_emails, emails or [], name or current.get("displayName") or "")
            if len(merged) > MAX_EMAILS:
                raise invalid(f"Outlook keeps at most three email addresses on a contact; this would make {len(merged)}. "
                              "Save the others on a second contact, or take some off with remove_emails. "
                              "Nothing was changed.")
            fields["emailAddresses"] = merged
        phones_now = _phones_of(current)
        if phones or remove_phones:
            business = list(current.get("businessPhones") or [])
            home = list(current.get("homePhones") or [])
            mobile = current.get("mobilePhone")
            if remove_phones:
                gone = {_phone_key(p) for p in remove_phones if p.strip()}
                not_removed += [p for p in remove_phones if p.strip() and _phone_key(p) not in {_phone_key(h) for h in phones_now}]
                if any(_phone_key(p) in gone for p in home):
                    fields["homePhones"] = home = [p for p in home if _phone_key(p) not in gone]
                if mobile and _phone_key(mobile) in gone:
                    fields["mobilePhone"], mobile = None, None
                business = [p for p in business if _phone_key(p) not in gone]
            kept = {_phone_key(p) for p in [*business, *home, *([mobile] if mobile else [])]}
            for p in phones or []:
                if p.strip() and _phone_key(p) not in kept:
                    kept.add(_phone_key(p))  # one number given twice is added once
                    business.append(p.strip())
            fields["businessPhones"] = business
        if company:
            fields["companyName"] = company
        if job_title:
            fields["jobTitle"] = job_title
        if notes is not None:
            fields["personalNotes"] = notes
        if id:
            return await _saved(c, id, await c.graph.patch(f"/me/contacts/{id}", fields), not_removed)
        path = "/me/contacts"
        if folder:
            folders, whole = await _contact_folders(c)
            if not whole:
                raise invalid("There are more contact folders than can be read here, so the name cannot be placed "
                              "for certain. Omit folder, and move the contact in Outlook. Nothing was changed.")
            key = folder.strip().strip("/")
            target = one(key, folders, kind="contact folder", label=lambda f: f["path"],
                         name=(lambda f: f["path"]) if "/" in key else (lambda f: f["displayName"]))
            path = f"/me/contactFolders/{target['id']}/contacts"
        return await _saved(c, None, await c.graph.post(path, fields), not_removed)


@tool(area="contacts", title="Delete a contact", scopes=["Contacts.ReadWrite"], irreversible=True)
async def delete_contact(
    id: Annotated[str, Field(description="A contact_id from list_people.")],
    user_confirmed: Annotated[bool, Field(description="True only when the person asked.")] = False,
) -> dict[str, Any]:
    """Remove a single personal contact from the person's Outlook address book. It cannot be undone
    from here: Outlook moves it to Deleted Items, where only the person can restore it, in Outlook. Ask the person first, then set user_confirmed."""
    require(user_confirmed, "Deleting a contact")
    async with call("delete_contact") as c:
        c.need("Contacts.ReadWrite")
        await _personal(c, id, "id")
        await c.graph.delete(f"/me/contacts/{id}")
        return {"deleted": True, "id": id, "note": "Moved to Deleted Items; the person can restore it in Outlook."}
