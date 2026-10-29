"""Names to ids. The model says "Anna" or "Inbox/Projects"; the server finds the one it means
or refuses with the candidates. It never picks the first of several."""

from __future__ import annotations

import re
import time
from collections.abc import Callable, Iterable
from dataclasses import dataclass, replace
from typing import Any, TypeVar

from m365_mcp.core.context import Call, scope_granted
from m365_mcp.core.errors import M365Error, invalid
from m365_mcp.graph.errors import GraphError

T = TypeVar("T")
_EMAIL = re.compile(r"^[^@\s<>]+@[^@\s<>]+\.[^@\s<>]+$")
_ANGLE = re.compile(r"<([^<>@\s]+@[^<>\s]+)>")


def one(query: str, items: Iterable[T], *, kind: str, name: Callable[[T], str],
        label: Callable[[T], str] | None = None) -> T:
    """Exact (case-insensitive) match wins; otherwise a single partial match; otherwise refuse."""
    items = list(items)
    q = query.strip().casefold()
    exact = [i for i in items if name(i).casefold() == q]
    if len(exact) == 1:
        return exact[0]
    pool = exact or [i for i in items if q in name(i).casefold()]
    if len(pool) == 1:
        return pool[0]
    if not pool:
        raise M365Error("unresolved", kind=kind, query=query)
    lab = label or name
    raise M365Error("ambiguous", kind=kind, query=query, candidates="; ".join(lab(i) for i in pool[:5]))


@dataclass(frozen=True)
class Who:
    name: str
    address: str
    is_group: bool = False
    job_title: str | None = None
    department: str | None = None
    source: str = "directory"
    contact_id: str | None = None  # a personal contact's own id, for save_contact and delete_contact

    def __str__(self) -> str:
        return f"{self.name} <{self.address}>" if self.name else self.address

    def recipient(self) -> dict[str, Any]:
        return {"emailAddress": {"address": self.address, "name": self.name or self.address}}


def _who_from_person(p: dict[str, Any]) -> Who | None:
    emails = p.get("scoredEmailAddresses") or []
    addr = emails[0].get("address") if emails else p.get("userPrincipalName")
    if not addr:
        return None
    ptype = (p.get("personType") or {}).get("class", "")
    sub = (p.get("personType") or {}).get("subclass", "")
    return Who(name=p.get("displayName") or "", address=addr, is_group=ptype == "Group",
               job_title=p.get("jobTitle"), department=p.get("department"),
               source="contacts" if sub == "PersonalContact" else "directory")


async def people(c: Call, query: str, *, limit: int = 10, directory: bool = True) -> list[Who]:
    """People relevant to the person: their contacts, colleagues and groups, ranked by Microsoft."""
    q = query.replace('"', "")
    found: list[Who] = []
    body = await c.graph.get("/me/people", params={"$search": f'"{q}"', "$top": str(limit),
                                                   "$select": "displayName,scoredEmailAddresses,userPrincipalName,personType,jobTitle,department"})
    for p in body.get("value", []):
        w = _who_from_person(p)
        if w:
            found.append(w)
    if not found and directory:
        body = await c.graph.get("/users", params={"$search": f'"displayName:{q}" OR "mail:{q}"', "$top": str(limit),
                                                   "$select": "displayName,mail,userPrincipalName,jobTitle,department"},
                                 headers={"ConsistencyLevel": "eventual"})
        for u in body.get("value", []):
            addr = u.get("mail") or u.get("userPrincipalName")
            if addr:
                found.append(Who(u.get("displayName") or "", addr, False, u.get("jobTitle"), u.get("department")))
    return found


MAX_CONTACT_PAGES = 8  # pages of contacts and folder lists together; past that the People index is the only source


def _words_match(query: str, *fields: str | None) -> bool:
    """Every word of the query appears in one of the fields, so "Kim Anna" finds "Anna Kim"."""
    hay = " ".join(f for f in fields if f).casefold()
    words = query.casefold().split()
    return bool(words) and all(w in hay for w in words)


class Contacts(list):
    """Matching contacts; `complete` is False when the person has more than the pages read."""
    complete = True


async def contacts(c: Call, query: str, *, names_only: bool = False) -> Contacts:
    """The person's own Outlook contacts that match, read from where they are kept: the People index
    learns a new contact late or never. Contacts without an address are not someone to write to."""
    out = Contacts()
    if not scope_granted("Contacts.Read", c.person.scopes):
        return out
    q = query.strip()
    select = {"$top": "100", "$select": "id,displayName,emailAddresses,companyName,jobTitle,department"}
    budget = [MAX_CONTACT_PAGES]  # pages, shared by every folder

    def take(items: list[dict[str, Any]]) -> None:
        for ct in items:
            emails = [e.get("address") for e in ct.get("emailAddresses") or [] if e.get("address")]
            if not emails:
                continue
            fields = [ct.get("displayName")] if names_only else [ct.get("displayName"), ct.get("companyName"),
                                                                 ct.get("jobTitle"), *emails]
            exact = next((e for e in emails if e.casefold() == q.casefold()), None)
            if exact or _words_match(q, *fields):
                out.append(Who(name=ct.get("displayName") or "", address=exact or emails[0], job_title=ct.get("jobTitle"),
                               department=ct.get("department"), source="contacts", contact_id=ct.get("id")))

    async def read_all(path: str, params: dict[str, str] | None = None) -> tuple[list[dict[str, Any]], bool]:
        """The items of a collection that the page budget allows, and whether that was all of them."""
        items: list[dict[str, Any]] = []
        nxt: str | None = path
        while nxt:
            if budget[0] == 0:
                return items, False
            budget[0] -= 1
            page = await c.graph.page(nxt, params=params if nxt == path else None)
            items += page.items
            nxt = page.next_link
        return items, True

    # The default folder, then every other contact folder and the folders inside them: a second
    # person of one name in "Clients" makes the name as ambiguous as one in the default folder.
    folders = ["/me"]
    while folders:
        base = folders.pop(0)
        got, whole = await read_all(f"{base}/contacts", select)
        take(got)  # what was read still counts when the budget ran out
        if not whole:
            out.complete = False
            return out
        children, whole = await read_all("/me/contactFolders" if base == "/me" else f"{base}/childFolders",
                                         {"$top": "100", "$select": "id"})
        if not whole:
            out.complete = False
            return out
        folders += [f"/me/contactFolders/{f['id']}" for f in children if f.get("id")]
    return out


DOMAIN_TTL = 3600.0


async def _organisation_domains(c: Call) -> frozenset[str]:
    hit = c.rt.domain_cache.get(c.person.tenant)
    if hit and hit[1] > time.monotonic():
        return hit[0]
    body = await c.graph.get("/organization", params={"$select": "verifiedDomains"})
    names = frozenset(d["name"].casefold() for org in body.get("value", [])
                      for d in org.get("verifiedDomains") or [] if d.get("name"))
    c.rt.domain_cache[c.person.tenant] = (names, time.monotonic() + DOMAIN_TTL)
    return names


async def outside_organisation(c: Call, addresses: list[str]) -> list[str] | None:
    """The addresses that are not at one of the organisation's own domains, in the order given; None when
    the domains cannot be read, which the caller says rather than reporting none. A draft already
    written is not failed over a courtesy check."""
    try:
        own = await _organisation_domains(c)
    except GraphError:
        return None
    if not own:
        return None
    out = []
    for a in dict.fromkeys(addresses):
        domain = a.rpartition("@")[2].casefold().rstrip(".")
        if not any(domain == d or domain.endswith("." + d) for d in own):
            out.append(a)
    return out


async def _directory_user(c: Call, address: str) -> dict[str, Any] | None:
    a = address.replace("'", "''")
    body = await c.graph.get("/users", params={
        "$filter": f"mail eq '{a}' or userPrincipalName eq '{a}'", "$top": "1",
        "$select": "displayName,mail,userPrincipalName,jobTitle,department"})
    users = body.get("value") or []
    return users[0] if users else None


async def classify(c: Call, address: str, known: Who | None = None) -> Who:
    """What an address reaches. An address at one of the organisation's own domains must be a user in
    its directory: anything else there is a Microsoft 365 group, a distribution list, a mail-enabled
    security group, an alias, or nothing at all, and none of those is a person to write to (rule 22).
    Groups are not read directly: that needs Group.Read.All, which no tool needs otherwise (rule 18).
    An address outside the organisation is taken as given; whether it is a list there cannot be known
    from here, and an installation that must limit it sets M365_ALLOWED_RECIPIENT_DOMAINS."""
    if known and known.is_group:
        return known
    user = await _directory_user(c, address)
    if user:
        return Who(name=user.get("displayName") or (known.name if known else ""),
                   address=user.get("mail") or address,
                   job_title=user.get("jobTitle") or (known.job_title if known else None),
                   department=user.get("department") or (known.department if known else None),
                   source=known.source if known else "directory")
    domain = address.rpartition("@")[2].casefold().rstrip(".")
    if any(domain == d or domain.endswith("." + d) for d in await _organisation_domains(c)):
        return Who(name=known.name if known else "", address=address, is_group=True, source="directory")
    return known or Who(name="", address=address, source="address")


def _label(w: Who) -> str:
    return f"{w} ({w.job_title})" if w.job_title else str(w)


async def person(c: Call, query: str) -> Who:
    """One person for a name or an address. A name that several people match is refused with all of
    them, even when one display name equals it exactly: "Megan" is also the first name of "Megan
    Bowen", and the person must say which (rule 7). An address is always exact."""
    q = query.strip()
    m = _ANGLE.search(q)
    if m:
        q = m.group(1)
    if not q:
        raise invalid("A person is needed: a name or an address.")
    if _EMAIL.match(q):
        hits = await people(c, q, limit=5, directory=False)
        known = next((h for h in hits if h.address.casefold() == q.casefold()), None)
        return await classify(c, q, known)
    own = await contacts(c, q, names_only=True)
    if not own.complete:
        # A second match past the pages read would make this name ambiguous without anyone knowing.
        raise invalid(f"There are too many personal contacts or contact folders to read them all here (the "
                      f"limit is {MAX_CONTACT_PAGES} requests), so a name cannot be checked against all of them. "
                      "Give the person's address instead.")
    ranked = await people(c, q)
    # People's word that an address is a group outlives a contact holding it, whatever either is called.
    groups = {w.address.casefold() for w in ranked if w.is_group}
    pool: dict[str, Who] = {}
    for w in [*own, *ranked]:
        key = w.address.casefold()
        if not _words_match(q, w.name) or key in pool:
            continue
        pool[key] = replace(w, is_group=True) if key in groups and not w.is_group else w
    found = list(pool.values())
    if not found:
        raise M365Error("unresolved", kind="person", query=q)
    if len(found) > 1:
        raise M365Error("ambiguous", kind="person", query=q, candidates="; ".join(_label(w) for w in found[:5]))
    w = found[0]
    # A personal contact can hold any address, a group's included; People's own groups are marked.
    return await classify(c, w.address, w) if w.source == "contacts" else w


async def address_of(c: Call, query: str) -> str:
    """An address as given (no lookup), or the one person a name resolves to; ambiguity refuses."""
    q = query.strip()
    m = _ANGLE.search(q)
    if m:
        q = m.group(1)
    return q if _EMAIL.match(q) else (await person(c, q)).address


async def recipients(c: Call, names: list[str] | None) -> list[Who]:
    return [await person(c, n) for n in (names or [])]


def groups_among(ws: list[Who]) -> list[str]:
    return [w.address for w in ws if w.is_group]


# -- mail folders -------------------------------------------------------------------------

WELL_KNOWN = {
    "inbox": "inbox", "sent": "sentitems", "sent items": "sentitems", "drafts": "drafts",
    "deleted": "deleteditems", "deleted items": "deleteditems", "trash": "deleteditems",
    "archive": "archive", "junk": "junkemail", "spam": "junkemail", "outbox": "outbox",
}
SYSTEM_FOLDER_IDS = set(WELL_KNOWN.values())


async def _all_folders(c: Call) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []

    async def walk(path: str, prefix: str, depth: int) -> None:
        page = await c.graph.page(path, params={"$top": "100", "$select": "id,displayName,totalItemCount,unreadItemCount,childFolderCount"})
        items = list(page.items)
        while page.next_link and len(items) < 1000:
            page = await c.graph.page(page.next_link)
            items += page.items
        for f in items:
            full = f"{prefix}/{f['displayName']}" if prefix else f["displayName"]
            f["path"] = full
            out.append(f)
            if f.get("childFolderCount") and depth < 6:
                await walk(f"/me/mailFolders/{f['id']}/childFolders", full, depth + 1)

    await walk("/me/mailFolders", "", 0)
    return out


async def mail_folder(c: Call, name: str) -> dict[str, Any]:
    key = name.strip().strip("/")
    wk = WELL_KNOWN.get(key.casefold())
    if wk:
        f = await c.graph.get(f"/me/mailFolders/{wk}", params={"$select": "id,displayName"})
        f["well_known"] = wk
        f["path"] = f.get("displayName", key)
        return f
    folders = await _all_folders(c)
    if "/" in key:
        return one(key, folders, kind="folder", name=lambda f: f["path"])
    return one(key, folders, kind="folder", name=lambda f: f["displayName"], label=lambda f: f["path"])
