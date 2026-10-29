"""Addressing in OneDrive and SharePoint: drives, items by id or path, sites.

A drive is the person's own (`/me/drive`) unless a drive id is given, which is how a library of
a SharePoint site or another person's shared drive is reached. Ids and paths come from the model,
so they are checked before they go into a URL.
"""

from __future__ import annotations

import re
from typing import Any
from urllib.parse import quote, urlsplit

from m365_mcp.core.context import Call, scope_granted
from m365_mcp.core.errors import M365Error, invalid
from m365_mcp.core.resolve import one
from m365_mcp.graph.errors import GraphError

ITEM_SELECT = ("id,name,size,webUrl,file,folder,package,createdDateTime,createdBy,lastModifiedDateTime,"
               "lastModifiedBy,parentReference")
_BAD_ID = re.compile(r"[\s/\\?#%]")
SITE_SELECT = "id,displayName,webUrl,description"
MAX_TEAMS = 20  # teams whose sites are read; each costs one request
SITE_ID = re.compile(r"^[^,\s/]+,[0-9a-fA-F-]{36},[0-9a-fA-F-]{36}$")
GUID = re.compile(r"^[0-9a-fA-F-]{36}$")


def check_id(value: str | None, what: str = "id") -> str:
    v = (value or "").strip()
    if not v or len(v) > 300 or _BAD_ID.search(v):
        raise invalid(f"The {what} is not valid. Use the {what} from a previous result.")
    return v


def root(drive: str | None) -> str:
    return f"/drives/{check_id(drive, 'drive')}" if drive else "/me/drive"


def item_url(drive: str | None, item_id: str) -> str:
    return f"{root(drive)}/items/{check_id(item_id)}"


def path_url(drive: str | None, path: str) -> str:
    p = path.strip().strip("/")
    if "#" in p:
        # Graph cannot take "#" in a path, even encoded, without opening it to rewriting; the id can.
        raise invalid("A path with # in it cannot be used. Use the item's id from list_files instead.")
    return f"{root(drive)}/root:/{quote(p)}" if p else f"{root(drive)}/root"


async def find(c: Call, item_id: str | None, path: str | None, drive: str | None,
               select: str = ITEM_SELECT) -> dict[str, Any]:
    """One drive item by id, or by path in the drive. Returns Graph's object."""
    if item_id:
        return await c.graph.get(item_url(drive, item_id), params={"$select": select})
    if path and path.strip().strip("/"):
        return await c.graph.get(path_url(drive, path), params={"$select": select})
    raise invalid("Give the file's id (from list_files) or its path.")


async def folder(c: Call, path: str | None, drive: str | None) -> dict[str, Any]:
    """A folder by path; an empty path is the drive's root."""
    f = await c.graph.get(path_url(drive, path or ""), params={"$select": "id,name,folder,parentReference"})
    if "folder" not in f and (path or "").strip().strip("/"):
        raise invalid("That path is a file, not a folder.")
    return f


async def own_drive(c: Call) -> str | None:
    body = await c.graph.get("/me/drive", params={"$select": "id"})
    return body.get("id")


def extension(name: str | None) -> str:
    n = (name or "").rsplit(".", 1)
    return n[1].lower() if len(n) == 2 else ""


async def known_sites(c: Call, cut: list[bool] | None = None) -> list[dict[str, Any]]:
    """The sites the person follows and the sites of their teams. Search leaves out a team's own site
    until it has indexed it, and followedSites only holds what was followed (both measured), so the
    teams are walked too, when Teams reading is granted. A source that refuses (403, 404) is passed
    over; throttling and outages fail the call. `cut` records that some were not read."""
    out: list[dict[str, Any]] = []
    try:
        body = await c.graph.get("/me/followedSites", params={"$select": SITE_SELECT})
        out += body.get("value", [])
        if body.get("@odata.nextLink") and cut is not None:
            cut.append(True)
    except GraphError as ge:
        if ge.status not in (403, 404):
            raise
    if scope_granted("Team.ReadBasic.All", c.person.scopes):
        try:
            teams = (await c.graph.get("/me/joinedTeams", params={"$select": "id"})).get("value", [])
        except GraphError as ge:
            if ge.status not in (403, 404):
                raise
            teams = []
        if len(teams) > MAX_TEAMS and cut is not None:
            cut.append(True)
        for t in teams[:MAX_TEAMS]:
            try:
                out.append(await c.graph.get(f"/groups/{t['id']}/sites/root", params={"$select": SITE_SELECT}))
            except GraphError as ge:
                if ge.status not in (403, 404):
                    raise
    return list({s["id"]: s for s in out if s.get("id")}.values())


async def search_sites(c: Call, query: str, top: int) -> tuple[list[dict[str, Any]], str | None, bool]:
    """Microsoft Search for sites: hits, the next link, and whether search answered at all."""
    try:
        body = await c.graph.get("/sites", params={"search": query, "$top": str(top), "$select": SITE_SELECT})
    except GraphError as ge:
        if ge.status < 500:
            raise
        return [], None, False  # measured: a tenant whose search is not set up answers 500 here
    return body.get("value", []), body.get("@odata.nextLink"), True


async def site(c: Call, query: str) -> dict[str, Any]:
    """A SharePoint site by name, URL or id. Several matches are refused with their names."""
    q = (query or "").strip()
    if not q:
        raise invalid("A site is needed: its name or its address.")
    if SITE_ID.match(q):
        return {"id": q, "displayName": q}
    if q.startswith("https://"):
        parts = urlsplit(q)
        return await c.graph.get(f"/sites/{parts.hostname}:{parts.path.rstrip('/')}",
                                 params={"$select": "id,displayName,webUrl"})
    known = [s for s in await known_sites(c) if q.casefold() in (s.get("displayName") or "").casefold()]
    hits = [s for s in known if (s.get("displayName") or "").casefold() == q.casefold()]
    if not hits:
        # A partial match among the person's own sites is not enough: Search may hold the exact one.
        found, _, answered = await search_sites(c, q, 25)
        hits = list({s["id"]: s for s in [*known, *found] if s.get("id")}.values())
        if not answered and not known:
            # Not "no site matches": Search did not answer, so it is not known whether one does.
            raise M365Error("unsupported", reason=(
                "No site you follow or of your teams matches, and Microsoft Search did not answer, so other sites "
                "cannot be looked up now. Give the site's address instead."))
    if len(hits) == 1:
        return hits[0]
    return one(q, hits, kind="site", name=lambda s: s.get("displayName") or "",
               label=lambda s: f"{s.get('displayName')} ({s.get('webUrl')})")
