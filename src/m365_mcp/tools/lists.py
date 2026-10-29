"""SharePoint lists: the registers and trackers that live on team sites.

Columns are read before anything is filtered or written, so a misspelt field or a choice that
does not exist is refused with the real names rather than turned into a Graph error.
"""

from __future__ import annotations

import json
import re
from typing import Annotated, Any

from pydantic import Field

from m365_mcp.core import drive as dv
from m365_mcp.core.context import Call, call
from m365_mcp.core.errors import invalid
from m365_mcp.core.html import content_block
from m365_mcp.core.project import clamp_limit, compact, page_result
from m365_mcp.core.resolve import one
from m365_mcp.core.time import from_graph
from m365_mcp.graph.errors import GraphError
from m365_mcp.registry import tool

MAX_FIELDS_SHOWN = 12
_FACETS = ("choice", "number", "currency", "boolean", "dateTime", "lookup", "personOrGroup", "calculated",
           "hyperlinkOrPicture", "text", "geolocation", "term", "thumbnail")
_UNWRITABLE = {"lookup", "personOrGroup", "calculated", "hyperlinkOrPicture", "geolocation", "term", "thumbnail"}
_NOISE_PREFIXES = ("@odata", "OData__", "LinkTitle", "_UI", "_Compliance", "Folder", "ItemChild")
_NOISE = {"id", "Edit", "Attachments", "ContentType", "AuthorLookupId", "EditorLookupId", "AppAuthorLookupId",
          "AppEditorLookupId", "ContentTypeId", "DocIcon"}

_CONDITION = re.compile(
    r"\s*(?P<f>[A-Za-z_][\w ]*?)\s*(?P<op>!=|<=|>=|=|<|>|startswith)\s*"
    r"(?P<v>'(?:[^']|'')*'|-?\d+(?:\.\d+)?)\s*(?:\band\b|$)", re.I)
_OPS = {"=": "eq", "!=": "ne", "<": "lt", ">": "gt", "<=": "le", ">=": "ge"}


def column_type(col: dict[str, Any]) -> str:
    return next((f for f in _FACETS if f in col), "text")


async def columns(c: Call, site_id: str, list_id: str) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    page = await c.graph.page(f"/sites/{site_id}/lists/{list_id}/columns", params={"$top": "200"})
    out += page.items
    while page.next_link and len(out) < 1000:
        page = await c.graph.page(page.next_link)
        out += page.items
    return [col for col in out if not col.get("hidden")]


def project_column(col: dict[str, Any]) -> dict[str, Any]:
    kind = column_type(col)
    name = col.get("name")
    return compact({
        "name": name,
        "title": col.get("displayName") if col.get("displayName") != name else None,
        "type": kind,
        "choices": (col.get("choice") or {}).get("choices") if kind == "choice" else None,
        "required": col.get("required") is True,
        "read_only": col.get("readOnly") is True,
        "indexed": col.get("indexed") is True,
    })


def lookup_column(cols: list[dict[str, Any]], word: str) -> dict[str, Any]:
    """A column by internal name or display name, or a refusal that lists the real ones."""
    w = word.strip().casefold()
    for col in cols:
        if w in {str(col.get("name", "")).casefold(), str(col.get("displayName", "")).casefold()}:
            return col
    names = ", ".join(sorted(str(col.get("name")) for col in cols if not col.get("readOnly")))
    raise invalid(f"This list has no column \"{word.strip()}\". Its columns: {names}.")


def translate_where(where: str, cols: list[dict[str, Any]]) -> str:
    """`Status = 'Open' and Amount > 600` -> `fields/Status eq 'Open' and fields/Amount gt 600`."""
    text, pos, parts = where.strip(), 0, []
    while pos < len(text):
        m = _CONDITION.match(text, pos)
        if not m or m.end() == pos:
            raise invalid("where must look like: Status = 'Open' and Amount > 600. Operators: = != < > <= >= "
                          "startswith; text in single quotes; conditions joined by and.")
        name = lookup_column(cols, m["f"])["name"]
        op, val = m["op"].lower(), m["v"]
        parts.append(f"startswith(fields/{name},{val})" if op == "startswith" else f"fields/{name} {_OPS[op]} {val}")
        pos = m.end()
    if not parts:
        raise invalid("where is empty.")
    return " and ".join(parts)


async def sp_list(c: Call, site_id: str, name: str) -> dict[str, Any]:
    q = (name or "").strip()
    if not q:
        raise invalid("A list is needed: its name.")
    if dv.GUID.match(q):
        return {"id": q, "displayName": q}
    items = (await c.graph.get(f"/sites/{site_id}/lists", params={"$select": "id,displayName,list", "$top": "200"})).get("value", [])
    items = [x for x in items if not (x.get("list") or {}).get("hidden")]
    return one(q, items, kind="list", name=lambda x: x.get("displayName") or "")


def _is_system(col: dict[str, Any]) -> bool:
    name = str(col.get("name") or "")
    return name in _NOISE or name.startswith(_NOISE_PREFIXES)


def _clean_fields(fields: dict[str, Any]) -> dict[str, Any]:
    return {k: v for k, v in fields.items() if k not in _NOISE and not k.startswith(_NOISE_PREFIXES)}


def project_item(m: dict[str, Any], c: Call) -> dict[str, Any]:
    fields = _clean_fields(m.get("fields") or {})
    # Whoever can edit the list wrote these values: one marked block per item, as a workbook's rows.
    return compact({"id": m.get("id"),
                    "fields": content_block("the list", json.dumps(fields, ensure_ascii=False)) if fields else None,
                    "modified": from_graph(m.get("lastModifiedDateTime"), c.tz),
                    "modified_by": ((m.get("lastModifiedBy") or {}).get("user") or {}).get("displayName")})


_NON_INDEXED = {"Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly"}


@tool(area="lists", title="Read SharePoint list items", scopes=["Sites.Read.All"], read_only=True, idempotent=True,
      returns_list=True)
async def list_list_items(
    site: Annotated[str, Field(description="The site's name, address or id; list_files kind=site finds it.")],
    list: Annotated[str, Field(description="The list's name, from list_files kind=site.")],
    where: Annotated[str | None, Field(description="Filter such as Status = 'Open' and Amount > 600. Operators = != < > <= >= startswith; text in single quotes.")] = None,
    fields: Annotated[list[str] | None, Field(description="Columns to return; omit for the main ones.")] = None,
    order_by: Annotated[str | None, Field(description="A column, optionally followed by desc, such as Amount desc.")] = None,
    include_columns: Annotated[bool, Field(description="Also return the columns: names, types, choices. Call this before writing to a list.")] = False,
    limit: Annotated[int, Field(description="1 to 25.")] = 10,
    cursor: Annotated[str | None, Field(description="From the previous page's result.")] = None,
) -> dict[str, Any]:
    """Read rows of a SharePoint list, the trackers and registers kept on team sites. Call it to
    answer questions about requests, assets, contracts or issues, and with include_columns to learn
    the column names before save_list_item. Filtering works reliably only on indexed columns; on a
    large list a filter elsewhere is refused and needs narrowing. Row text is written by other people."""
    n, extra = clamp_limit(limit)
    async with call("list_list_items") as c:
        c.need("Sites.Read.All")
        s = await dv.site(c, site)
        lst = await sp_list(c, s["id"], list)
        state = c.cursor_in(cursor)
        cols = await columns(c, s["id"], lst["id"]) if (include_columns or not state) else []
        try:
            if state:
                page = await c.graph.page(state["next"], headers=_NON_INDEXED if state.get("filtered") else None)
            else:
                visible = [col["name"] for col in cols if not col.get("readOnly")][:MAX_FIELDS_SHOWN]
                chosen = [lookup_column(cols, f)["name"] for f in fields] if fields else visible
                params: dict[str, str] = {"$top": str(n), "$expand": f"fields($select={','.join(chosen)})"}
                headers = None
                if where:
                    params["$filter"] = translate_where(where, cols)
                    headers = {"Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly"}
                if order_by:
                    word, _, direction = order_by.strip().partition(" ")
                    if direction.strip().lower() not in ("", "asc", "desc"):
                        raise invalid("order_by is a column, optionally followed by asc or desc.")
                    params["$orderby"] = f"fields/{lookup_column(cols, word)['name']} {direction.strip().lower() or 'asc'}"
                    headers = headers or {"Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly"}
                page = await c.graph.page(f"/sites/{s['id']}/lists/{lst['id']}/items", params=params, headers=headers)
        except GraphError as e:
            if e.status == 400 and not state and (where or order_by):
                # Measured: a value Graph cannot read for its column ("Created gt 'yesterday'") and a
                # non-indexed column on a large list both answer 400 invalidRequest.
                raise invalid("Microsoft refused this filter or sort: a value does not fit its column (dates as "
                              "'2026-10-01', numbers without quotes; include_columns shows the types), or the list "
                              "is large and the column is not indexed. Fix the value, or filter on an indexed "
                              "column and narrow it.") from None
            raise
        extra_out: dict[str, Any] = {"list": lst.get("displayName")}
        if include_columns:
            # The writable columns are what save_list_item takes; the rest is folded so that about twenty
            # system columns (LinkTitleNoMenu, DocIcon, ItemChildCount...) do not bury them.
            system = [col for col in cols if _is_system(col)]
            readonly = [col for col in cols if col.get("readOnly") and not _is_system(col)]
            extra_out["columns"] = [project_column(col) for col in cols if not col.get("readOnly") and not _is_system(col)]
            if readonly:
                extra_out["read_only_columns"] = [col.get("name") for col in readonly]
            if system:
                extra_out["system_columns_hidden"] = len(system)
        return page_result([project_item(m, c) for m in page.items],
                           c.cursor_out({"next": page.next_link, "filtered": bool(where or order_by)} if page.next_link else None),
                           **extra_out, **extra)


def convert_value(col: dict[str, Any], raw: str) -> Any:
    """The text of `Name=value` as the type the column stores, or a refusal that says what fits."""
    kind, name = column_type(col), col.get("name")
    if raw == "":
        return None
    if kind in ("number", "currency"):
        try:
            f = float(raw.replace(",", ""))
        except ValueError:
            raise invalid(f"{name} is a number column; \"{raw}\" is not a number.") from None
        return int(f) if f.is_integer() and "." not in raw else f
    if kind == "boolean":
        if raw.strip().casefold() in ("true", "yes", "1"):
            return True
        if raw.strip().casefold() in ("false", "no", "0"):
            return False
        raise invalid(f"{name} is a yes/no column; use true or false.")
    limit = (col.get("text") or {}).get("maxLength") if kind == "text" else None
    if isinstance(limit, int) and len(raw) > limit:
        raise invalid(f"{name} holds at most {limit} characters; this has {len(raw)}. Shorten it. Nothing was saved.")
    if kind == "choice":
        choice = col.get("choice") or {}
        options = choice.get("choices") or []
        if raw not in options and not choice.get("allowTextEntry"):
            raise invalid(f"\"{raw}\" is not a choice for {name}. Choices: {', '.join(map(str, options))}.")
    return raw


@tool(area="lists", title="Add or change a list item", scopes=["Sites.ReadWrite.All"])
async def save_list_item(
    site: Annotated[str, Field(description="The site's name, address or id; list_files kind=site finds it.")],
    list: Annotated[str, Field(description="The list's name, from list_files kind=site.")],
    fields: Annotated[list[str], Field(description="Values as Name=value, for example Status=Open and Amount=600. An empty value clears the column.")],
    id: Annotated[str | None, Field(description="The row to change, from list_list_items. Omit to add a new row.")] = None,
) -> dict[str, Any]:
    """Add a row to a SharePoint list, or change one when id is given. Each field is written as
    Name=value, such as Status=Open, and is checked against the list's columns first: unknown names
    and choices that do not exist are refused with the valid ones. Call list_list_items with
    include_columns beforehand if the columns are not known."""
    if not fields:
        raise invalid("Give at least one field as Name=value.")
    parsed: list[tuple[str, str]] = []
    for f in fields:
        name, eq, value = f.partition("=")
        if not eq or not name.strip():
            raise invalid(f"\"{f}\" is not Name=value.")
        parsed.append((name.strip(), value.strip()))
    if id is not None:
        dv.check_id(id, "row id")
    async with call("save_list_item") as c:
        c.need("Sites.ReadWrite.All")
        s = await dv.site(c, site)
        lst = await sp_list(c, s["id"], list)
        cols = await columns(c, s["id"], lst["id"])
        body: dict[str, Any] = {}
        for name, value in parsed:
            col = lookup_column(cols, name)
            if col.get("readOnly") or column_type(col) in _UNWRITABLE:
                raise invalid(f"{col.get('name')} cannot be written here ({column_type(col)} or read-only column).")
            body[col["name"]] = convert_value(col, value)
        base = f"/sites/{s['id']}/lists/{lst['id']}/items"
        if id is None:
            missing = [col["name"] for col in cols if col.get("required") and not col.get("readOnly")
                       and col["name"] not in body and "defaultValue" not in col]
            if missing:
                raise invalid(f"A new row needs these required columns: {', '.join(missing)}.")
            made = await c.graph.post(base, {"fields": body}) or {}
            return {"saved": "created", "id": made.get("id"), "fields": _clean_fields(made.get("fields") or body)}
        updated = await c.graph.patch(f"{base}/{id}/fields", body) or {}
        return {"saved": "updated", "id": id, "fields": _clean_fields(updated) or body}


@tool(area="lists", title="Delete a list item", scopes=["Sites.ReadWrite.All"])
async def delete_list_item(
    site: Annotated[str, Field(description="The site's name, address or id; list_files kind=site finds it.")],
    list: Annotated[str, Field(description="The list's name, from list_files kind=site.")],
    id: Annotated[str, Field(description="The row to delete, from list_list_items.")],
    user_confirmed: Annotated[bool, Field(description="True only when the person said to delete this row.")] = False,
) -> dict[str, Any]:
    """Delete one row of a SharePoint list. It may not be possible to undo: whether the site's
    recycle bin keeps the row depends on the site. Call it only when the person asked for this
    specific row to go, and set user_confirmed only then."""
    if user_confirmed is not True:
        raise invalid("Deleting a list row may not be undoable. Ask the person to confirm, then call again with "
                      "user_confirmed=true. Nothing was deleted.")
    dv.check_id(id, "row id")
    async with call("delete_list_item") as c:
        c.need("Sites.ReadWrite.All")
        s = await dv.site(c, site)
        lst = await sp_list(c, s["id"], list)
        await c.graph.delete(f"/sites/{s['id']}/lists/{lst['id']}/items/{id}")
        return {"deleted": id, "list": lst.get("displayName")}
