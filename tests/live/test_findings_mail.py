"""Probe for mail: what Graph's attachment `size` measures. Records the answer as a finding."""

from __future__ import annotations

import pytest

import builders
from test_live import find_message

pytestmark = pytest.mark.live


async def test_attachment_size_against_real_bytes(live):
    b = await live.address("b")
    subject = f"{live.run} size probe"
    draft = await live.ok("a", "save_draft", to=[b], subject=subject, body="size probe", attachments_upload=3)
    sizes = [55, 5000, 200_000]
    names = [f"{live.run}-{n}.bin" for n in sizes]
    for link, n, name in zip(draft["attachment_uploads"], sizes, names, strict=True):
        put = await live.upload(live.link_in(link), builders.random_bytes(n),
                                {"Content-Type": "application/octet-stream", "X-File-Name": name})
        assert put.status_code == 200, put.text[:300]
    on_draft = await live.ok("a", "get_message", id=draft["draft_id"], include_attachments=True)
    await live.ok("a", "send_draft", id=draft["draft_id"], user_confirmed=True)
    arrived = await find_message(live, "b", live.run, subject_starts=subject)
    got = await live.ok("b", "get_message", id=arrived["id"], include_attachments=True)
    seen: dict[str, dict] = {}
    for where, listing in (("draft", on_draft["attachments"]), ("received", got["attachments"])):
        for a in listing:
            real = len((await live.fetch(live.link_in(a["download"]))).content)
            seen[f"{where} {real}"] = a.get("size_incl_overhead")
    live.finding("attachment_size_vs_real_bytes", seen)


async def test_messages_exclude_deleted_items_with_a_filter_beside_orderby(live):
    """Does /me/messages honour `parentFolderId ne <Deleted Items id>` next to $orderby, with immutable ids?
    list_messages relies on it for a listing without a folder; $search cannot take any $filter."""
    folder = await live.raw("a", "GET", "/me/mailFolders/deleteditems", params={"$select": "id"})
    assert folder.status_code == 200, folder.text[:300]
    fid = folder.json()["id"]
    params = {"$top": "25", "$select": "id,parentFolderId", "$orderby": "receivedDateTime desc",
              "$filter": f"receivedDateTime ge 1900-01-01T00:00:00Z and parentFolderId ne '{fid}'"}
    got = await live.raw("a", "GET", "/me/messages", params=params)
    unfiltered = await live.raw("a", "GET", "/me/messages", params={"$top": "100", "$select": "parentFolderId"})
    search = await live.raw("a", "GET", "/me/messages", params={"$top": "5", "$search": '"a"', "$filter": f"parentFolderId ne '{fid}'"})
    rows = got.json().get("value", []) if got.status_code == 200 else []
    live.finding("messages_exclude_deleted_items", {
        "filter_with_orderby": observed_status(got, live),
        "deleted_rows_in_filtered_page": sum(1 for m in rows if m.get("parentFolderId") == fid),
        "rows_in_filtered_page": len(rows),
        "deleted_rows_in_unfiltered_100": sum(1 for m in unfiltered.json().get("value", []) if m.get("parentFolderId") == fid)
        if unfiltered.status_code == 200 else None,
        "search_with_filter": observed_status(search, live)})
    assert got.status_code == 200, got.text[:300]


def observed_status(resp, live) -> str:
    return "ok" if resp.status_code < 300 else live.error_of(resp)
