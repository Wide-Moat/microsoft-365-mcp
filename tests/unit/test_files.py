import re

import httpx
import pytest

from m365_mcp.links.tokens import LinkRefused

OWN = "b!own-drive"
SITE_DRIVE = "b!site-drive"
DELETE_OK = re.compile(r"^/(me/drive|drives/[^/]+)/items/[^/]+(/permissions/[^/]+)?$")


def drive_item(i, name, *, folder=False, drive=OWN, path="/drive/root:/Reports", size=1234, **kw):
    m = {"id": i, "name": name, "size": size, "webUrl": f"https://x.test/{name}",
         "lastModifiedDateTime": "2026-10-20T10:00:00Z",
         "lastModifiedBy": {"user": {"displayName": "Bo"}},
         "parentReference": {"driveId": drive, "path": path}, "@odata.etag": "W/1", "eTag": "abc", **kw}
    m["folder" if folder else "file"] = {"childCount": 3} if folder else {"mimeType": "application/octet-stream"}
    return m


def person(graph, name, addr, group=False):
    def handler(seen):
        q = seen.params.get("$search", "").strip('"').lower()
        if q and (q in name.lower() or q in addr.lower()):
            return httpx.Response(200, json={"value": [{
                "displayName": name, "scoredEmailAddresses": [{"address": addr}],
                "personType": {"class": "Group" if group else "Person", "subclass": "OrganizationUser"}}]})
        return httpx.Response(200, json={"value": []})
    graph.on("GET", "/me/people", handler)


def link_of(text, op):
    m = re.search(rf"https://links\.test/{op}/([\w.\-]+)", text)
    assert m, text
    return m.group(1)


def quiet(graph):
    """Requests other than the time-zone lookup every call makes once."""
    return [s for s in graph.seen if "mailboxSettings" not in s.path]


def assert_deletes_stay_in_the_allowlist(graph):
    for s in graph.seen:
        if s.method == "DELETE":
            assert DELETE_OK.match(s.path), s.path


# -- list_files ------------------------------------------------------------------------------

async def test_list_files_my_projects_items(graph, invoke):
    graph.on("GET", "/me/drive/root/children", {"value": [
        drive_item("f1", "Q3 plan.docx"), drive_item("d1", "Archive", folder=True),
        drive_item("f2", "chart.png"), drive_item("f3", "report.pdf")]})
    r = await invoke("list_files")
    assert not r.error, r.text
    names = {i["name"]: i for i in r.data["items"]}
    assert names["Q3 plan.docx"]["kind"] == "document" and names["Archive"]["kind"] == "folder"
    assert names["chart.png"]["kind"] == "image" and names["report.pdf"]["kind"] == "pdf"
    assert names["Q3 plan.docx"]["path"] == "Reports/Q3 plan.docx"
    assert names["Q3 plan.docx"]["modified_by"] == "Bo" and "drive" not in names["Q3 plan.docx"]
    assert "size" not in names["Archive"] and names["Archive"]["children"] == 3
    assert "@odata" not in r.text and "eTag" not in r.text
    assert graph.calls("GET", "/me/drive/root/children")[0].params["$top"] == "10"
    assert graph.writes() == []


async def test_list_files_in_a_folder_uses_path_addressing(graph, invoke):
    graph.on("GET", "/me/drive/root:/Reports/2026:/children", {"value": []})
    r = await invoke("list_files", folder="Reports/2026")
    assert not r.error, r.text


async def test_query_without_folder_searches_everything_and_pages_by_from(graph, invoke):
    graph.on("GET", "/me/drive", {"id": OWN})
    hit = lambda i: {"resource": drive_item(i, f"{i}.docx", drive=SITE_DRIVE)}  # noqa: E731
    graph.on("POST", "/search/query", {"value": [{"hitsContainers": [{
        "hits": [hit("a"), hit("b")], "moreResultsAvailable": True}]}]})
    r = await invoke("list_files", query="budget", limit=2)
    assert not r.error, r.text
    body = graph.calls("POST", "/search/query")[0].json
    assert body == {"requests": [{"entityTypes": ["driveItem"], "query": {"queryString": "budget"}, "from": 0, "size": 2}]}
    assert r.data["items"][0]["drive"] == SITE_DRIVE and r.data["has_more"]
    await invoke("list_files", query="budget", limit=2, cursor=r.data["cursor"])
    assert graph.calls("POST", "/search/query")[1].json["requests"][0]["from"] == 2


async def test_shared_goes_through_search_never_the_deprecated_endpoint(graph, invoke):
    graph.on("GET", "/me/drive", {"id": OWN})
    graph.on("POST", "/search/query", {"value": [{"hitsContainers": [{"hits": [
        {"resource": drive_item("mine", "mine.docx")},
        {"resource": drive_item("theirs", "theirs.docx", drive="b!other")}]}]}]})
    r = await invoke("list_files", location="shared")
    assert not r.error, r.text
    assert [i["name"] for i in r.data["items"]] == ["theirs.docx"]
    assert not [s for s in graph.seen if "sharedWithMe" in s.path]
    assert graph.calls("POST", "/search/query")


async def test_a_recent_query_keeps_filtering_on_the_next_page(graph, invoke):
    # Found by review: the cursor of /me/drive/recent carried only the next link, so page two
    # returned every recent file, whatever the query.
    graph.on("GET", "/me/drive", {"id": OWN})
    nxt = "https://graph.microsoft.com/v1.0/me/drive/recent?$skiptoken=p2"

    def recent(seen):
        if seen.params.get("$skiptoken") == "p2":
            return httpx.Response(200, json={"value": [drive_item("c", "budget 2.xlsx"), drive_item("d", "other.docx")]})
        return httpx.Response(200, json={"value": [drive_item("a", "budget.xlsx"), drive_item("b", "notes.txt")],
                                         "@odata.nextLink": nxt})
    graph.on("GET", "/me/drive/recent", recent)
    r = await invoke("list_files", location="recent", query="budget")
    assert [i["name"] for i in r.data["items"]] == ["budget.xlsx"] and r.data["has_more"]
    r = await invoke("list_files", cursor=r.data["cursor"])
    assert not r.error, r.text
    assert [i["name"] for i in r.data["items"]] == ["budget 2.xlsx"]


async def test_recent_and_kind_filter(graph, invoke):
    graph.on("GET", "/me/drive", {"id": OWN})
    graph.on("GET", "/me/drive/recent", {"value": [drive_item("a", "a.xlsx"), drive_item("b", "b.docx")]})
    r = await invoke("list_files", location="recent", kind="spreadsheet")
    assert [i["name"] for i in r.data["items"]] == ["a.xlsx"]


async def test_modified_since_filters_the_page(graph, invoke):
    old = drive_item("o", "old.docx", lastModifiedDateTime="2026-01-01T00:00:00Z")
    graph.on("GET", "/me/drive/root:/Reports:/children", {"value": [old, drive_item("n", "new.docx")]})
    r = await invoke("list_files", folder="Reports", modified_since="2026-10-01")
    assert [i["name"] for i in r.data["items"]] == ["new.docx"]


async def test_site_location_resolves_the_site_then_lists_its_library(graph, invoke):
    graph.on("GET", "/me/drive", {"id": OWN})
    graph.on("GET", "/sites", {"value": [{"id": "S1", "displayName": "Marketing", "webUrl": "https://x.test/sites/m"}]})
    graph.on("GET", "/sites/S1/drive/root/children", {"value": [drive_item("s1", "brand.pptx", drive=SITE_DRIVE)]})
    r = await invoke("list_files", location="site", site="Marketing")
    assert not r.error, r.text
    assert graph.calls("GET", "/sites")[0].params["search"] == "Marketing"
    assert r.data["items"][0]["drive"] == SITE_DRIVE and r.data["site"] == "Marketing"


async def test_site_location_without_a_site_is_refused(graph, invoke):
    r = await invoke("list_files", location="site")
    assert r.error and "kind=site" in r.text and quiet(graph) == []


async def test_two_sites_with_one_name_are_ambiguous(graph, invoke):
    graph.on("GET", "/me/drive", {"id": OWN})
    graph.on("GET", "/sites", {"value": [
        {"id": "S1", "displayName": "Sales", "webUrl": "https://x.test/sites/s1"},
        {"id": "S2", "displayName": "Sales", "webUrl": "https://x.test/sites/s2"}]})
    r = await invoke("list_files", location="site", site="Sales")
    assert r.error and "https://x.test/sites/s1" in r.text and "https://x.test/sites/s2" in r.text


# -- list_files(kind=site) ------------------------------------------------------------------------------

async def test_kind_site_shows_libraries_and_visible_lists_only(graph, invoke):
    graph.on("GET", "/me/followedSites", {"value": [{"id": "S1", "displayName": "Ops", "webUrl": "https://x.test/sites/ops"}]})
    graph.on("GET", "/sites/S1/drives", {"value": [{"id": "d1", "name": "Documents"}]})
    graph.on("GET", "/sites/S1/lists", {"value": [
        {"id": "l1", "displayName": "Requests", "list": {"template": "genericList", "hidden": False}},
        {"id": "l2", "displayName": "appdata", "list": {"template": "genericList", "hidden": True}},
        {"id": "l3", "displayName": "Documents", "list": {"template": "documentLibrary"}}]})
    r = await invoke("list_files", kind="site", include_lists=True)
    assert not r.error, r.text
    s = r.data["items"][0]
    assert s["libraries"] == [{"id": "d1", "name": "Documents"}]
    assert s["lists"] == [{"id": "l1", "name": "Requests"}]
    plain = await invoke("list_files", kind="site")
    assert "lists" not in plain.data["items"][0]
    assert [s for s in graph.seen if "/drive/" in s.path or s.path.startswith("/search")] == []


async def test_kind_site_searches_sites_by_name(graph, invoke):
    graph.on("GET", "/sites", {"value": []})
    r = await invoke("list_files", kind="site", query="ops")
    assert not r.error and graph.calls("GET", "/sites")[0].params["search"] == "ops"


async def test_a_file_cursor_is_not_a_site_cursor(graph, invoke, rt):
    from m365_mcp.identity import current_person
    cur = rt.cursors.make("list_files", current_person().key, {"from": 10, "q": "budget", "loc": "my"})
    r = await invoke("list_files", kind="site", cursor=cur)
    assert r.error and "cursor" in r.text and quiet(graph) == []


async def test_include_lists_is_for_sites_only(graph, invoke):
    r = await invoke("list_files", include_lists=True)
    assert r.error and "kind=site" in r.text and quiet(graph) == []


# -- get_file --------------------------------------------------------------------------------

async def test_get_file_permissions_never_show_links(graph, invoke):
    graph.on("GET", "/me/drive/items/f1", drive_item("f1", "plan.docx", createdBy={"user": {"displayName": "Ann"}}))
    graph.on("GET", "/me/drive/items/f1/permissions", {"value": [
        {"id": "p1", "roles": ["owner"], "grantedToV2": {"user": {"displayName": "Ann", "email": "ann@x.test"}}},
        {"id": "p2", "roles": ["read"], "link": {"type": "view", "scope": "organization",
                                                  "webUrl": "https://x.test/SECRETLINK", "token": "TOK"}},
        {"id": "p3", "roles": ["write"], "inheritedFrom": {"id": "parent"},
         "grantedToV2": {"user": {"displayName": "Bo", "email": "bo@x.test"}}}]})
    graph.on("GET", "/me/drive/items/f1/versions", {"value": [
        {"id": "3.0", "lastModifiedDateTime": "2026-10-20T10:00:00Z", "lastModifiedBy": {"user": {"displayName": "Bo"}}, "size": 9}]})
    r = await invoke("get_file", id="f1", include_permissions=True, include_versions=True)
    assert not r.error, r.text
    assert "SECRETLINK" not in r.text and "TOK" not in r.text
    access = r.data["access"]
    assert access[0]["who"] == ["Ann <ann@x.test>"] and access[1]["link"] == "view link, organization"
    assert access[2]["inherited"] is True
    assert r.data["versions"][0]["id"] == "3.0" and r.data["created_by"] == "Ann"


async def test_get_file_by_path(graph, invoke):
    graph.on("GET", "/me/drive/root:/Reports/q3.xlsx", drive_item("f9", "q3.xlsx"))
    r = await invoke("get_file", path="Reports/q3.xlsx")
    assert not r.error and r.data["id"] == "f9"


async def test_get_file_needs_an_id_or_a_path(graph, invoke):
    r = await invoke("get_file")
    assert r.error and quiet(graph) == []


# -- download_file ---------------------------------------------------------------------------

async def test_download_issues_a_quoted_curl_with_a_drive_link(graph, invoke, rt):
    graph.on("GET", "/me/drive/items/f1", drive_item("f1", "it's $(rm -rf x) q3.docx"))
    r = await invoke("download_file", id="f1")
    assert not r.error, r.text
    cmd = re.search(r"curl -sS --fail-with-body -o '[^\n]*?' 'https://links", r.text).group(0)  # the text holds the one copy
    assert "$(" not in cmd.split("'")[1] and "'" not in cmd.split("'")[1]
    link = rt.signer.verify(link_of(r.text, "d"), op="d")
    assert link.kind == "drive" and link.target == {"i": "f1", "n": "it's $(rm -rf x) q3.docx"}
    assert r.data["expires_in_seconds"] == 300 and graph.writes() == []


async def test_download_in_another_drive_and_as_pdf(graph, invoke, rt):
    graph.on("GET", "/drives/b!site/items/f2", drive_item("f2", "deck.pptx"))
    r = await invoke("download_file", id="f2", drive="b!site", format="pdf")
    assert not r.error, r.text
    link = rt.signer.verify(link_of(r.text, "d"), op="d")
    assert link.target == {"i": "f2", "d": "b!site", "f": "pdf", "n": "deck.pdf"}
    assert "'deck.pdf'" in r.text


async def test_download_of_a_folder_is_refused(graph, invoke):
    graph.on("GET", "/me/drive/items/d1", drive_item("d1", "Archive", folder=True))
    r = await invoke("download_file", id="d1")
    assert r.error and "folder" in r.text and "https://links.test" not in r.text


async def test_pdf_of_an_unsupported_type_is_refused_with_the_supported_ones(graph, invoke):
    graph.on("GET", "/me/drive/items/z1", drive_item("z1", "data.zip"))
    r = await invoke("download_file", id="z1", format="pdf")
    assert r.error and "docx" in r.text and "https://links.test" not in r.text


async def test_download_without_link_keys_says_so(graph, invoke, rt):
    rt.signer = None
    graph.on("GET", "/me/drive/items/f1", drive_item("f1", "a.docx"))
    r = await invoke("download_file", id="f1")
    assert r.error and "not configured" in r.text


# -- upload_file -----------------------------------------------------------------------------

async def test_upload_resolves_the_folder_and_returns_a_link(graph, invoke, rt):
    graph.on("GET", "/me/drive/root:/Reports/2026", {"id": "P1", "name": "2026", "folder": {}})
    r = await invoke("upload_file", folder="Reports/2026", name="q3.xlsx", on_conflict="replace")
    assert not r.error, r.text
    link = rt.signer.verify(link_of(r.text, "u"), op="u")
    assert link.kind == "drive_upload" and link.target == {"p": "P1", "n": "q3.xlsx", "c": "replace"}
    assert link.max_bytes == 250 * 1024 * 1024
    assert "curl -sS --fail-with-body -T '<path of the local file>' '" in r.text and graph.writes() == []


async def test_upload_to_the_top_level_of_a_site_library(graph, invoke, rt):
    graph.on("GET", "/drives/b!site/root", {"id": "R1", "folder": {}})
    r = await invoke("upload_file", folder="", name="a.txt", drive="b!site")
    assert not r.error, r.text
    assert rt.signer.verify(link_of(r.text, "u"), op="u").target == {"p": "R1", "n": "a.txt", "c": "fail", "d": "b!site"}


@pytest.mark.parametrize("name", ["", "a/b.txt", "a:b.txt", "q?.txt", "...", "x" * 300])
async def test_upload_refuses_bad_names(graph, invoke, name):
    r = await invoke("upload_file", folder="Reports", name=name)
    assert r.error and quiet(graph) == []


async def test_upload_to_a_file_path_is_refused(graph, invoke):
    graph.on("GET", "/me/drive/root:/Reports/a.txt", {"id": "F", "name": "a.txt", "file": {}})
    r = await invoke("upload_file", folder="Reports/a.txt", name="b.txt")
    assert r.error and "not a folder" in r.text


# -- update_file -----------------------------------------------------------------------------

async def test_move_patches_the_parent_reference(graph, invoke):
    graph.on("GET", "/me/drive/root:/Archive", {"id": "A1", "name": "Archive", "folder": {}})
    graph.on("PATCH", "/me/drive/items/f[12]", {"id": "x"})
    r = await invoke("update_file", action="move", ids=["f1", "f2"], target_folder="Archive")
    assert not r.error, r.text
    assert [(s.path, s.json) for s in graph.writes()] == [
        ("/me/drive/items/f1", {"parentReference": {"id": "A1"}}),
        ("/me/drive/items/f2", {"parentReference": {"id": "A1"}})]


async def test_rename_patches_the_name_and_needs_one_item(graph, invoke):
    graph.on("PATCH", "/me/drive/items/f1", {"id": "f1"})
    r = await invoke("update_file", action="rename", ids=["f1"], new_name="final.docx")
    assert not r.error and graph.writes()[0].json == {"name": "final.docx"}
    two = await invoke("update_file", action="rename", ids=["f1", "f2"], new_name="x.docx")
    assert two.error and len(graph.writes()) == 1


async def test_create_folder_posts_the_folder_and_fails_on_a_conflict(graph, invoke):
    graph.on("GET", "/me/drive/root:/Reports", {"id": "P1", "name": "Reports", "folder": {}})
    graph.on("POST", "/me/drive/items/P1/children", {"id": "N1", "name": "2027", "webUrl": "https://x.test/2027"})
    r = await invoke("update_file", action="create_folder", target_folder="Reports", new_name="2027")
    assert not r.error, r.text
    assert graph.writes()[0].json == {"name": "2027", "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}


async def test_copy_waits_on_a_graph_monitor_url(graph, invoke):
    graph.on("GET", "/me/drive", {"id": OWN})
    graph.on("GET", "/me/drive/root:/Backup", {"id": "B1", "name": "Backup", "folder": {}})
    graph.on("POST", "/me/drive/items/f1/copy", None, status=202,
             headers={"Location": "https://graph.microsoft.com/v1.0/monitor/m1"})
    polls = iter([{"status": "inProgress"}, {"status": "completed", "resourceId": "n1"}])
    graph.on("GET", "/monitor/m1", lambda s: httpx.Response(200, json=next(polls)))
    r = await invoke("update_file", action="copy", ids=["f1"], target_folder="Backup")
    assert not r.error, r.text
    assert graph.writes()[0].json == {"parentReference": {"driveId": OWN, "id": "B1"}}
    assert r.data["copies"] == {"f1": "completed"} and len(graph.calls("GET", "/monitor/m1")) == 2


def _copies(graph, outcomes):
    """Copy of each id in `outcomes` answers 202 with a monitor that reports that outcome."""
    graph.on("GET", "/me/drive", {"id": OWN})
    graph.on("GET", "/me/drive/root:/Backup", {"id": "B1", "name": "Backup", "folder": {}})
    for i, outcome in outcomes.items():
        graph.on("POST", f"/me/drive/items/{i}/copy", None, status=202,
                 headers={"Location": f"https://graph.microsoft.com/v1.0/monitor/{i}"})
        graph.on("GET", f"/monitor/{i}", outcome)


NAME_CLASH = {"status": "failed", "error": {"code": "nameAlreadyExists", "message": "Name already exists"}}


async def test_a_copy_that_failed_at_the_monitor_is_an_error_with_its_reason(graph, invoke):
    # Measured: copying into a folder that holds the name answers 202, and the monitor then reports
    # {"status": "failed", "error": {"code": "nameAlreadyExists"}}; the tool said done 1.
    _copies(graph, {"f1": NAME_CLASH})
    r = await invoke("update_file", action="copy", ids=["f1"], target_folder="Backup")
    assert r.error and "No item was confirmed copied" in r.text
    assert "f1: An item with that name already exists there" in r.text and "new_name" in r.text


async def test_a_copy_failing_for_some_items_lists_them_under_failed(graph, invoke):
    _copies(graph, {"f1": NAME_CLASH, "f2": {"status": "completed"}})
    r = await invoke("update_file", action="copy", ids=["f1", "f2"], target_folder="Backup")
    assert not r.error, r.text
    assert r.data["done"] == 1 and r.data["copies"] == {"f2": "completed"}
    assert [f["id"] for f in r.data["failed"]] == ["f1"] and "already exists" in r.data["failed"][0]["reason"]


@pytest.mark.parametrize("action", ["move", "copy"])
async def test_a_folder_into_its_own_subfolder_is_refused_with_why(graph, invoke, action):
    # Measured: Graph answers 400 invalidRequest ("Invalid request" for a move); QA read only
    # "refused the request".
    graph.on("GET", "/me/drive", {"id": OWN})
    graph.on("GET", "/me/drive/root:/A/sub", {"id": "S1", "name": "sub", "folder": {}, "parentReference": {"path": "/drive/root:/A"}})
    graph.on("GET", "/me/drive/items/A1", {"id": "A1", "name": "A", "folder": {}, "parentReference": {"path": "/drive/root:"}})
    refused = {"error": {"code": "invalidRequest", "message": "Invalid request"}}
    graph.on("PATCH", "/me/drive/items/A1", refused, status=400)
    graph.on("POST", "/me/drive/items/A1/copy", refused, status=400)
    r = await invoke("update_file", action=action, ids=["A1"], target_folder="A/sub")
    assert r.error and "A1: A folder cannot go into itself or into a folder inside it" in r.text, r.text


async def test_copy_polls_the_sharepoint_monitor_without_the_bearer(graph, invoke):
    graph.on("GET", "/me/drive", {"id": OWN})
    graph.on("GET", "/me/drive/root:/Backup", {"id": "B1", "name": "Backup", "folder": {}})
    graph.on("POST", "/me/drive/items/f1/copy", None, status=202,
             headers={"Location": "https://contoso.sharepoint.com/_api/v2.0/monitor/zzz"})
    graph.on("GET", "https://contoso.sharepoint.com/_api/v2.0/monitor/zzz", {"status": "completed"})
    r = await invoke("update_file", action="copy", ids=["f1"], target_folder="Backup")
    assert not r.error and r.data["copies"] == {"f1": "completed"}
    polls = [s for s in graph.seen if "sharepoint.com" in s.path]
    assert polls and all("authorization" not in s.headers for s in polls)


async def test_copy_never_polls_a_monitor_outside_microsoft_file_hosts(graph, invoke):
    graph.on("GET", "/me/drive", {"id": OWN})
    graph.on("GET", "/me/drive/root:/Backup", {"id": "B1", "name": "Backup", "folder": {}})
    graph.on("POST", "/me/drive/items/f1/copy", None, status=202, headers={"Location": "https://evil.test/monitor"})
    r = await invoke("update_file", action="copy", ids=["f1"], target_folder="Backup")
    assert not r.error and r.data["copies"] == {"f1": "unknown"}
    assert not [s for s in graph.seen if "evil.test" in s.path]


async def test_delete_a_file_goes_to_the_recycle_bin_and_says_where_to_restore(graph, invoke):
    graph.on("GET", "/me/drive/items/f1", drive_item("f1", "old.docx"))
    graph.on("DELETE", "/me/drive/items/f1", None, status=204)
    r = await invoke("update_file", action="delete", ids=["f1"])
    assert not r.error, r.text
    assert "recycle bin" in r.data["note"] and "cannot restore" in r.data["note"]
    assert [s.method for s in graph.writes()] == ["DELETE"]
    assert_deletes_stay_in_the_allowlist(graph)


async def test_delete_a_folder_without_confirmation_is_refused_and_nothing_is_deleted(graph, invoke):
    graph.on("GET", "/me/drive/items/f1", drive_item("f1", "old.docx"))
    graph.on("GET", "/me/drive/items/d1", drive_item("d1", "Projects", folder=True))
    graph.on("DELETE", "/me/drive/items/[fd]1", None, status=204)
    r = await invoke("update_file", action="delete", ids=["f1", "d1"])
    assert r.error and "Projects" in r.text and "Nothing was deleted" in r.text
    assert graph.writes() == []


async def test_delete_a_folder_with_confirmation(graph, invoke):
    graph.on("GET", "/me/drive/items/d1", drive_item("d1", "Projects", folder=True))
    graph.on("DELETE", "/me/drive/items/d1", None, status=204)
    r = await invoke("update_file", action="delete", ids=["d1"], user_confirmed=True)
    assert not r.error and [s.path for s in graph.writes()] == ["/me/drive/items/d1"]
    assert_deletes_stay_in_the_allowlist(graph)


async def test_restore_version(graph, invoke):
    graph.on("POST", "/me/drive/items/f1/versions/2.0/restoreVersion", None, status=204)
    r = await invoke("update_file", action="restore_version", ids=["f1"], version="2.0")
    assert not r.error and graph.writes()[0].path == "/me/drive/items/f1/versions/2.0/restoreVersion"
    assert (await invoke("update_file", action="restore_version", ids=["f1"])).error


async def test_more_than_20_ids_is_refused(graph, invoke):
    r = await invoke("update_file", action="delete", ids=[f"f{i}" for i in range(21)])
    assert r.error and "20" in r.text and quiet(graph) == []


@pytest.mark.parametrize("ident", ["../x", "a/b", "a?b", "a b"])
async def test_ids_that_could_change_the_url_are_refused(graph, invoke, ident):
    r = await invoke("update_file", action="delete", ids=[ident], user_confirmed=True)
    assert r.error and quiet(graph) == []


async def test_one_failure_does_not_hide_the_others(graph, invoke):
    graph.on("PATCH", "/me/drive/items/f1", {"id": "f1"})
    graph.on("PATCH", "/me/drive/items/f2", {"error": {"code": "itemNotFound"}}, status=404)
    graph.on("GET", "/me/drive/root:/Archive", {"id": "A1", "name": "Archive", "folder": {}})
    r = await invoke("update_file", action="move", ids=["f1", "f2"], target_folder="Archive")
    assert not r.error and r.data["done"] == 1 and r.data["failed"][0]["id"] == "f2"


# -- share_file ------------------------------------------------------------------------------

async def test_share_link_is_organisation_scoped(graph, invoke):
    graph.on("POST", "/me/drive/items/f1/createLink", {"link": {"type": "view", "scope": "organization", "webUrl": "https://x.test/l"}})
    r = await invoke("share_file", id="f1", action="link", user_confirmed=True)
    assert not r.error, r.text
    assert graph.writes()[0].json == {"type": "view", "scope": "organization"}
    assert r.data["link"] == "https://x.test/l" and r.data["reach"] == "organization"


async def test_share_link_edit_still_organisation_scoped(graph, invoke):
    graph.on("POST", "/me/drive/items/f1/createLink", {"link": {"type": "edit", "scope": "organization", "webUrl": "https://x.test/l"}})
    await invoke("share_file", id="f1", action="link", role="edit", user_confirmed=True)
    assert graph.writes()[0].json == {"type": "edit", "scope": "organization"}


@pytest.mark.parametrize("action", ["link", "invite"])
async def test_share_without_confirmation_writes_nothing(graph, invoke, action):
    person(graph, "Bo", "bo@x.test")
    r = await invoke("share_file", id="f1", action=action, people=["Bo"])
    assert r.error and "confirm" in r.text
    assert graph.writes() == []


async def test_invite_sends_a_signed_in_invitation(graph, invoke):
    person(graph, "Bo", "bo@x.test")
    graph.on("POST", "/me/drive/items/f1/invite", {"value": []})
    r = await invoke("share_file", id="f1", action="invite", people=["Bo"], role="edit", message="Please review",
                     user_confirmed=True)
    assert not r.error, r.text
    assert graph.writes()[0].json == {"recipients": [{"email": "bo@x.test"}], "roles": ["write"],
                                      "sendInvitation": True, "requireSignIn": True, "message": "Please review"}
    assert r.data["invited"] == ["Bo <bo@x.test>"] and r.data["sends_left_now"] == 9


async def test_invite_more_than_20_people_is_refused(graph, invoke):
    r = await invoke("share_file", id="f1", action="invite", people=[f"p{i}" for i in range(21)], user_confirmed=True)
    assert r.error and "20" in r.text and graph.writes() == []


async def test_invite_of_a_distribution_list_is_refused(graph, invoke):
    person(graph, "Everyone", "all@x.test", group=True)
    r = await invoke("share_file", id="f1", action="invite", people=["Everyone"], user_confirmed=True)
    assert r.error and "all@x.test" in r.text and graph.writes() == []


async def test_invite_needs_people(graph, invoke):
    r = await invoke("share_file", id="f1", action="invite", user_confirmed=True)
    assert r.error and quiet(graph) == []


async def test_share_draws_from_the_shared_bucket(graph, invoke):
    graph.on("POST", "/me/drive/items/f1/createLink", {"link": {"webUrl": "https://x.test/l"}})
    for _ in range(10):
        assert not (await invoke("share_file", id="f1", action="link", user_confirmed=True)).error
    r = await invoke("share_file", id="f1", action="link", user_confirmed=True)
    assert r.error and "Wait" in r.text and len(graph.writes()) == 10


PERMS = {"value": [
    {"id": "own", "roles": ["owner"], "grantedToV2": {"user": {"displayName": "Ann", "email": "ann@x.test"}}},
    {"id": "bo", "roles": ["read"], "grantedToV2": {"user": {"displayName": "Bo", "email": "bo@x.test"}}},
    {"id": "inh", "roles": ["read"], "inheritedFrom": {"id": "p"},
     "grantedToV2": {"user": {"displayName": "Cy", "email": "cy@x.test"}}},
    {"id": "lnk", "roles": ["read"], "link": {"type": "view", "scope": "users"},
     "grantedToIdentitiesV2": [{"user": {"displayName": "Di", "email": "di@x.test"}}, {"user": {"email": "ed@x.test"}}]}]}


async def test_revoke_deletes_only_the_direct_permission_of_the_named_person(graph, invoke):
    person(graph, "Bo", "bo@x.test")
    graph.on("GET", "/me/drive/items/f1/permissions", PERMS)
    graph.on("DELETE", "/me/drive/items/f1/permissions/bo", None, status=204)
    r = await invoke("share_file", id="f1", action="revoke", people=["Bo"])
    assert not r.error, r.text
    assert [s.path for s in graph.writes()] == ["/me/drive/items/f1/permissions/bo"]
    assert r.data["revoked"] == ["Bo <bo@x.test>"]
    assert_deletes_stay_in_the_allowlist(graph)


async def test_revoke_of_the_owner_is_refused(graph, invoke):
    person(graph, "Ann", "ann@x.test")
    graph.on("GET", "/me/drive/items/f1/permissions", PERMS)
    r = await invoke("share_file", id="f1", action="revoke", people=["Ann"])
    assert r.error and "owner" in r.text and graph.writes() == []


async def test_revoke_leaves_inherited_and_shared_link_access_alone_and_says_so(graph, invoke):
    person(graph, "Cy", "cy@x.test")
    graph.on("GET", "/me/drive/items/f1/permissions", PERMS)
    r = await invoke("share_file", id="f1", action="revoke", people=["cy@x.test", "di@x.test", "zed@x.test"])
    assert not r.error, r.text
    assert graph.writes() == []
    assert any("parent folder" in s for s in r.data["not_revoked"])
    assert any("sharing link" in s for s in r.data["not_revoked"])
    assert r.data["no_direct_access"] == ["zed@x.test"]


async def test_revoke_needs_no_confirmation(graph, invoke):
    person(graph, "Bo", "bo@x.test")
    graph.on("GET", "/me/drive/items/f1/permissions", PERMS)
    graph.on("DELETE", "/me/drive/items/f1/permissions/bo", None, status=204)
    r = await invoke("share_file", id="f1", action="revoke", people=["Bo"], user_confirmed=False)
    assert not r.error and len(graph.writes()) == 1


async def test_a_link_issued_for_a_download_cannot_upload(rt):
    s = rt.signer.issue(op="d", kind="drive", target={"i": "x"}, person="p", token="t", token_expires=2_000_000_000, ttl=300)
    with pytest.raises(LinkRefused):
        rt.signer.verify(s, op="u")


# -- what QA on a real tenant found ---------------------------------------------------------

SEARCH_DOWN = {"error": {"code": "InternalServerError", "message": "The call failed, please try again."}}


def team_site(graph):
    # Measured: /sites?search= and /me/followedSites both came back empty or 500 for a team's own
    # site; /me/joinedTeams then /groups/{id}/sites/root finds it.
    graph.on("GET", "/me/joinedTeams", {"value": [{"id": "G1", "displayName": "Ops team"}]})
    graph.on("GET", "/groups/G1/sites/root", {"id": "S9", "displayName": "Ops team", "webUrl": "https://x.test/sites/ops"})
    graph.on("GET", "/me/followedSites", {"value": []})
    graph.on("GET", "/sites", SEARCH_DOWN, status=500)
    graph.on("GET", "/sites/S9/drives", {"value": [{"id": "d9", "name": "Documents"}]})
    graph.on("GET", "/sites/S9/lists", {"value": [
        {"id": "l9", "displayName": "QA list", "list": {"template": "genericList", "hidden": False}}]})


async def test_the_sites_of_the_persons_teams_are_listed(graph, invoke):
    team_site(graph)
    r = await invoke("list_files", kind="site", include_lists=True)
    assert not r.error, r.text
    [s] = r.data["items"]
    assert s["name"] == "Ops team" and s["libraries"] == [{"id": "d9", "name": "Documents"}]
    assert s["lists"] == [{"id": "l9", "name": "QA list"}]


async def test_a_site_query_finds_a_team_site_while_search_is_down(graph, invoke):
    team_site(graph)
    r = await invoke("list_files", kind="site", query="ops")
    assert not r.error, r.text
    assert [s["name"] for s in r.data["items"]] == ["Ops team"]


async def test_a_team_site_is_found_by_name_for_its_files_and_lists(graph, invoke):
    team_site(graph)
    graph.on("GET", "/me/drive", {"id": OWN})
    graph.on("GET", "/sites/S9/drive/root/children", {"value": [drive_item("s1", "brand.pptx", drive="d9")]})
    r = await invoke("list_files", location="site", site="Ops team")
    assert not r.error, r.text
    assert r.data["site"] == "Ops team" and graph.calls("GET", "/sites") == []


async def test_a_file_search_microsoft_cannot_answer_falls_back_to_names_in_the_folder(graph, invoke):
    # Measured: /search/query (driveItem) and the drive's own search(q=) both answered 500 on the test
    # tenant for every query.
    graph.on("GET", "/me/drive", {"id": OWN})
    graph.on("POST", "/search/query", SEARCH_DOWN, status=500)
    graph.on("GET", "/me/drive/root/children", {"value": [drive_item("a", "Budget 2026.xlsx"), drive_item("b", "notes.txt")]})
    r = await invoke("list_files", query="budget")
    assert not r.error, r.text
    assert [i["name"] for i in r.data["items"]] == ["Budget 2026.xlsx"]
    assert "Microsoft Search" in r.data["note"]


async def test_the_name_fallback_continues_by_cursor_instead_of_dropping_matches(graph, invoke):
    # Found by QA: eleven matching names, limit 2: two came back with has_more false and no cursor.
    graph.on("GET", "/me/drive", {"id": OWN})
    graph.on("GET", "/me/drive/root:/Reports:/search\\(q='n'\\)", SEARCH_DOWN, status=500)
    graph.on("GET", "/me/drive/root:/Reports:/children", {"value": [drive_item(f"f{i}", f"n{i}.txt") for i in range(5)]
                                                          + [drive_item("x", "other.txt")]})
    r = await invoke("list_files", folder="Reports", query="n", limit=2)
    assert not r.error, r.text
    seen = [i["name"] for i in r.data["items"]]
    assert seen == ["n0.txt", "n1.txt"] and r.data["has_more"] is True and "Microsoft Search" in r.data["note"]
    while r.data.get("cursor"):
        r = await invoke("list_files", cursor=r.data["cursor"], limit=2)
        assert not r.error, r.text
        seen += [i["name"] for i in r.data["items"]]
    assert seen == [f"n{i}.txt" for i in range(5)]
    assert graph.writes() == []


async def test_a_folder_that_is_a_file_is_said_to_be_a_file(graph, invoke):
    # Measured: children of a file answer 422 notSupported, and its search(q=) 500; QA saw a bare
    # "refused the request" for both.
    graph.on("GET", "/me/drive", {"id": OWN})
    not_a_folder = {"error": {"code": "notSupported", "message": "Children cannot be listed from an item that is not a folder"}}
    graph.on("GET", "/me/drive/root:/a.txt:/children", not_a_folder, status=422)
    graph.on("GET", "/me/drive/root:/a.txt:/search\\(q='x'\\)", SEARCH_DOWN, status=500)
    for args in ({"folder": "a.txt"}, {"folder": "a.txt", "query": "x"}):
        r = await invoke("list_files", **args)
        assert r.error and "That folder path is a file, not a folder" in r.text and "get_file" in r.text, r.text
    assert graph.writes() == []


@pytest.mark.parametrize("name", ["con", "CON.txt", "aux.txt", "lpt1", "com9.log", "trail.", "x_vti_y.txt"])
async def test_a_name_onedrive_reserves_is_refused_before_a_link_is_made(graph, invoke, name):
    # Measured: PUT of "con" and "aux.txt" answered 400 invalidRequest, "trail." 400 "illegal characters",
    # "_vti_x" 403; the tool had handed out an upload link for each, and rename got the bare refusal.
    graph.on("GET", "/me/drive/root", {"id": "R", "name": "root", "folder": {}})
    up = await invoke("upload_file", folder="", name=name)
    rn = await invoke("update_file", action="rename", ids=["f1"], new_name=name)
    mk = await invoke("update_file", action="create_folder", new_name=name)
    for r in (up, rn, mk):
        assert r.error and "OneDrive does not accept this name" in r.text and "reserved" in r.text, r.text
    assert "/u/" not in up.text
    assert graph.writes() == []


@pytest.mark.parametrize("name", ["console.txt", "a..b", "x.lock", "~$x.docx", "desktop.ini"])
async def test_names_onedrive_accepts_get_a_link(graph, invoke, name):
    # Measured: each of these was created by a PUT.
    graph.on("GET", "/me/drive/root", {"id": "R", "name": "root", "folder": {}})
    r = await invoke("upload_file", folder="", name=name)
    assert not r.error, r.text


async def test_shared_files_while_microsoft_search_is_down_is_said_plainly(graph, invoke):
    graph.on("GET", "/me/drive", {"id": OWN})
    graph.on("POST", "/search/query", SEARCH_DOWN, status=500)
    r = await invoke("list_files", location="shared")
    assert r.error and "Microsoft Search" in r.text and "shared" in r.text
    assert graph.writes() == [g for g in graph.writes() if g.path == "/search/query"]


async def test_a_name_with_a_hash_is_refused_everywhere(graph, invoke):
    # Measured: the upload link took "a#b.txt" and then answered 400; a folder "a#b" was made but could
    # not be reached by path. The characters are refused up front, consistently.
    graph.on("GET", "/me/drive/root", {"id": "R", "name": "root", "folder": {}})
    graph.on("POST", "/me/drive/items/R/children", {"id": "N"}, status=201)
    graph.on("PATCH", "/me/drive/items/f1", {"id": "f1"})
    up = await invoke("upload_file", folder="", name="a#b.txt")
    mk = await invoke("update_file", action="create_folder", new_name="a#b")
    rn = await invoke("update_file", action="rename", ids=["f1"], new_name="a#b.txt")
    dots = await invoke("update_file", action="create_folder", new_name="..")
    for r in (up, mk, rn, dots):
        assert r.error and "#" in r.text, r.text
    assert graph.writes() == []


async def test_a_path_with_a_hash_is_refused_with_the_way_to_reach_it(graph, invoke):
    r = await invoke("get_file", path="qa/a#b")
    assert r.error and "id" in r.text and "#" in r.text
    assert quiet(graph) == []


async def test_pdf_conversion_offers_the_types_microsoft_converts(graph, invoke):
    # Measured: csv, txt, json and xml answer 406 at the download URL; md, html, rtf and eml convert.
    graph.on("GET", "/me/drive/items/c1", {"id": "c1", "name": "notes.csv", "file": {}})
    r = await invoke("download_file", id="c1", format="pdf")
    assert r.error and "csv" not in r.text.split("Convertible types:")[1]
    graph.on("GET", "/me/drive/items/x1", {"id": "x1", "name": "book.xlsm", "file": {}})
    assert not (await invoke("download_file", id="x1", format="pdf")).error


async def test_a_change_that_failed_for_every_item_is_an_error(graph, invoke):
    graph.on("PATCH", "/me/drive/items/f1", {"error": {"code": "nameAlreadyExists"}}, status=409)
    r = await invoke("update_file", action="rename", ids=["f1"], new_name="taken.docx")
    assert r.error and "f1" in r.text and "No item was confirmed renamed" in r.text


async def test_a_folder_that_exists_is_refused_with_its_path(graph, invoke):
    # QA: create_folder made "sub1 1" next to an existing "sub1" and reported success.
    graph.on("GET", "/me/drive/root:/Reports", {"id": "P1", "name": "Reports", "folder": {}})
    graph.on("POST", "/me/drive/items/P1/children", {"error": {"code": "nameAlreadyExists"}}, status=409)
    r = await invoke("update_file", action="create_folder", target_folder="Reports", new_name="2027")
    assert r.error and "Reports/2027" in r.text and "already exists" in r.text
    assert graph.writes()[0].json["@microsoft.graph.conflictBehavior"] == "fail"


async def test_more_matching_sites_than_the_page_holds_continue_by_cursor(graph, invoke):
    graph.on("GET", "/me/joinedTeams", {"value": [{"id": "G1"}, {"id": "G2"}]})
    graph.on("GET", "/groups/G1/sites/root", {"id": "S1", "displayName": "Ops one"})
    graph.on("GET", "/groups/G2/sites/root", {"id": "S2", "displayName": "Ops two"})
    graph.on("GET", "/me/followedSites", {"value": []})
    graph.on("GET", "/sites", {"value": [{"id": "S3", "displayName": "Ops three"}]})
    graph.on("GET", "/sites/S[0-9]/drives", {"value": []})
    seen, cursor = [], None
    for _ in range(4):
        r = await invoke("list_files", kind="site", query="ops", limit=1, **({"cursor": cursor} if cursor else {}))
        assert not r.error, r.text
        seen += [s["id"] for s in r.data["items"]]
        cursor = r.data.get("cursor")
        if not cursor:
            break
    assert seen == ["S1", "S2", "S3"]


async def test_site_search_pages_skip_nothing_after_the_known_sites(graph, invoke):
    graph.on("GET", "/me/joinedTeams", {"value": [{"id": "G1"}]})
    graph.on("GET", "/groups/G1/sites/root", {"id": "S1", "displayName": "Ops one"})
    graph.on("GET", "/me/followedSites", {"value": []})
    graph.on("GET", "/sites", {"value": [{"id": "S1", "displayName": "Ops one"}, {"id": "S2", "displayName": "Ops two"},
                                         {"id": "S3", "displayName": "Ops three"}]})
    graph.on("GET", "/sites/S[0-9]/drives", {"value": []})
    first = await invoke("list_files", kind="site", query="ops", limit=2)
    assert [s["id"] for s in first.data["items"]] == ["S1"] and first.data["has_more"] is True
    second = await invoke("list_files", kind="site", query="ops", limit=2, cursor=first.data["cursor"])
    assert [s["id"] for s in second.data["items"]] == ["S2", "S3"]


async def test_a_site_name_while_search_is_down_is_not_called_missing(graph, invoke):
    graph.on("GET", "/me/joinedTeams", {"value": []})
    graph.on("GET", "/me/followedSites", {"value": []})
    graph.on("GET", "/sites", SEARCH_DOWN, status=500)
    graph.on("GET", "/me/drive", {"id": OWN})
    r = await invoke("list_files", location="site", site="Marketing")
    assert r.error and "Microsoft Search did not answer" in r.text and "address" in r.text


async def test_a_continued_listing_returns_the_whole_page_whatever_the_limit(graph, invoke):
    nxt = "https://graph.microsoft.com/v1.0/me/drive/root/children?$skiptoken=2"
    graph.on("GET", "/me/drive", {"id": OWN})
    graph.on("GET", "/me/drive/root/children", lambda s: httpx.Response(200, json={
        "value": [drive_item(f"n{i}", f"n{i}.txt") for i in range(3)]} if s.params.get("$skiptoken") else {
        "value": [drive_item("a", "a.txt")], "@odata.nextLink": nxt}))
    first = await invoke("list_files", limit=3)
    assert not first.error, first.text
    more = await invoke("list_files", limit=1, cursor=first.data["cursor"])
    assert not more.error, more.text
    assert [i["id"] for i in more.data["items"]] == ["n0", "n1", "n2"]


async def test_a_partial_match_among_ones_own_sites_does_not_hide_the_exact_one(graph, invoke):
    graph.on("GET", "/me/joinedTeams", {"value": []})
    graph.on("GET", "/me/followedSites", {"value": [{"id": "S1", "displayName": "Finance Archive"}]})
    graph.on("GET", "/sites", {"value": [{"id": "S2", "displayName": "Finance"}]})
    graph.on("GET", "/me/drive", {"id": OWN})
    graph.on("GET", "/sites/S2/drive/root/children", {"value": []})
    r = await invoke("list_files", location="site", site="Finance")
    assert not r.error, r.text
    assert r.data["site"] == "Finance"


# -- round 3 -------------------------------------------------------------------------------------

def remote_item(i, name, **kw):
    """What /me/drive/recent returns for a file in someone else's drive: the facts sit in remoteItem."""
    inner = {"id": i, "name": name, "size": 555, "webUrl": f"https://x.test/{name}", "file": {"mimeType": "text/plain"},
             "parentReference": {"driveId": "b!other"}, "fileSystemInfo": {"lastModifiedDateTime": "2026-10-21T08:00:00Z"}}
    return {"id": i, "remoteItem": inner, "fileSystemInfo": {"lastAccessedDateTime": "2026-10-22T08:00:00Z"}, **kw}


async def test_recent_items_show_their_name_not_only_their_id(graph, invoke):
    # Round 3 (F01): $select on /me/drive/recent drops everything but the id, and the real facts of
    # a file that lives elsewhere sit in remoteItem.
    graph.on("GET", "/me/drive", {"id": OWN})
    graph.on("GET", "/me/drive/recent", {"value": [remote_item("r1", "plan.docx"), drive_item("a", "a.xlsx")]})
    r = await invoke("list_files", location="recent")
    assert not r.error, r.text
    assert "$select" not in graph.calls("GET", "/me/drive/recent")[0].params
    first, second = r.data["items"]
    assert first["name"] == "plan.docx" and first["kind"] == "document" and first["size"] == 555
    assert first["drive"] == "b!other" and first["modified"].startswith("2026-10-21")
    assert first["last_used"].startswith("2026-10-22") and second["name"] == "a.xlsx"


async def test_modified_since_works_on_recent_items_too(graph, invoke):
    graph.on("GET", "/me/drive", {"id": OWN})
    old = remote_item("r2", "old.txt")
    old["remoteItem"]["fileSystemInfo"]["lastModifiedDateTime"] = "2026-01-01T00:00:00Z"
    graph.on("GET", "/me/drive/recent", {"value": [remote_item("r1", "new.txt"), old]})
    r = await invoke("list_files", location="recent", modified_since="2026-10-01")
    assert [i["name"] for i in r.data["items"]] == ["new.txt"]


async def test_shared_items_say_who_made_them_and_drop_empty_site_folders(graph, invoke):
    # Round 3 (F12): "files Megan shared" could not be answered, and site folders and notebooks of
    # size 0 came back as files.
    graph.on("GET", "/me/drive", {"id": OWN})
    theirs = drive_item("t", "theirs.docx", drive="b!other", createdBy={"user": {"displayName": "Megan", "email": "megan@x.test"}})
    noise = drive_item("n", "Notebooks", drive="b!other", size=0)
    graph.on("POST", "/search/query", {"value": [{"hitsContainers": [{"hits": [
        {"resource": theirs}, {"resource": noise}]}]}]})
    r = await invoke("list_files", location="shared")
    assert [i["name"] for i in r.data["items"]] == ["theirs.docx"]
    assert r.data["items"][0]["created_by"] == "Megan"


async def test_an_empty_shared_listing_warns_about_the_lag(graph, invoke):
    graph.on("GET", "/me/drive", {"id": OWN})
    graph.on("POST", "/search/query", {"value": [{"hitsContainers": [{"hits": []}]}]})
    r = await invoke("list_files", location="shared")
    assert r.data["items"] == [] and "minute" in r.data["note"]


async def test_rename_echoes_the_new_name_and_path(graph, invoke):
    # Round 3 (F05): the result said done:1 and a model had to call get_file to see what changed.
    graph.on("PATCH", "/me/drive/items/f1", drive_item("f1", "final.docx"))
    r = await invoke("update_file", action="rename", ids=["f1"], new_name="final.docx")
    assert r.data["items"] == [{"id": "f1", "name": "final.docx", "path": "Reports/final.docx"}]


async def test_move_echoes_where_the_item_went(graph, invoke):
    graph.on("GET", "/me/drive/root:/Archive", {"id": "A1", "name": "Archive", "folder": {}})
    graph.on("PATCH", "/me/drive/items/f1", drive_item("f1", "a.docx", path="/drive/root:/Archive"))
    r = await invoke("update_file", action="move", ids=["f1"], target_folder="Archive")
    assert r.data["items"][0]["path"] == "Archive/a.docx"


async def test_a_download_command_is_in_the_structured_result_the_agent_reads(graph, invoke):
    # Round 3 replaced the command in the structured copy with "(the command is
    # in the text result)" to save context. Claude Code shows the model the structured copy, so the
    # agent got no command and could not download any file.
    graph.on("GET", "/me/drive/items/f1", drive_item("f1", "a.docx"))
    r = await invoke("download_file", id="f1")
    assert r.data["download"].startswith("curl ") and "https://links.test/d/" in r.data["download"]
    assert "text result" not in r.text
    assert "data written by someone else" in r.data["content_note"]


async def test_an_upload_command_is_in_the_structured_result_the_agent_reads(graph, invoke):
    graph.on("GET", "/me/drive/root:/Reports", {"id": "P1", "name": "Reports", "folder": {}})
    r = await invoke("upload_file", folder="Reports", name="a.txt")
    assert "https://links.test/u/" in r.data["upload"] and "text result" not in r.text


async def test_the_upload_description_names_the_call_that_makes_a_folder(server):
    tools = {t.name: t for t in await server.list_tools()}
    assert "create_folder" in tools["upload_file"].description


async def test_a_recent_stub_is_addressed_by_the_remote_items_id_and_drive(graph, invoke):
    # Review finding: the stub's id with the remote drive made an address that exists nowhere.
    graph.on("GET", "/me/drive", {"id": OWN})
    stub = remote_item("remote-id", "plan.txt")
    stub["id"] = "stub-id"
    graph.on("GET", "/me/drive/recent", {"value": [stub]})
    item = (await invoke("list_files", location="recent")).data["items"][0]
    assert item["id"] == "remote-id" and item["drive"] == "b!other"


async def test_an_empty_file_with_an_extension_is_still_shared_with_you(graph, invoke):
    graph.on("GET", "/me/drive", {"id": OWN})
    empty = drive_item("e", "empty.txt", drive="b!other", size=0)
    empty["file"]["mimeType"] = "text/plain"
    noise = drive_item("n", "Contoso - All Staff Notebook", drive="b!other", size=0)
    graph.on("POST", "/search/query", {"value": [{"hitsContainers": [{"hits": [{"resource": empty}, {"resource": noise}]}]}]})
    r = await invoke("list_files", location="shared")
    assert [i["name"] for i in r.data["items"]] == ["empty.txt"]


# -- round 4 ---------------------------------------------------------------------------------------

def search_hits(*items):
    return {"value": [{"hitsContainers": [{"hits": [{"resource": m} for m in items]}]}]}


def search_page(items, more):
    return {"value": [{"hitsContainers": [{"hits": [{"resource": m} for m in items], "moreResultsAvailable": more}]}]}


def searched(m):
    """A folder as Microsoft Search returns it (measured): a file facet of type octet-stream, size 0 or none."""
    return {k: v for k, v in m.items() if k != "folder"} | {"file": {"mimeType": "application/octet-stream"}, "size": 0}


async def test_recent_items_without_a_modified_time_are_dated_from_their_item(graph, invoke):
    # Round 4 (F01): /me/drive/recent carries lastAccessedDateTime only, measured live; the filter compared
    # a field that is never there and answered 0, and the items showed no modified time or path.
    graph.on("GET", "/me/drive", {"id": OWN})
    stub = {"id": "r1", "remoteItem": {"id": "r1", "name": "rec.mp4", "size": 5, "file": {"mimeType": "video/mp4"},
            "parentReference": {"driveId": OWN}}, "fileSystemInfo": {"lastAccessedDateTime": "2026-10-25T06:00:00Z"}}
    graph.on("GET", f"/drives/{OWN}/items/r1", drive_item("r1", "rec.mp4", path="/drive/root:/Recordings",
             lastModifiedDateTime="2026-10-23T19:19:07Z"))
    graph.on("GET", "/me/drive/recent", {"value": [stub]})
    r = await invoke("list_files", location="recent", modified_since="2026-10-19")
    assert not r.error, r.text
    item = r.data["items"][0]
    assert item["name"] == "rec.mp4" and item["modified"].startswith("2026-10-23") and item["path"] == "Recordings/rec.mp4"
    old = await invoke("list_files", location="recent", modified_since="2026-10-24")
    assert old.data["items"] == []


async def test_recent_items_that_cannot_be_dated_are_counted_not_hidden(graph, invoke):
    graph.on("GET", "/me/drive", {"id": OWN})
    stub = {"id": "r1", "remoteItem": {"id": "r1", "name": "x.txt", "parentReference": {"driveId": "b!other"}}}
    graph.on("GET", "/drives/b!other/items/r1", {"error": {"code": "accessDenied"}}, status=403)
    graph.on("GET", "/me/drive/recent", {"value": [stub]})
    r = await invoke("list_files", location="recent", modified_since="2026-10-19")
    assert not r.error and r.data["items"] == [] and "1" in r.data["note"] and "date" in r.data["note"]


async def test_modified_since_in_my_files_looks_inside_folders_through_search(graph, invoke):
    # Round 4 (F01): the top level lists folders whose own time changed, not the files inside them.
    graph.on("GET", "/me/drive", {"id": OWN})
    new = drive_item("n", "new.docx", path="/drive/root:/Deep/er", lastModifiedDateTime="2026-10-22T10:00:00Z")
    old = drive_item("o", "old.docx", lastModifiedDateTime="2026-10-01T10:00:00Z")
    folder = searched(drive_item("d", "Deep", folder=True, lastModifiedDateTime="2026-10-22T10:00:00Z"))
    theirs = drive_item("t", "theirs.docx", drive="b!other", lastModifiedDateTime="2026-10-22T11:00:00Z")
    graph.on("POST", "/search/query", search_hits(new, old, folder, theirs))
    r = await invoke("list_files", location="my", modified_since="2026-10-19")
    assert not r.error, r.text
    assert [i["name"] for i in r.data["items"]] == ["new.docx"]
    q = graph.calls("POST", "/search/query")[0].json["requests"][0]["query"]["queryString"]
    assert "lastModifiedTime>=2026-10-19" in q
    assert "index" in r.data["note"]  # Search lags: a file changed a minute ago may be missing
    assert not graph.calls("GET", "/me/drive/root/children")


async def test_modified_since_keeps_folders_when_asked_for_folders(graph, invoke):
    graph.on("GET", "/me/drive", {"id": OWN})
    folder = searched(drive_item("d", "Deep", folder=True, lastModifiedDateTime="2026-10-22T10:00:00Z"))
    graph.on("POST", "/search/query", search_hits(folder))
    r = await invoke("list_files", location="my", modified_since="2026-10-19", kind="folder")
    assert [i["name"] for i in r.data["items"]] == ["Deep"]


async def test_modified_since_while_search_is_down_says_only_the_top_level_was_read(graph, invoke):
    graph.on("GET", "/me/drive", {"id": OWN})
    graph.on("POST", "/search/query", {"error": {"code": "generalException"}}, status=500)
    graph.on("GET", "/me/drive/root/children", {"value": [drive_item("n", "new.docx")]})
    r = await invoke("list_files", location="my", modified_since="2026-10-01")
    assert not r.error and [i["name"] for i in r.data["items"]] == ["new.docx"]
    assert "top level" in r.data["note"]


async def test_the_shared_listing_fills_the_page_past_dropped_items(graph, invoke):
    # Round 4 (F12): one item and has_more, because filtering took the rest of a page of 10.
    graph.on("GET", "/me/drive", {"id": OWN})
    mine = [drive_item(f"m{i}", f"mine{i}.docx") for i in range(3)]
    theirs = [drive_item(f"t{i}", f"theirs{i}.docx", drive="b!other") for i in range(3)]

    def search(seen):
        start = seen.json["requests"][0]["from"]
        pages = {0: search_page(mine + theirs[:1], True), 4: search_page(theirs[1:], False)}
        return httpx.Response(200, json=pages[start])
    graph.on("POST", "/search/query", search)
    r = await invoke("list_files", location="shared", limit=3)
    assert [i["name"] for i in r.data["items"]] == ["theirs0.docx", "theirs1.docx", "theirs2.docx"]
    assert not r.data.get("has_more")


async def test_the_shared_listing_leaves_out_site_system_files(graph, invoke):
    graph.on("GET", "/me/drive", {"id": OWN})
    icon = drive_item("i", "__siteIcon__.png", drive="b!site", createdBy={"user": {"displayName": "System Account"}})
    real = drive_item("t", "plan.docx", drive="b!other")
    graph.on("POST", "/search/query", search_hits(icon, real))
    r = await invoke("list_files", location="shared")
    assert [i["name"] for i in r.data["items"]] == ["plan.docx"]


async def test_a_query_that_finds_nothing_says_search_lags_and_matches_names_in_the_folder(graph, invoke):
    # Round 4 (F02): a file uploaded a minute ago, asked for by its exact name, was "not there".
    graph.on("GET", "/me/drive", {"id": OWN})
    graph.on("GET", "/me/drive/root:/Reports:/search\\(q='plan-final.txt'\\)", {"value": []})
    graph.on("GET", "/me/drive/root:/Reports:/children", {"value": [
        drive_item("f1", "plan-final.txt"), drive_item("f2", "other.txt")]})
    r = await invoke("list_files", folder="Reports", query="plan-final.txt")
    assert [i["name"] for i in r.data["items"]] == ["plan-final.txt"]
    assert "minutes" in r.data["note"] and "name" in r.data["note"]


async def test_a_query_that_finds_nothing_anywhere_still_warns_about_the_lag(graph, invoke):
    graph.on("GET", "/me/drive", {"id": OWN})
    graph.on("GET", "/me/drive/root:/Reports:/search\\(q='nothing'\\)", {"value": []})
    graph.on("GET", "/me/drive/root:/Reports:/children", {"value": [drive_item("f2", "other.txt")]})
    r = await invoke("list_files", folder="Reports", query="nothing")
    assert r.data["items"] == [] and "minutes" in r.data["note"]


async def test_uploading_into_a_folder_that_is_not_there_names_the_next_step(graph, invoke):
    # Round 4 (F03): the answer talked of an id and of sharing, for a folder path.
    r = await invoke("upload_file", folder="Nope/Deeper", name="a.txt")
    assert r.error and "Nope/Deeper" in r.text and "create_folder" in r.text and "shared with you" not in r.text
    assert graph.writes() == []


async def test_moving_into_a_folder_that_is_not_there_names_the_next_step(graph, invoke):
    r = await invoke("update_file", action="move", ids=["f1"], target_folder="Nope")
    assert r.error and "Nope" in r.text and "create_folder" in r.text and "shared with you" not in r.text
    assert graph.writes() == []


async def test_the_upload_command_quotes_the_local_path(graph, invoke):
    graph.on("GET", "/me/drive/root:/Reports", {"id": "P1", "name": "Reports", "folder": {}})
    r = await invoke("upload_file", folder="Reports", name="a b.txt")
    assert "-T '<path of the local file>' '" in r.text


async def test_an_item_not_found_without_a_drive_hints_at_the_drive_of_a_shared_listing(graph, invoke):
    # Round 4 (F06): B had the id from a shared listing and not the drive, and read "not shared with you".
    graph.on("GET", "/me/drive/items/x1", {"error": {"code": "itemNotFound"}}, status=404)
    for tool_name, args in (("get_file", {"id": "x1"}), ("download_file", {"id": "x1"})):
        r = await invoke(tool_name, **args)
        assert r.error and "drive" in r.text and "list_files" in r.text, (tool_name, r.text)
    graph.on("GET", "/drives/b!other/items/x1", {"error": {"code": "itemNotFound"}}, status=404)
    r = await invoke("get_file", id="x1", drive="b!other")
    assert r.error and "pass its drive" not in r.text


async def test_the_link_confirmation_says_the_link_is_for_the_organisation_only(graph, invoke):
    # Round 4 (F07): the refusal gave no hint of the reach, and a careless model could promise "anyone".
    r = await invoke("share_file", id="f1", action="link")
    assert r.error and "organisation" in r.text and "user_confirmed" in r.text and graph.writes() == []


async def test_the_invite_confirmation_shows_who_gets_what_and_the_message(graph, invoke):
    # Round 4 (X03): the person confirms what goes out, so the refusal repeats the note and the role.
    person(graph, "Nick", "nick@x.test")
    r = await invoke("share_file", id="f1", action="invite", people=["Nick"], role="edit", message="Please review")
    assert r.error and "nick@x.test" in r.text and "Please review" in r.text and "edit" in r.text
    assert graph.writes() == []


async def test_the_description_says_who_created_means_and_that_kind_site_needs_no_location(server):
    tools = {t.name: t for t in await server.list_tools()}
    text = tools["list_files"].description + str(tools["list_files"].input_schema)
    assert "not necessarily who shared" in text and "modified_since" in text
    assert "kind=site" in str(tools["list_files"].input_schema["properties"]["location"])


# -- a delete names who loses access ------------------------------------------

async def test_deleting_a_shared_file_names_the_people_who_lose_access(graph, invoke):
    graph.on("GET", "/me/drive/items/f1", drive_item("f1", "plan.docx"))
    graph.on("GET", "/me/drive/items/f1/permissions", PERMS)
    graph.on("DELETE", "/me/drive/items/f1", None, status=204)
    r = await invoke("update_file", action="delete", ids=["f1"])
    assert not r.error, r.text
    assert r.data["shared_with"] == {"plan.docx (f1)": ["Bo <bo@x.test>", "Cy <cy@x.test>", "view link, users",
                                                   "Di <di@x.test>", "ed@x.test"]}
    assert "no longer reach" in r.data["shared_note"]
    assert [s.method for s in graph.writes()] == ["DELETE"]


async def test_deleting_an_unshared_file_has_no_sharing_note(graph, invoke):
    graph.on("GET", "/me/drive/items/f1", drive_item("f1", "old.docx"))
    graph.on("GET", "/me/drive/items/f1/permissions", {"value": [PERMS["value"][0]]})
    graph.on("DELETE", "/me/drive/items/f1", None, status=204)
    r = await invoke("update_file", action="delete", ids=["f1"])
    assert "shared_with" not in r.data and "shared_note" not in r.data


async def test_a_delete_still_happens_when_the_permissions_cannot_be_read(graph, invoke):
    graph.on("GET", "/me/drive/items/f1", drive_item("f1", "old.docx"))
    graph.on("GET", "/me/drive/items/f1/permissions", {"error": {"code": "accessDenied", "message": "no"}}, status=403)
    graph.on("DELETE", "/me/drive/items/f1", None, status=204)
    r = await invoke("update_file", action="delete", ids=["f1"])
    assert not r.error, r.text
    assert [s.method for s in graph.writes()] == ["DELETE"]
    assert "could not read who had access to: old.docx (f1)" in r.data["shared_note"] and "shared_with" not in r.data


async def test_a_throttled_permissions_read_is_not_swallowed(graph, invoke):
    graph.on("GET", "/me/drive/items/f1", drive_item("f1", "old.docx"))
    graph.on("GET", "/me/drive/items/f1/permissions", {"error": {"code": "serviceNotAvailable", "message": "x"}}, status=503)
    graph.on("DELETE", "/me/drive/items/f1", None, status=204)
    r = await invoke("update_file", action="delete", ids=["f1"])
    assert r.error and graph.writes() == []


async def test_two_files_with_one_name_stay_two_entries(graph, invoke):
    graph.on("GET", "/me/drive/items/f1", drive_item("f1", "plan.docx"))
    graph.on("GET", "/me/drive/items/f2", drive_item("f2", "plan.docx"))
    graph.on("GET", "/me/drive/items/f[12]/permissions", PERMS)
    graph.on("DELETE", "/me/drive/items/f[12]", None, status=204)
    r = await invoke("update_file", action="delete", ids=["f1", "f2"])
    assert set(r.data["shared_with"]) == {"plan.docx (f1)", "plan.docx (f2)"}


async def test_a_refused_folder_delete_reads_no_permissions(graph, invoke):
    graph.on("GET", "/me/drive/items/d1", drive_item("d1", "Projects", folder=True))
    graph.on("GET", "/me/drive/items/d1/permissions", PERMS)
    r = await invoke("update_file", action="delete", ids=["d1"])
    assert r.error and not graph.calls("GET", "/me/drive/items/d1/permissions")


async def test_a_download_result_is_short_enough_to_cost_little_context(graph, invoke):
    # Each signed link was ~6,300 characters, ~7k per download_file result and 70k
    # for ten attachment links. The link now carries a short id; the whole result stays small.
    graph.on("GET", "/me/drive/items/f1", drive_item("f1", "a.docx"))
    r = await invoke("download_file", id="f1")
    assert len(link_of(r.text, "d")) < 40 and len(r.text) < 1500, len(r.text)


async def test_a_full_link_store_refuses_the_download_with_a_sentence(graph, invoke, rt):
    from m365_mcp.links.store import LinkStore
    from m365_mcp.links.tokens import LinkSigner
    graph.on("GET", "/me/drive/items/f1", drive_item("f1", "a.docx"))
    rt.signer = LinkSigner((b"s" * 32,), (b"e" * 32,), LinkStore(max_entries=0))
    r = await invoke("download_file", id="f1")
    assert r.error and "Too many file links" in r.text


async def test_download_file_is_annotated_as_reading_only(server):
    # destructive_hint true on a tool that changes nothing, so clients asked
    # for confirmation without need.
    tools = {t.name: t for t in await server.list_tools()}
    a = tools["download_file"].annotations
    assert a.read_only_hint is True and a.destructive_hint is False


async def test_kind_document_leaves_out_plain_text(graph, invoke):
    # kind=document included .txt, and the docx was on the second page.
    from m365_mcp.tools.files import _KIND_BY_EXT
    assert "txt" not in _KIND_BY_EXT["document"] and "docx" in _KIND_BY_EXT["document"]
