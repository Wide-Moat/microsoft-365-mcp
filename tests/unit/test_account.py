from helpers import make_token

from m365_mcp.identity import current_token


async def test_get_connection_reports_person_zone_and_areas(graph, invoke):
    graph.on("GET", "/me", {"displayName": "Ann Lee", "mail": "ann@contoso.test", "jobTitle": "PM"})
    graph.on("GET", "/me/mailboxSettings", {"workingHours": {"daysOfWeek": ["monday"], "startTime": "09:00:00.0000000",
                                                             "endTime": "17:00:00.0000000"},
                                            "automaticRepliesSetting": {"status": "scheduled"}})
    r = await invoke("get_connection")
    assert not r.error, r.text
    assert r.data["name"] == "Ann Lee" and r.data["address"] == "ann@contoso.test"
    assert r.data["working_hours"]["start"] == "09:00"
    assert r.data["automatic_reply"] == "scheduled"
    assert r.data["areas"]["mail"] == "on"


async def test_get_connection_names_a_missing_permission(graph, invoke):
    current_token.set(make_token("User.Read MailboxSettings.Read Mail.ReadWrite"))
    graph.on("GET", "/me", {"displayName": "Ann"})
    graph.on("GET", "/me/mailboxSettings", {})
    r = await invoke("get_connection")
    assert "Mail.Send" in r.data["areas"]["mail"]
    assert r.data["areas"]["calendar"].startswith("missing permission")


async def test_no_token_means_reconnect_and_no_graph_call(graph, invoke):
    current_token.set(None)
    r = await invoke("get_connection")
    assert r.error and "Reconnect" in r.text
    assert graph.seen == []


async def test_expired_token_means_reconnect(graph, invoke):
    current_token.set(make_token(exp=1000))
    r = await invoke("get_connection")
    assert r.error and "Reconnect" in r.text
    assert graph.seen == []


async def test_wrong_audience_means_reconnect(graph, invoke):
    current_token.set(make_token(aud="api://something-else"))
    r = await invoke("get_connection")
    assert r.error and "Reconnect" in r.text


async def test_graph_401_means_reconnect(graph, invoke):
    graph.on("GET", "/me", {"error": {"code": "InvalidAuthenticationToken", "message": "x"}}, status=401)
    r = await invoke("get_connection")
    assert r.error and "Reconnect" in r.text


async def test_every_request_carries_immutable_ids_and_a_request_id(graph, invoke):
    graph.on("GET", "/me", {"displayName": "Ann"})
    graph.on("GET", "/me/mailboxSettings", {})
    await invoke("get_connection")
    assert graph.seen
    for s in graph.seen:
        assert 'IdType="ImmutableId"' in s.headers["prefer"]
        assert s.headers["client-request-id"]
        assert s.headers["authorization"].startswith("Bearer ")


async def test_list_people_ranks_and_marks_groups(graph, invoke):
    graph.on("GET", "/me/people", {"value": [
        {"displayName": "Anna Berg", "scoredEmailAddresses": [{"address": "anna@contoso.test"}],
         "personType": {"class": "Person", "subclass": "OrganizationUser"}, "jobTitle": "CFO"},
        {"displayName": "Anna Team", "scoredEmailAddresses": [{"address": "annateam@contoso.test"}],
         "personType": {"class": "Group", "subclass": "PublicDistributionList"}},
        {"displayName": "Anna Accountant", "scoredEmailAddresses": [{"address": "acc@example.org"}],
         "personType": {"class": "Person", "subclass": "PersonalContact"}},
    ]})
    r = await invoke("list_people", query="Anna")
    assert not r.error, r.text
    items = r.data["items"]
    assert items[0] == {"name": "Anna Berg", "address": "anna@contoso.test", "job_title": "CFO", "source": "directory"}
    assert items[1]["group"] is True
    assert items[2]["source"] == "contacts"
    r = await invoke("list_people", query="Anna", source="contacts")
    assert [i["address"] for i in r.data["items"]] == ["acc@example.org"]


async def test_list_people_refuses_an_empty_query(graph, invoke):
    r = await invoke("list_people", query="  ")
    assert r.error and "Give a name or an address" in r.text
    assert graph.calls("GET", "/me/people") == []


async def test_list_people_clamps_the_limit(graph, invoke):
    graph.on("GET", "/me/people", {"value": []})
    graph.on("GET", "/users", {"value": []})
    r = await invoke("list_people", query="x", limit=500)
    assert r.data["limit_clamped_to"] == 25
    assert graph.calls("GET", "/me/people")[0].params["$top"] == "26"  # one more, to know has_more


async def test_a_personal_contact_is_found_at_once_without_the_people_index(graph, invoke):
    # The People index learns a new contact late or never; contacts are read where they are kept.
    graph.on("GET", "/me/people", {"value": []})
    graph.on("GET", "/users", {"value": []})
    graph.on("GET", "/me/contacts", {"value": [
        {"id": "c1", "displayName": "qa Contact One", "companyName": "Ledger Ltd", "jobTitle": "Accountant",
         "emailAddresses": [{"address": "one@example.org"}]},
        {"id": "c2", "displayName": "Someone Else", "emailAddresses": [{"address": "else@example.org"}]}]})
    for q in ("Contact One", "Ledger", "accountant", "one@example.org"):
        for source in ("contacts", "all"):
            r = await invoke("list_people", query=q, source=source)
            assert not r.error, r.text
            assert [i["address"] for i in r.data["items"]] == ["one@example.org"], (q, source, r.data)
            assert r.data["items"][0]["source"] == "contacts"


async def test_a_personal_contact_comes_with_the_id_save_and_delete_take(graph, invoke):
    # Only the creating call returned the id, so an older contact could be
    # neither changed nor deleted.
    graph.on("GET", "/me/people", {"value": []})
    graph.on("GET", "/me/contacts", {"value": [
        {"id": "c1", "displayName": "Ann Ledger", "emailAddresses": [{"address": "ann@example.org"}]}]})
    graph.on("GET", "/users", {"value": [{"id": "u-ann", "displayName": "Ann Lee", "mail": "ann.lee@x.test",
                                           "userPrincipalName": "ann.lee@x.test"}]})
    r = await invoke("list_people", query="Ann")
    contact = next(i for i in r.data["items"] if i["source"] == "contacts")
    assert contact["contact_id"] == "c1"
    assert all("contact_id" not in i for i in r.data["items"] if i["source"] != "contacts")
    graph.on("GET", "/me/contacts/c1", {"id": "c1", "displayName": "Ann Ledger", "parentFolderId": "contacts"})
    graph.on("GET", "/me/contactFolders/contacts", {"id": "contacts", "displayName": "Contacts"})
    graph.on("GET", "/me/contactFolders", {"value": []})
    graph.on("GET", "/me/contactFolders/contacts/childFolders", {"value": []})
    graph.on("PATCH", "/me/contacts/c1", {})
    saved = await invoke("save_contact", id=contact["contact_id"], job_title="Accountant")
    assert not saved.error and [w.path for w in graph.writes()] == ["/me/contacts/c1"]


async def test_a_contact_found_by_a_second_address_is_given_that_address(graph, invoke):
    graph.on("GET", "/me/contacts", {"value": [{"id": "c1", "displayName": "Mia", "emailAddresses": [
        {"address": "mia@home.example"}, {"address": "mia@work.example"}]}]})
    r = await invoke("list_people", query="mia@work.example", source="contacts")
    assert [i["address"] for i in r.data["items"]] == ["mia@work.example"]


async def test_more_matches_than_the_limit_say_there_are_more(graph, invoke):
    graph.on("GET", "/me/contacts", {"value": [
        {"id": f"c{i}", "displayName": f"Ann {i}", "emailAddresses": [{"address": f"ann{i}@example.org"}]} for i in range(5)]})
    r = await invoke("list_people", query="Ann", source="contacts", limit=2)
    assert len(r.data["items"]) == 2 and r.data["has_more"] is True and r.data["cursor"]


async def test_list_people_in_the_directory_only_reads_no_contacts(graph, invoke):
    r = await invoke("list_people", query="Ann", source="directory")
    assert not r.error, r.text
    assert graph.calls("GET", "/me/contacts") == []


async def test_a_contacts_search_past_its_pages_says_the_list_was_cut(graph, invoke):
    graph.on("GET", "/me/contacts", {"value": [{"id": "c1", "displayName": "Ann", "emailAddresses": [{"address": "ann@example.org"}]}],
                                     "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/contacts?$skip=100"})
    r = await invoke("list_people", query="Ann", source="contacts")
    assert r.data["truncated"] is True


async def test_a_contact_holding_a_group_known_to_people_is_shown_as_a_group(graph, invoke):
    graph.on("GET", "/me/contacts", {"value": [{"id": "c1", "displayName": "All Staff", "emailAddresses": [{"address": "all@x.test"}]}]})
    graph.on("GET", "/me/people", {"value": [{"displayName": "All Staff", "scoredEmailAddresses": [{"address": "all@x.test"}],
                                              "personType": {"class": "Group", "subclass": "PublicDistributionList"}}]})
    for source in ("all", "contacts"):
        r = await invoke("list_people", query="All Staff", source=source)
        assert [i.get("group") for i in r.data["items"]] == [True], source


async def test_contacts_read_before_the_budget_ran_out_are_still_listed(graph, invoke):
    graph.on("GET", "/me/contacts", {"value": [{"id": "c1", "displayName": "Ann", "emailAddresses": [{"address": "ann@example.org"}]}],
                                     "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/contacts?$skip=100"})
    r = await invoke("list_people", query="Ann", source="contacts")
    assert r.data["items"] and r.data["items"][0]["address"] == "ann@example.org" and r.data["truncated"] is True


async def test_one_person_saved_twice_is_listed_once(graph, invoke):
    graph.on("GET", "/me/contacts", {"value": [{"id": "c1", "displayName": "Ann", "emailAddresses": [{"address": "ann@example.org"}]}]})
    graph.on("GET", "/me/contactFolders", {"value": [{"id": "F1"}]})
    graph.on("GET", "/me/contactFolders/F1/contacts", {"value": [{"id": "c2", "displayName": "Ann", "emailAddresses": [{"address": "ANN@example.org"}]}]})
    graph.on("GET", "/me/contactFolders/F1/childFolders", {"value": []})
    r = await invoke("list_people", query="Ann", source="contacts")
    assert len(r.data["items"]) == 1


async def test_get_connection_says_when_no_working_hours_are_set(graph, invoke):
    # Measured: both test users' mailboxSettings answer "workingHours": {}; find_meeting_times then uses
    # its documented default, and get_connection says the same thing rather than leaving the key out.
    graph.on("GET", "/me", {"displayName": "Ann"})
    graph.on("GET", "/me/mailboxSettings", {"workingHours": {}, "automaticRepliesSetting": {"status": "disabled"}})
    r = await invoke("get_connection")
    assert not r.error, r.text
    assert r.data["working_hours"] == "not set in Outlook"



async def test_list_people_continues_with_a_cursor(graph, invoke):
    # Every list carries has_more and cursor.
    import httpx

    def people(seen):
        top = int(seen.params["$top"])
        return httpx.Response(200, json={"value": [{
            "displayName": f"Ann {i}", "scoredEmailAddresses": [{"address": f"ann{i}@x.test"}],
            "personType": {"class": "Person", "subclass": "OrganizationUser"}} for i in range(min(top, 5))]})
    graph.on("GET", "/me/people", people)
    first = await invoke("list_people", query="Ann", limit=2)
    assert [i["address"] for i in first.data["items"]] == ["ann0@x.test", "ann1@x.test"]
    assert first.data["has_more"] is True
    second = await invoke("list_people", query="Ann", limit=2, cursor=first.data["cursor"])
    assert [i["address"] for i in second.data["items"]] == ["ann2@x.test", "ann3@x.test"]
    last = await invoke("list_people", query="Ann", limit=2, cursor=second.data["cursor"])
    assert [i["address"] for i in last.data["items"]] == ["ann4@x.test"] and last.data["has_more"] is False
    other = await invoke("list_people", query="Bob", limit=2, cursor=first.data["cursor"])
    assert other.error and "without the cursor" in other.text



async def test_list_people_reads_on_past_what_the_source_filter_drops(graph, invoke):
    import httpx

    def people(seen):
        top = int(seen.params["$top"])
        rows = [{"displayName": f"Ann {i}", "scoredEmailAddresses": [{"address": f"ann{i}@x.test"}],
                 "personType": {"class": "Person", "subclass": "PersonalContact" if i < 3 else "OrganizationUser"}}
                for i in range(10)]
        return httpx.Response(200, json={"value": rows[:top]})
    graph.on("GET", "/me/people", people)
    r = await invoke("list_people", query="Ann", source="directory", limit=2)
    assert [i["address"] for i in r.data["items"]] == ["ann3@x.test", "ann4@x.test"] and r.data["has_more"] is True


async def test_list_people_ends_its_last_page_at_the_cap(graph, invoke, rt):
    import httpx
    from m365_mcp.tools.account import MAX_PEOPLE
    graph.on("GET", "/me/people", lambda seen: httpx.Response(200, json={"value": [{
        "displayName": f"Ann {i}", "scoredEmailAddresses": [{"address": f"ann{i}@x.test"}],
        "personType": {"class": "Person", "subclass": "OrganizationUser"}} for i in range(int(seen.params["$top"]))]}))
    from m365_mcp.identity import current_person
    cursor = rt.cursors.make("list_people", current_person().key, {"args": ["Ann", "all", False], "skip": MAX_PEOPLE - 5})
    r = await invoke("list_people", query="Ann", limit=10, cursor=cursor)
    assert not r.error, r.text
    assert r.data["count"] == 5 and r.data["has_more"] is False and r.data["truncated"] is True


async def test_a_contact_people_also_ranks_counts_once(graph, invoke):
    # Review of this branch: six saved contacts that People ranks too were counted twice, so the
    # read stopped before the contacts People holds further down.
    import httpx
    graph.on("GET", "/me/contacts", {"value": [
        {"id": f"c{i}", "displayName": f"Ann {i}", "emailAddresses": [{"address": f"ann{i}@x.test"}]} for i in range(6)]})

    def people(seen):
        kinds = ["PersonalContact"] * 6 + ["OrganizationUser"] * 5 + ["PersonalContact"] * 4
        rows = [{"displayName": f"Ann {i}", "scoredEmailAddresses": [{"address": f"ann{i}@x.test"}],
                 "personType": {"class": "Person", "subclass": k}} for i, k in enumerate(kinds)]
        return httpx.Response(200, json={"value": rows[:int(seen.params["$top"])]})
    graph.on("GET", "/me/people", people)
    r = await invoke("list_people", query="Ann", source="contacts", limit=10)
    assert [i["address"] for i in r.data["items"]] == [f"ann{i}@x.test" for i in (*range(6), *range(11, 15))]
    assert r.data["has_more"] is False


async def test_get_connection_shows_when_a_scheduled_reply_runs(graph, invoke):
    # "scheduled" alone could not remind the person when the reply runs.
    graph.on("GET", "/me", {"displayName": "Ann"})
    graph.on("GET", "/me/mailboxSettings", {"workingHours": {}, "automaticRepliesSetting": {
        "status": "scheduled",
        "scheduledStartDateTime": {"dateTime": "2026-10-30T08:00:00.0000000", "timeZone": "UTC"},
        "scheduledEndDateTime": {"dateTime": "2026-11-02T08:00:00.0000000", "timeZone": "UTC"}}})
    r = await invoke("get_connection")
    assert r.data["automatic_reply"] == "scheduled"
    assert r.data["automatic_reply_period"] == {"from": "2026-10-30T08:00+00:00", "until": "2026-11-02T08:00+00:00"}


async def test_get_connection_has_no_period_when_the_reply_is_off(graph, invoke):
    graph.on("GET", "/me", {"displayName": "Ann"})
    graph.on("GET", "/me/mailboxSettings", {"workingHours": {}, "automaticRepliesSetting": {"status": "disabled"}})
    assert "automatic_reply_period" not in (await invoke("get_connection")).data


async def test_list_people_with_no_match_says_what_was_searched(graph, invoke):
    # Round 4 (P05): a bare {items: [], count: 0} did not say whether the directory was searched.
    graph.on("GET", "/me/people", {"value": []})
    graph.on("GET", "/users", {"value": []})
    r = await invoke("list_people", query="Nobody")
    assert not r.error, r.text
    assert r.data["count"] == 0 and "contacts and the organisation's directory" in r.data["note"]


async def test_an_incomplete_contact_search_does_not_claim_nobody_matched(graph, invoke, monkeypatch):
    from m365_mcp.core.resolve import Contacts
    from m365_mcp.tools import account

    async def partial(c, query):
        out = Contacts()
        out.complete = False
        return out
    monkeypatch.setattr(account, "contacts", partial)
    graph.on("GET", "/me/people", {"value": []})
    graph.on("GET", "/users", {"value": []})
    r = await invoke("list_people", query="Nobody")
    assert not r.error, r.text
    assert r.data["truncated"] is True and "note" not in r.data and "Not every contact" in r.data["contacts_note"]
