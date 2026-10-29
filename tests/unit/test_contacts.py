import httpx
import pytest


def _graph_applies_writes(graph, method, path, base, status=200):
    """Graph answers a write with the contact as it now stands: what it held, with the sent fields on top."""
    graph.on(method, path, lambda seen: httpx.Response(status, json={**base, **seen.json}))


@pytest.fixture(autouse=True)
def _folders(graph):
    # The default folder, and Clients with Key inside it. Deleted Items is a mail folder, not one of these.
    graph.on("GET", "/me/contactFolders/contacts", {"id": "contacts", "displayName": "Contacts"})
    graph.on("GET", "/me/contactFolders/contacts/childFolders", {"value": []})
    graph.on("GET", "/me/contactFolders", {"value": [{"id": "f-clients", "displayName": "Clients"}]})
    graph.on("GET", "/me/contactFolders/f-clients/childFolders", {"value": [{"id": "f-key", "displayName": "Key"}]})
    graph.on("GET", "/me/contactFolders/f-key/childFolders", {"value": []})


async def test_create_splits_the_name_and_sends_structured_fields(graph, invoke):
    _graph_applies_writes(graph, "POST", "/me/contacts", {"id": "c1", "emailAddresses": [], "businessPhones": []}, 201)
    r = await invoke("save_contact", name="Anna Maria Berg", emails=["anna@x.test"], phones=["+4712345678"],
                     company="Acme", job_title="CFO", notes="met at fair")
    assert not r.error and r.data["id"] == "c1", r.text
    assert graph.writes()[0].json == {
        "displayName": "Anna Maria Berg", "givenName": "Anna Maria", "surname": "Berg",
        "emailAddresses": [{"address": "anna@x.test", "name": "Anna Maria Berg"}],
        "businessPhones": ["+4712345678"], "companyName": "Acme", "jobTitle": "CFO", "personalNotes": "met at fair"}


async def test_update_reads_first_and_sends_the_merged_lists_whole(graph, invoke):
    graph.on("GET", "/me/contacts/c1", {"id": "c1", "displayName": "Anna Berg", "parentFolderId": "contacts",
                                         "emailAddresses": [{"address": "a@x.test", "name": "Anna"}],
                                         "businessPhones": ["111"]})
    graph.on("PATCH", "/me/contacts/c1", {"id": "c1"})
    r = await invoke("save_contact", id="c1", emails=["A@x.test", "b@x.test"], phones=["111", "222"])
    assert not r.error, r.text
    assert graph.calls("GET", "/me/contacts/c1")[0].params["$select"] == "displayName,emailAddresses,businessPhones,homePhones,mobilePhone,parentFolderId"
    assert graph.writes()[0].json == {
        "emailAddresses": [{"address": "a@x.test", "name": "Anna"}, {"address": "b@x.test", "name": "Anna Berg"}],
        "businessPhones": ["111", "222"]}


@pytest.mark.parametrize("tool, extra", [("save_contact", {"emails": ["x@x.test"]}), ("delete_contact", {"user_confirmed": True})])
@pytest.mark.parametrize("ident", ["AAD-user-id", "bo@x.test"])
async def test_a_directory_person_is_not_a_contact(graph, invoke, tool, extra, ident):
    # A colleague's directory id or address is no id under /me/contacts.
    r = await invoke(tool, id=ident, **extra)
    assert r.error and "Only personal contacts are changed here" in r.text and "list_people" in r.text
    assert graph.writes() == []


async def test_a_new_contact_needs_a_name(graph, invoke):
    r = await invoke("save_contact", emails=["x@x.test"])
    assert r.error and "needs a name" in r.text and graph.writes() == []


async def test_delete_needs_confirmation(graph, invoke):
    r = await invoke("delete_contact", id="c1")
    assert r.error and "user_confirmed" in r.text and graph.writes() == []
    graph.on("GET", "/me/contacts/c1", {"id": "c1", "parentFolderId": "contacts"})
    graph.on("DELETE", "/me/contacts/c1", None, status=204)
    assert not (await invoke("delete_contact", id="c1", user_confirmed=True)).error
    assert graph.writes()[0].method == "DELETE"


def _deleted_contact(graph):
    # Measured: a deleted contact keeps its id and answers 200 from Deleted Items; PATCH and a second
    # DELETE succeed on it, and the second DELETE takes it out of Deleted Items too.
    graph.on("GET", "/me/contacts/c1", {"id": "c1", "displayName": "Ann", "parentFolderId": "del-folder"})
    graph.on("PATCH", "/me/contacts/c1", {})
    graph.on("DELETE", "/me/contacts/c1", None, status=204)


@pytest.mark.parametrize("tool, extra", [("save_contact", {"job_title": "CFO"}), ("delete_contact", {"user_confirmed": True})])
async def test_a_contact_in_deleted_items_is_not_changed_or_deleted_again(graph, invoke, tool, extra):
    _deleted_contact(graph)
    r = await invoke(tool, id="c1", **extra)
    assert r.error and "Deleted Items" in r.text and "restore" in r.text and "Nothing was changed" in r.text
    assert graph.writes() == []


@pytest.mark.parametrize("bad", ["not-an-email", "a@b", "two@@x.test", "a b@x.test"])
async def test_an_email_that_is_not_an_address_is_refused(graph, invoke, bad):
    r = await invoke("save_contact", name="Ann", emails=[bad])
    assert r.error and "not an email address" in r.text and graph.writes() == []


async def test_a_fourth_email_names_outlooks_limit_of_three(graph, invoke):
    graph.on("GET", "/me/contacts/c1", {"id": "c1", "displayName": "Ann", "parentFolderId": "contacts",
                                         "emailAddresses": [{"address": f"a{i}@x.test"} for i in range(3)]})
    r = await invoke("save_contact", id="c1", emails=["d@x.test"])
    assert r.error and "at most three email addresses" in r.text and graph.writes() == []


async def test_a_new_contact_goes_into_the_folder_named(graph, invoke):
    _graph_applies_writes(graph, "POST", "/me/contactFolders/f-clients/contacts", {"id": "c9", "emailAddresses": [], "businessPhones": []}, 201)
    _graph_applies_writes(graph, "POST", "/me/contactFolders/f-key/contacts", {"id": "c10", "emailAddresses": [], "businessPhones": []}, 201)
    r = await invoke("save_contact", name="Ann", folder="clients")
    assert not r.error, r.text
    nested = await invoke("save_contact", name="Bo", folder="Clients/Key")
    assert not nested.error, nested.text
    assert [w.path for w in graph.writes()] == ["/me/contactFolders/f-clients/contacts", "/me/contactFolders/f-key/contacts"]
    unknown = await invoke("save_contact", name="Ann", folder="Nope")
    assert unknown.error and "Nope" in unknown.text


async def test_a_contact_in_a_subfolder_is_changed_and_no_mail_permission_is_asked(graph, invoke):
    graph.on("GET", "/me/contacts/c1", {"id": "c1", "displayName": "Ann", "parentFolderId": "f-key"})
    graph.on("PATCH", "/me/contacts/c1", {})
    r = await invoke("save_contact", id="c1", job_title="CFO")
    assert not r.error, r.text
    assert not [s for s in graph.seen if "/mailFolders" in s.path]


async def test_a_contact_in_a_folder_under_the_default_one_is_not_taken_for_deleted(graph, invoke):
    graph.on("GET", "/me/contactFolders/contacts/childFolders", {"value": [{"id": "f-under", "displayName": "Under"}]})
    graph.on("GET", "/me/contactFolders/f-under/childFolders", {"value": []})
    graph.on("GET", "/me/contacts/c1", {"id": "c1", "displayName": "Ann", "parentFolderId": "f-under"})
    graph.on("PATCH", "/me/contacts/c1", {})
    r = await invoke("save_contact", id="c1", job_title="CFO")
    assert not r.error, r.text


async def test_a_contact_whose_folder_cannot_be_placed_is_not_changed(graph, invoke, monkeypatch):
    # The folders could not all be read: a contact in none of those read may be deleted, so neither
    # tool writes rather than DELETE a second time from Deleted Items.
    import m365_mcp.tools.contacts as contacts
    monkeypatch.setattr(contacts, "MAX_FOLDER_REQUESTS", 1)
    _deleted_contact(graph)
    for tool, extra in (("save_contact", {"job_title": "CFO"}), ("delete_contact", {"user_confirmed": True})):
        r = await invoke(tool, id="c1", **extra)
        assert r.error and "cannot be told from a deleted one" in r.text and "Nothing was changed" in r.text
    assert graph.writes() == []


async def test_one_address_given_twice_counts_once_against_the_limit(graph, invoke):
    graph.on("GET", "/me/contacts/c1", {"id": "c1", "displayName": "Ann", "parentFolderId": "contacts",
                                         "emailAddresses": [{"address": f"a{i}@x.test"} for i in range(2)]})
    graph.on("PATCH", "/me/contacts/c1", {})
    r = await invoke("save_contact", id="c1", emails=["new@x.test", "NEW@x.test"])
    assert not r.error, r.text
    assert [e["address"] for e in graph.writes()[0].json["emailAddresses"]] == ["a0@x.test", "a1@x.test", "new@x.test"]


async def test_a_contact_without_a_folder_is_not_changed(graph, invoke):
    graph.on("GET", "/me/contacts/c1", {"id": "c1", "displayName": "Ann"})
    graph.on("PATCH", "/me/contacts/c1", {})
    r = await invoke("save_contact", id="c1", job_title="CFO")
    assert r.error and "cannot be told from a deleted one" in r.text and graph.writes() == []


def _anna(graph, **extra):
    graph.on("GET", "/me/contacts/c1", {"id": "c1", "displayName": "Anna Berg", "parentFolderId": "contacts",
                                         "emailAddresses": [{"address": "a@x.test", "name": "Anna"}],
                                         "businessPhones": ["+47 111 22"], **extra})
    _graph_applies_writes(graph, "PATCH", "/me/contacts/c1", {
        "id": "c1", "displayName": "Anna Berg", "emailAddresses": [{"address": "a@x.test", "name": "Anna"}],
        "businessPhones": ["+47 111 22"], **extra})


async def test_a_phone_is_replaced_by_giving_the_new_and_the_old(graph, invoke):
    # Phones were only ever added, so "change the phone" left two numbers.
    _anna(graph)
    r = await invoke("save_contact", id="c1", phones=["+47 999"], remove_phones=["+4711122"])
    assert not r.error, r.text
    assert graph.writes()[0].json == {"businessPhones": ["+47 999"]}
    assert r.data["phones"] == ["+47 999"] and "not_removed" not in r.data


async def test_the_result_says_which_numbers_the_contact_holds_now(graph, invoke):
    _anna(graph, mobilePhone="555")
    r = await invoke("save_contact", id="c1", phones=["222"])
    assert r.data["phones"] == ["+47 111 22", "222", "555"]


async def test_a_number_held_as_mobile_can_be_removed(graph, invoke):
    _anna(graph, mobilePhone="555", homePhones=["777"])
    r = await invoke("save_contact", id="c1", remove_phones=["555", "777"])
    assert not r.error, r.text
    assert graph.writes()[0].json == {"homePhones": [], "mobilePhone": None, "businessPhones": ["+47 111 22"]}
    assert r.data["phones"] == ["+47 111 22"]


async def test_an_email_is_replaced_and_one_that_is_not_held_is_named(graph, invoke):
    _anna(graph)
    r = await invoke("save_contact", id="c1", emails=["new@x.test"], remove_emails=["A@x.test", "ghost@x.test"])
    assert not r.error, r.text
    assert graph.writes()[0].json == {"emailAddresses": [{"address": "new@x.test", "name": "Anna Berg"}]}
    assert r.data["emails"] == ["new@x.test"] and r.data["not_removed"] == ["ghost@x.test"]


async def test_removing_on_a_new_contact_is_refused(graph, invoke):
    r = await invoke("save_contact", name="Bo", remove_phones=["1"])
    assert r.error and "existing contact" in r.text and graph.writes() == []


async def test_the_result_shows_what_microsoft_holds_not_what_was_asked(graph, invoke):
    # Round 4 (P03, P02), review: the lists come from Graph's answer to the write. The mocked answer
    # differs from the request (Outlook rewrote the address, kept a number, and holds an address with
    # no value), so a result echoing the request fails here.
    graph.on("GET", "/me/contacts/c1", {"id": "c1", "displayName": "Anna Berg", "parentFolderId": "contacts",
                                         "emailAddresses": [{"address": "a@x.test", "name": "Anna"}],
                                         "businessPhones": ["111"]})
    graph.on("PATCH", "/me/contacts/c1", {
        "id": "c1", "displayName": "Anna B. Berg",
        "emailAddresses": [{"address": "anna@kept.test"}, {"name": "no address"}],
        "businessPhones": ["111", "222"], "mobilePhone": "333"})
    r = await invoke("save_contact", id="c1", phones=["999"])
    assert not r.error, r.text
    assert r.data["name"] == "Anna B. Berg" and r.data["emails"] == ["anna@kept.test"]
    assert r.data["phones"] == ["111", "222", "333"]


async def test_when_the_write_answers_without_the_lists_they_are_read_back(graph, invoke):
    graph.on("POST", "/me/contacts", {"id": "c2"}, status=201)
    graph.on("GET", "/me/contacts/c2", {"id": "c2", "displayName": "Bo Lund",
                                         "emailAddresses": [{"address": "bo@kept.test"}], "businessPhones": []})
    made = await invoke("save_contact", name="Bo Lund", emails=["BO@x.test"])
    assert not made.error, made.text
    assert made.data["id"] == "c2" and made.data["emails"] == ["bo@kept.test"] and made.data["phones"] == []
    assert graph.calls("GET", "/me/contacts/c2")[0].params["$select"].startswith("displayName,emailAddresses")
