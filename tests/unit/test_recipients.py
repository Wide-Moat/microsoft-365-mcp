"""Who a message, an invitation or a share reaches: the shared resolver, through every tool that uses it.

Two defects found on a real tenant hold this file together. A group address (a Microsoft 365 group,
a distribution list) is not in the People index, so it passed as a person and a draft to it was sent.
And a first name that two colleagues share resolved to the one whose display name it equals,
without a word about the other.
"""

from __future__ import annotations

import re

import httpx
import pytest

from m365_mcp.tools.mail import DRAFT_MARK

from people import me

OWN = {"singleValueExtendedProperties": [{"id": DRAFT_MARK, "value": "1"}]}
GROUP = "wm-all@x.test"


def directory(graph, users: dict[str, str], domains=("x.test",)) -> None:
    """A directory holding exactly `users` (address -> display name) on the organisation's `domains`."""
    def route(seen):
        f = seen.params.get("$filter", "")
        m = re.search(r"mail eq '((?:[^']|'')*)'", f)
        if not m:
            return httpx.Response(200, json={"value": []})
        addr = m.group(1).replace("''", "'").casefold()
        hits = [{"displayName": n, "mail": a, "userPrincipalName": a} for a, n in users.items() if a.casefold() == addr]
        return httpx.Response(200, json={"value": hits})
    graph.on("GET", "/users", route)
    graph.on("GET", "/organization", {"value": [{"verifiedDomains": [{"name": d} for d in domains]}]})


def two_megans(graph) -> None:
    """People ranks both colleagues for the first name; one display name is exactly that name."""
    def route(seen):
        q = seen.params.get("$search", "").strip('"').casefold()
        people = [("Megan", "megan@x.test"), ("Megan Bowen", "mb@x.test")]
        hits = [{"displayName": n, "scoredEmailAddresses": [{"address": a}], "jobTitle": "Engineer",
                 "personType": {"class": "Person", "subclass": "OrganizationUser"}}
                for n, a in people if q and (q in n.casefold() or q in a)]
        return httpx.Response(200, json={"value": hits})
    graph.on("GET", "/me/people", route)


# -- groups and lists ------------------------------------------------------------------------

async def test_save_draft_refuses_an_internal_address_that_is_no_person(graph, invoke):
    directory(graph, {"bo@x.test": "Bo"})
    r = await invoke("save_draft", to=["bo@x.test", GROUP], subject="s", body="x")
    assert r.error and GROUP in r.text and "distribution list" in r.text
    assert graph.writes() == []


async def test_send_draft_refuses_a_draft_to_a_group_the_people_index_does_not_know(graph, invoke):
    directory(graph, {"bo@x.test": "Bo"})
    graph.on("GET", "/me/messages/d1", {"id": "d1", "isDraft": True, "subject": "S", **OWN,
                                        "toRecipients": [{"emailAddress": {"address": GROUP}}]})
    graph.on("POST", "/me/messages/d1/send", None, status=202)
    r = await invoke("send_draft", id="d1", user_confirmed=True)
    assert r.error and GROUP in r.text and "distribution list" in r.text
    assert graph.writes() == []


async def test_an_invitation_to_a_group_is_refused(graph, invoke):
    directory(graph, {})
    r = await invoke("save_event", subject="Sync", start="2026-10-23 15:00", duration_minutes=30,
                     attendees=[GROUP], user_confirmed=True)
    assert r.error and GROUP in r.text and graph.writes() == []


async def test_sharing_a_file_with_a_group_is_refused(graph, invoke):
    directory(graph, {})
    r = await invoke("share_file", id="f1", action="invite", people=[GROUP], user_confirmed=True)
    assert r.error and GROUP in r.text and graph.writes() == []


async def test_a_teams_message_to_a_group_is_refused(graph, invoke):
    directory(graph, {})
    me(graph)
    r = await invoke("send_message", to=[GROUP], body="hi", user_confirmed=True)
    assert r.error and GROUP in r.text and graph.writes() == []


async def test_a_directory_user_is_a_person_and_keeps_their_name(graph, invoke):
    directory(graph, {"bo@x.test": "Bo Berg"})
    graph.on("POST", "/me/messages", {"id": "d1"})
    graph.on("GET", "/me/messages/d1", {"id": "d1", "toRecipients": [{"emailAddress": {"address": "bo@x.test"}}]})
    r = await invoke("save_draft", to=["bo@x.test"], subject="s", body="x")
    assert not r.error, r.text
    assert graph.calls("POST", "/me/messages")[0].json["toRecipients"] == [
        {"emailAddress": {"address": "bo@x.test", "name": "Bo Berg"}}]


async def test_an_address_outside_the_organisation_is_taken_as_given(graph, invoke):
    directory(graph, {})
    graph.on("POST", "/me/messages", {"id": "d1"})
    graph.on("GET", "/me/messages/d1", {"id": "d1"})
    r = await invoke("save_draft", to=["client@example.org"], subject="s", body="x")
    assert not r.error, r.text


# -- one name, several people ------------------------------------------------------------------

async def test_a_shared_first_name_is_refused_in_a_draft(graph, invoke):
    two_megans(graph)
    r = await invoke("save_draft", to=["Megan"], subject="s", body="x")
    assert r.error and "megan@x.test" in r.text and "mb@x.test" in r.text
    assert graph.writes() == []


async def test_a_shared_first_name_is_refused_in_a_teams_message(graph, invoke):
    two_megans(graph)
    me(graph)
    r = await invoke("send_message", to=["Megan"], body="hi", user_confirmed=True)
    assert r.error and "megan@x.test" in r.text and "mb@x.test" in r.text
    assert graph.writes() == []


async def test_a_shared_first_name_is_refused_when_sharing(graph, invoke):
    two_megans(graph)
    r = await invoke("share_file", id="f1", action="invite", people=["Megan"], user_confirmed=True)
    assert r.error and "megan@x.test" in r.text and "mb@x.test" in r.text
    assert graph.writes() == []


async def test_a_shared_first_name_is_refused_as_an_attendee(graph, invoke):
    two_megans(graph)
    r = await invoke("save_event", subject="Sync", start="2026-10-23 15:00", duration_minutes=30,
                     attendees=["Megan"], user_confirmed=True)
    assert r.error and "megan@x.test" in r.text and "mb@x.test" in r.text
    assert graph.writes() == []


async def test_a_shared_first_name_is_refused_as_a_sender_filter(graph, invoke):
    two_megans(graph)
    graph.on("GET", "/me/messages", {"value": []})
    r = await invoke("list_messages", sender="Megan")
    assert r.error and "megan@x.test" in r.text and "mb@x.test" in r.text
    assert graph.calls("GET", "/me/messages") == []


async def test_the_candidates_carry_the_job_title(graph, invoke):
    two_megans(graph)
    r = await invoke("save_draft", to=["Megan"], subject="s", body="x")
    assert "Engineer" in r.text


async def test_a_full_name_still_resolves_to_one_person(graph, invoke):
    two_megans(graph)
    graph.on("POST", "/me/messages", {"id": "d1"})
    graph.on("GET", "/me/messages/d1", {"id": "d1"})
    r = await invoke("save_draft", to=["Megan Bowen"], subject="s", body="x")
    assert not r.error, r.text
    assert graph.calls("POST", "/me/messages")[0].json["toRecipients"][0]["emailAddress"]["address"] == "mb@x.test"


async def test_an_address_resolves_even_when_the_name_is_shared(graph, invoke):
    two_megans(graph)
    graph.on("POST", "/me/messages", {"id": "d1"})
    graph.on("GET", "/me/messages/d1", {"id": "d1"})
    r = await invoke("save_draft", to=["megan@x.test"], subject="s", body="x")
    assert not r.error, r.text


# -- personal contacts -------------------------------------------------------------------------

async def test_a_personal_contact_resolves_by_name(graph, invoke):
    graph.on("GET", "/me/people", {"value": []})
    graph.on("GET", "/users", {"value": []})
    graph.on("GET", "/me/contacts", {"value": [
        {"id": "c1", "displayName": "Mia Accountant", "emailAddresses": [{"address": "mia@example.org"}]}]})
    graph.on("POST", "/me/messages", {"id": "d1"})
    graph.on("GET", "/me/messages/d1", {"id": "d1"})
    r = await invoke("save_draft", to=["Mia"], subject="s", body="x")
    assert not r.error, r.text
    assert graph.calls("POST", "/me/messages")[0].json["toRecipients"][0]["emailAddress"]["address"] == "mia@example.org"


@pytest.mark.parametrize("tool, args", [
    ("save_draft", {"to": ["Team"], "subject": "s", "body": "x"}),
])
async def test_a_personal_contact_holding_a_group_address_is_refused(graph, invoke, tool, args):
    directory(graph, {})
    graph.on("GET", "/me/people", {"value": []})
    graph.on("GET", "/me/contacts", {"value": [
        {"id": "c1", "displayName": "Team", "emailAddresses": [{"address": GROUP}]}]})
    r = await invoke(tool, **args)
    assert r.error and GROUP in r.text and graph.writes() == []


async def test_a_name_is_not_resolved_when_contacts_could_not_all_be_checked(graph, invoke):
    graph.on("GET", "/me/contacts", {"value": [
        {"id": "c1", "displayName": "Mia Accountant", "emailAddresses": [{"address": "mia@example.org"}]}],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/contacts?$skip=100"})
    r = await invoke("save_draft", to=["Mia"], subject="s", body="x")
    assert r.error and "too many personal contacts" in r.text and "address" in r.text
    assert graph.writes() == []


async def test_a_shared_first_name_is_refused_when_finding_meeting_times(graph, invoke):
    two_megans(graph)
    r = await invoke("find_meeting_times", people=["Megan"], start="2026-10-23 09:00", end="2026-10-23 18:00")
    assert r.error and "megan@x.test" in r.text and "mb@x.test" in r.text
    assert graph.calls("POST", "/me/calendar/getSchedule") == [] and graph.calls("POST", "/me/findMeetingTimes") == []


async def test_a_shared_first_name_is_refused_when_opening_a_one_to_one_conversation(graph, invoke):
    two_megans(graph)
    me(graph)
    r = await invoke("get_conversation", person="Megan")
    assert r.error and "megan@x.test" in r.text and "mb@x.test" in r.text
    assert graph.calls("GET", "/me/chats") == []


async def test_contacts_in_other_folders_count_for_a_shared_name(graph, invoke):
    graph.on("GET", "/me/people", {"value": []})
    graph.on("GET", "/me/contacts", {"value": [
        {"id": "c1", "displayName": "Mia Berg", "emailAddresses": [{"address": "mia.b@example.org"}]}]})
    graph.on("GET", "/me/contactFolders", {"value": [{"id": "F1", "displayName": "Clients"}]})
    graph.on("GET", "/me/contactFolders/F1/contacts", {"value": [
        {"id": "c2", "displayName": "Mia Kim", "emailAddresses": [{"address": "mia.k@example.org"}]}]})
    graph.on("GET", "/me/contactFolders/F1/childFolders", {"value": []})
    r = await invoke("save_draft", to=["Mia"], subject="s", body="x")
    assert r.error and "mia.b@example.org" in r.text and "mia.k@example.org" in r.text
    assert graph.writes() == []


async def test_a_contact_holding_a_group_people_knows_is_refused_even_outside(graph, invoke):
    # An outside list cannot be told from the directory; People's own word that it is a group counts.
    graph.on("GET", "/me/contacts", {"value": [
        {"id": "c1", "displayName": "Partners List", "emailAddresses": [{"address": "list@partner.example"}]}]})
    graph.on("GET", "/me/people", {"value": [{"displayName": "Partners List", "scoredEmailAddresses": [{"address": "list@partner.example"}],
                                              "personType": {"class": "Group", "subclass": "PublicDistributionList"}}]})
    r = await invoke("save_draft", to=["Partners List"], subject="s", body="x")
    assert r.error and "list@partner.example" in r.text and graph.writes() == []


async def test_a_group_people_knows_under_another_name_still_marks_the_contact(graph, invoke):
    graph.on("GET", "/me/contacts", {"value": [
        {"id": "c1", "displayName": "Sales", "emailAddresses": [{"address": "sales@partner.example"}]}]})
    graph.on("GET", "/me/people", {"value": [{"displayName": "Regional Team", "scoredEmailAddresses": [{"address": "sales@partner.example"}],
                                              "personType": {"class": "Group", "subclass": "PublicDistributionList"}}]})
    r = await invoke("save_draft", to=["Sales"], subject="s", body="x")
    assert r.error and "sales@partner.example" in r.text and graph.writes() == []
