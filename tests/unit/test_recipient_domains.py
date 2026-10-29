"""An installation may limit whom the assistant reaches to its own domains. When it does, every action
that notifies people refuses a recipient outside them, before Graph is asked to do anything, and the
refusal names the domains and never the addresses."""

from dataclasses import replace

import httpx
import pytest

from m365_mcp.core.context import configure
from m365_mcp.core.errors import M365Error
from m365_mcp.core.guard import Guard
from m365_mcp.tools.mail import DRAFT_MARK

OWN = {"singleValueExtendedProperties": [{"id": DRAFT_MARK, "value": "1"}]}
ALLOWED = frozenset({"x.test", "partner.test"})


@pytest.fixture
def limited(rt):
    rt.settings = replace(rt.settings, allowed_recipient_domains=ALLOWED)
    configure(rt)
    return rt


def people(graph, *entries):
    def route(seen):
        q = seen.params.get("$search", "").strip('"').lower()
        hit = [(n, a) for n, a in entries if q and (q in n.lower() or q in a.lower())]
        return httpx.Response(200, json={"value": [{
            "displayName": n, "scoredEmailAddresses": [{"address": a}],
            "personType": {"class": "Person", "subclass": "OrganizationUser"}} for n, a in hit]})
    graph.on("GET", "/me/people", route)


# -- the guard ---------------------------------------------------------------------------------

def test_guard_refuses_an_address_outside_the_allowed_domains():
    with pytest.raises(M365Error) as e:
        Guard().check("p", kind="mail", user_confirmed=True, recipients=2,
                      addresses=["ann@x.test", "eve@Outside.test"], allowed_domains=ALLOWED)
    assert e.value.key == "outside_domains"
    assert "outside.test" in e.value.text and "eve" not in e.value.text


def test_guard_accepts_subdomains_and_any_case():
    r = Guard().check("p", kind="mail", user_confirmed=True, recipients=2,
                      addresses=["ann@X.TEST", "bo@eu.partner.test"], allowed_domains=ALLOWED)
    assert r["recipients"] == 2


def test_a_lookalike_domain_is_not_a_subdomain():
    with pytest.raises(M365Error) as e:
        Guard().check("p", kind="mail", user_confirmed=True, recipients=1,
                      addresses=["mallory@evilx.test"], allowed_domains=ALLOWED)
    assert "evilx.test" in e.value.text


def test_an_address_without_a_domain_is_refused():
    with pytest.raises(M365Error) as e:
        Guard().check("p", kind="mail", user_confirmed=True, recipients=1, addresses=["nobody"],
                      allowed_domains=ALLOWED)
    assert e.value.key == "outside_domains"


def test_without_allowed_domains_any_domain_passes():
    r = Guard().check("p", kind="mail", user_confirmed=True, recipients=1, addresses=["eve@outside.test"])
    assert r["recipients"] == 1


def test_a_refused_domain_takes_nothing_from_the_bucket():
    g = Guard()
    with pytest.raises(M365Error):
        g.check("p", kind="mail", user_confirmed=True, recipients=1, addresses=["eve@outside.test"],
                allowed_domains=ALLOWED)
    assert g._buckets == {}


# -- the tools ---------------------------------------------------------------------------------

async def test_send_draft_refuses_a_recipient_outside_the_allowed_domains(graph, invoke, limited):
    graph.on("GET", "/me/messages/d1", {"id": "d1", "isDraft": True, "subject": "S", **OWN,
                                        "toRecipients": [{"emailAddress": {"address": "ann@x.test"}}],
                                        "bccRecipients": [{"emailAddress": {"address": "eve@outside.test"}}]})
    people(graph, ("Ann", "ann@x.test"), ("Eve", "eve@outside.test"))
    r = await invoke("send_draft", id="d1", user_confirmed=True)
    assert r.error and "outside.test" in r.text and "eve@" not in r.text
    assert graph.writes() == []


async def test_send_draft_inside_the_allowed_domains_sends(graph, invoke, limited):
    graph.on("GET", "/me/messages/d1", {"id": "d1", "isDraft": True, "subject": "S", **OWN,
                                        "toRecipients": [{"emailAddress": {"address": "ann@x.test"}}]})
    people(graph, ("Ann", "ann@x.test"))
    graph.on("PATCH", "/me/messages/d1", {})  # the sent mark
    graph.on("POST", "/me/messages/d1/send", None, status=202)
    r = await invoke("send_draft", id="d1", user_confirmed=True)
    assert not r.error, r.text


async def test_inviting_an_attendee_outside_the_allowed_domains_is_refused(graph, invoke, limited):
    people(graph, ("Ann", "ann@x.test"), ("Eve", "eve@outside.test"))
    r = await invoke("save_event", subject="Sync", start="2026-10-23 15:00", duration_minutes=30,
                     attendees=["ann@x.test", "eve@outside.test"], user_confirmed=True)
    assert r.error and "outside.test" in r.text and "eve@" not in r.text
    assert graph.writes() == []


async def test_sharing_with_someone_outside_the_allowed_domains_is_refused(graph, invoke, limited):
    people(graph, ("Eve", "eve@outside.test"))
    r = await invoke("share_file", id="f1", action="invite", people=["Eve"], user_confirmed=True)
    assert r.error and "outside.test" in r.text and "eve@" not in r.text
    assert graph.writes() == []


async def test_a_teams_message_to_someone_outside_the_allowed_domains_is_refused(graph, invoke, limited):
    graph.on("GET", "/me", {"id": "me-id", "displayName": "Me", "mail": "me@x.test", "userPrincipalName": "me@x.test"})
    people(graph, ("Eve", "eve@outside.test"))
    r = await invoke("send_message", to=["eve@outside.test"], body="hi", user_confirmed=True)
    assert r.error and "outside.test" in r.text and "eve@" not in r.text
    assert graph.writes() == []


async def test_adding_a_chat_member_outside_the_allowed_domains_is_refused(graph, invoke, limited):
    graph.on("GET", "/me", {"id": "me-id", "displayName": "Me", "mail": "me@x.test", "userPrincipalName": "me@x.test"})
    graph.on("GET", "/me/people", {"value": []})
    r = await invoke("save_chat", topic="T", add_members=["a@x.test", "eve@outside.test"], user_confirmed=True)
    assert r.error and "outside.test" in r.text and graph.writes() == []


def test_allowed_domains_are_read_as_a_comma_list(monkeypatch):
    from m365_mcp.config import Settings
    monkeypatch.setenv("M365_AREAS", "account")
    monkeypatch.setenv("M365_ALLOWED_RECIPIENT_DOMAINS", " X.test, @partner.test ,")
    assert Settings.from_env().allowed_recipient_domains == ALLOWED


CHAT = "19:abc@thread.v2"


def _me(graph):
    graph.on("GET", "/me", {"id": "me-id", "displayName": "Me", "mail": "me@x.test", "userPrincipalName": "me@x.test"})


async def test_a_message_to_an_existing_chat_with_an_outside_member_is_refused(graph, invoke, limited):
    _me(graph)
    graph.on("GET", f"/chats/{CHAT}/members", {"value": [{"email": "me@x.test", "userId": "me-id"},
                                                        {"email": "eve@outside.test", "userId": "u-eve"}]})
    r = await invoke("send_message", to=[CHAT], body="hi", user_confirmed=True)
    assert r.error and "outside.test" in r.text and "eve@" not in r.text
    assert graph.writes() == []


async def test_a_message_to_an_existing_chat_inside_the_domains_is_sent(graph, invoke, limited):
    _me(graph)
    graph.on("GET", f"/chats/{CHAT}/members", {"value": [{"email": "me@x.test"}, {"email": "ann@x.test"}]})
    graph.on("POST", f"/chats/{CHAT}/messages", {"id": "m1"}, status=201)
    r = await invoke("send_message", to=[CHAT], body="hi", user_confirmed=True)
    assert not r.error, r.text


async def test_without_allowed_domains_an_existing_chat_is_not_listed(graph, invoke):
    _me(graph)
    graph.on("POST", f"/chats/{CHAT}/messages", {"id": "m1"}, status=201)
    r = await invoke("send_message", to=[CHAT], body="hi", user_confirmed=True)
    assert not r.error and not graph.calls("GET", f"/chats/{CHAT}/members")


async def test_a_chat_member_without_an_address_is_refused_even_when_the_sender_has_no_mail(graph, invoke, limited):
    graph.on("GET", "/me", {"id": "me-id", "displayName": "Me", "mail": None, "userPrincipalName": "me@x.test"})
    graph.on("GET", f"/chats/{CHAT}/members", {"value": [{"userId": "me-id", "email": None},
                                                        {"userId": "u-guest", "email": None}]})
    r = await invoke("send_message", to=[CHAT], body="hi", user_confirmed=True)
    assert r.error and "no domain" in r.text and graph.writes() == []


async def test_every_page_of_chat_members_is_checked(graph, invoke, limited):
    _me(graph)
    nxt = f"https://graph.microsoft.com/v1.0/chats/{CHAT}/members?$skiptoken=2"
    graph.on("GET", f"/chats/{CHAT}/members", lambda s: httpx.Response(200, json=(
        {"value": [{"userId": "u-eve", "email": "eve@outside.test"}]} if "$skiptoken" in s.params else
        {"value": [{"userId": "u-ann", "email": "ann@x.test"}], "@odata.nextLink": nxt})))
    r = await invoke("send_message", to=[CHAT], body="hi", user_confirmed=True)
    assert r.error and "outside.test" in r.text and graph.writes() == []


async def test_a_comment_to_an_organiser_outside_the_allowed_domains_is_refused(graph, invoke, limited):
    graph.on("GET", "/me/events/e1", {"id": "e1", "subject": "S", "isOrganizer": False, "isCancelled": False,
                                      "organizer": {"emailAddress": {"name": "Eve", "address": "eve@outside.test"}}})
    graph.on("POST", "/me/events/e1/accept", None, status=202)
    r = await invoke("respond_to_event", id="e1", response="accept", comment="ok", user_confirmed=True)
    assert r.error and "outside.test" in r.text and "eve@" not in r.text
    assert graph.writes() == []


# -- the audiences that are not named recipients ------------------------------------------------

@pytest.mark.parametrize("audience", ["all", "contacts"])
async def test_an_automatic_reply_to_people_outside_is_refused_when_domains_are_limited(graph, invoke, limited, audience):
    graph.on("PATCH", "/me/mailboxSettings", None, status=204)
    r = await invoke("save_automatic_reply", status="on", internal_message="away", external_message="away",
                     audience=audience, user_confirmed=True)
    assert r.error and "audience=none" in r.text and "x.test" in r.text
    assert graph.writes() == []


async def test_an_automatic_reply_for_the_organisation_only_is_set_when_domains_are_limited(graph, invoke, limited):
    graph.on("PATCH", "/me/mailboxSettings", None, status=204)
    r = await invoke("save_automatic_reply", status="on", internal_message="away", audience="none", user_confirmed=True)
    assert not r.error, r.text
    off = await invoke("save_automatic_reply", status="off")
    assert not off.error, off.text


async def test_a_cancellation_comment_to_an_attendee_outside_the_allowed_domains_is_refused(graph, invoke, limited):
    graph.on("GET", "/me/events/e1", {"id": "e1", "subject": "S", "isOrganizer": True, "isCancelled": False,
                                      "type": "singleInstance", "attendees": [
                                          {"type": "required", "emailAddress": {"address": "ann@x.test"}},
                                          {"type": "required", "emailAddress": {"address": "eve@outside.test"}}]})
    graph.on("POST", "/me/events/e1/cancel", None, status=202)
    r = await invoke("cancel_event", id="e1", comment="off", user_confirmed=True)
    assert r.error and "outside.test" in r.text and "eve@" not in r.text
    assert graph.writes() == []
    # Without a comment the cancellation is Exchange's own notice, and it must reach everyone invited.
    r = await invoke("cancel_event", id="e1", user_confirmed=True)
    assert not r.error, r.text


def _channel(graph):
    _me(graph)
    graph.on("GET", "/me/joinedTeams", {"value": [{"id": "t1", "displayName": "Ops"}]})
    graph.on("GET", "/teams/t1/channels", {"value": [{"id": "c1", "displayName": "General"}]})
    graph.on("POST", "/teams/t1/channels/c1/messages", {"id": "p1"}, status=201)


async def test_a_channel_post_with_a_member_outside_the_allowed_domains_is_refused(graph, invoke, limited):
    _channel(graph)
    graph.on("GET", "/teams/t1/channels/c1/members", {"value": [{"userId": "me-id", "email": "me@x.test"},
                                                                {"userId": "u-eve", "email": "eve@outside.test"}]})
    r = await invoke("send_message", to=["Ops/General"], body="hi", user_confirmed=True)
    assert r.error and "outside.test" in r.text and "eve@" not in r.text
    assert graph.writes() == []


async def test_a_channel_post_inside_the_allowed_domains_is_sent(graph, invoke, limited):
    _channel(graph)
    graph.on("GET", "/teams/t1/channels/c1/members", {"value": [{"userId": "me-id", "email": "me@x.test"},
                                                                {"userId": "u-ann", "email": "ann@x.test"}]})
    r = await invoke("send_message", to=["Ops/General"], body="hi", user_confirmed=True)
    assert not r.error, r.text


async def test_a_channel_post_whose_members_cannot_be_read_is_refused_when_domains_are_limited(graph, invoke, limited):
    from helpers import ALL_SCOPES, make_token
    from m365_mcp.identity import current_token
    _channel(graph)
    assert "ChannelMember.Read.All" in ALL_SCOPES
    t = current_token.set(make_token(ALL_SCOPES.replace(" ChannelMember.Read.All", "")))
    try:
        r = await invoke("send_message", to=["Ops/General"], body="hi", user_confirmed=True)
    finally:
        current_token.reset(t)
    assert r.error and "ChannelMember.Read.All" in r.text and graph.writes() == []


# -- changing a message: the new or restored text reaches the conversation as it is now ------------

def _own_message(graph, path):
    graph.on("GET", path, {"id": "m1", "messageType": "message", "from": {"user": {"id": "me-id"}},
                           "body": {"contentType": "html", "content": "old"}})


@pytest.mark.parametrize("action", ["edit", "restore"])
async def test_changing_a_chat_message_with_an_outside_member_is_refused(graph, invoke, limited, action):
    _me(graph)
    _own_message(graph, f"/chats/{CHAT}/messages/m1")
    graph.on("GET", f"/chats/{CHAT}/members", {"value": [{"email": "me@x.test", "userId": "me-id"},
                                                        {"email": "eve@outside.test", "userId": "u-eve"}]})
    graph.on("PATCH", f"/chats/{CHAT}/messages/m1", None, status=204)
    graph.on("POST", f"/users/me-id/chats/{CHAT}/messages/m1/undoSoftDelete", None, status=204)
    r = await invoke("update_message", id="m1", chat=CHAT, action=action, body="new text", user_confirmed=True)
    assert r.error and "outside.test" in r.text and "eve@" not in r.text
    assert graph.writes() == []


async def test_changing_a_channel_post_with_an_outside_member_is_refused(graph, invoke, limited):
    _channel(graph)
    _own_message(graph, "/teams/t1/channels/c1/messages/m1")
    graph.on("GET", "/teams/t1/channels/c1/members", {"value": [{"userId": "me-id", "email": "me@x.test"},
                                                                {"userId": "u-eve", "email": "eve@outside.test"}]})
    graph.on("PATCH", "/teams/t1/channels/c1/messages/m1", None, status=204)
    r = await invoke("update_message", id="m1", channel="Ops/General", action="edit", body="new", user_confirmed=True)
    assert r.error and "outside.test" in r.text and "eve@" not in r.text
    assert graph.writes() == []


async def test_changing_a_message_inside_the_domains_is_done_and_a_delete_reads_no_members(graph, invoke, limited):
    _me(graph)
    _own_message(graph, f"/chats/{CHAT}/messages/m1")
    graph.on("GET", f"/chats/{CHAT}/members", {"value": [{"email": "me@x.test", "userId": "me-id"},
                                                        {"email": "ann@x.test", "userId": "u-ann"}]})
    graph.on("PATCH", f"/chats/{CHAT}/messages/m1", None, status=204)
    graph.on("POST", f"/users/me-id/chats/{CHAT}/messages/m1/softDelete", None, status=204)
    r = await invoke("update_message", id="m1", chat=CHAT, action="edit", body="new", user_confirmed=True)
    assert not r.error, r.text
    assert len(graph.calls("GET", f"/chats/{CHAT}/members")) == 1
    # A delete delivers no text: it is not refused for who is in the conversation.
    r = await invoke("update_message", id="m1", chat=CHAT, action="delete", user_confirmed=True)
    assert not r.error, r.text
    assert len(graph.calls("GET", f"/chats/{CHAT}/members")) == 1


async def test_an_automatic_reply_without_an_audience_is_for_the_organisation_when_domains_are_limited(graph, invoke, limited):
    graph.on("PATCH", "/me/mailboxSettings", None, status=204)
    r = await invoke("save_automatic_reply", status="on", internal_message="away", user_confirmed=True)
    assert not r.error, r.text
    assert graph.writes()[0].json["automaticRepliesSetting"]["externalAudience"] == "none"
    assert r.data["audience"] == "none" and "note" not in r.data
