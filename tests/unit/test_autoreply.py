import pytest

BERLIN = "W. Europe Standard Time"


@pytest.fixture(autouse=True)
def berlin(graph):
    graph.on("GET", "/me/mailboxSettings/timeZone", {"value": BERLIN})
    graph.on("PATCH", "/me/mailboxSettings", {})


def sent(graph):
    [w] = graph.writes()
    assert w.method == "PATCH" and w.path == "/me/mailboxSettings"
    return w.json["automaticRepliesSetting"]


async def test_turning_it_off_needs_no_confirmation_and_no_bucket(graph, invoke, guard):
    r = await invoke("save_automatic_reply", status="off")
    assert not r.error, r.text
    assert sent(graph) == {"status": "disabled"}
    assert r.data == {"automatic_reply": "off"} and guard._buckets == {}


async def test_turning_it_on_without_confirmation_writes_nothing(graph, invoke):
    r = await invoke("save_automatic_reply", status="on", internal_message="Away")
    assert r.error and "confirm" in r.text and graph.writes() == []


async def test_scheduling_without_confirmation_writes_nothing(graph, invoke):
    r = await invoke("save_automatic_reply", status="scheduled", internal_message="Away",
                     start="2099-01-01 08:00", end="2099-01-08 08:00")
    assert r.error and "confirm" in r.text and graph.writes() == []


async def test_turning_it_on_sends_html_messages_and_the_audience(graph, invoke):
    r = await invoke("save_automatic_reply", status="on", internal_message="Out until Monday.\nCall Bo.",
                     external_message="Away, back Monday.", audience="all", user_confirmed=True)
    assert not r.error, r.text
    s = sent(graph)
    assert s["status"] == "alwaysEnabled" and s["externalAudience"] == "all"
    assert "Out until Monday." in s["internalReplyMessage"] and "<" in s["internalReplyMessage"]
    assert "Away, back Monday." in s["externalReplyMessage"]
    assert "scheduledStartDateTime" not in s
    assert r.data["automatic_reply"] == "on" and r.data["audience"] == "all" and not ({"sends_left_now", "kind", "recipients"} & set(r.data))


async def test_the_default_audience_is_contacts_and_a_missing_external_text_is_reported(graph, invoke):
    r = await invoke("save_automatic_reply", status="on", internal_message="Away", user_confirmed=True)
    assert sent(graph)["externalAudience"] == "contactsOnly" and "externalReplyMessage" not in sent(graph)
    assert "No external_message" in r.data["note"]


async def test_audience_none_is_sent_as_none(graph, invoke):
    r = await invoke("save_automatic_reply", status="on", internal_message="Away", audience="none", user_confirmed=True)
    assert sent(graph)["externalAudience"] == "none" and "note" not in r.data


async def test_message_text_cannot_carry_active_html(graph, invoke):
    await invoke("save_automatic_reply", status="on", internal_message="<script>steal()</script>Hi",
                 user_confirmed=True)
    assert "<script" not in sent(graph)["internalReplyMessage"]


async def test_a_schedule_is_sent_in_utc_from_the_mailbox_zone(graph, invoke):
    r = await invoke("save_automatic_reply", status="scheduled", internal_message="Holiday",
                     start="2099-07-01 08:00", end="2099-07-15 18:00", user_confirmed=True)
    assert not r.error, r.text
    s = sent(graph)
    assert s["status"] == "scheduled"
    assert s["scheduledStartDateTime"] == {"dateTime": "2099-07-01T06:00:00", "timeZone": "UTC"}
    assert s["scheduledEndDateTime"] == {"dateTime": "2099-07-15T16:00:00", "timeZone": "UTC"}


@pytest.mark.parametrize("dates", [{}, {"start": "2099-07-01 08:00"}, {"end": "2099-07-15 18:00"}])
async def test_scheduled_without_both_dates_is_refused(graph, invoke, dates):
    r = await invoke("save_automatic_reply", status="scheduled", internal_message="Away", user_confirmed=True, **dates)
    assert r.error and "both start and end" in r.text and graph.writes() == []


async def test_an_end_in_the_past_is_refused(graph, invoke):
    r = await invoke("save_automatic_reply", status="scheduled", internal_message="Away",
                     start="2020-01-01 08:00", end="2020-01-08 08:00", user_confirmed=True)
    assert r.error and "in the past" in r.text and graph.writes() == []


async def test_an_end_before_the_start_is_refused(graph, invoke):
    r = await invoke("save_automatic_reply", status="scheduled", internal_message="Away",
                     start="2099-01-08 08:00", end="2099-01-01 08:00", user_confirmed=True)
    assert r.error and "end must be after start" in r.text and graph.writes() == []


async def test_dates_with_a_status_that_ignores_them_are_refused(graph, invoke):
    r = await invoke("save_automatic_reply", status="on", internal_message="Away", start="2099-01-01 08:00",
                     end="2099-01-08 08:00", user_confirmed=True)
    assert r.error and "scheduled" in r.text and graph.writes() == []


async def test_turning_it_on_with_no_text_anywhere_is_refused(graph, invoke):
    graph.on("GET", "/me/mailboxSettings/automaticRepliesSetting",
             {"status": "disabled", "internalReplyMessage": "", "externalReplyMessage": ""})
    r = await invoke("save_automatic_reply", status="on", user_confirmed=True)
    assert r.error and "no reply text" in r.text and graph.writes() == []


async def test_turning_it_on_reuses_the_text_already_set(graph, invoke):
    graph.on("GET", "/me/mailboxSettings/automaticRepliesSetting",
             {"status": "disabled", "internalReplyMessage": "<p>Old text</p>", "externalReplyMessage": ""})
    r = await invoke("save_automatic_reply", status="on", user_confirmed=True)
    assert not r.error, r.text
    s = sent(graph)
    assert s["status"] == "alwaysEnabled" and "internalReplyMessage" not in s


async def test_html_in_the_reply_text_passes_the_outgoing_pipeline(graph, invoke):
    # Everyone who writes receives this text; HTML in it is cleaned like a letter, not escaped into tags.
    r = await invoke("save_automatic_reply", status="on", audience="none", user_confirmed=True,
                     internal_message='<p onclick="x()">Away <script>alert(1)</script><b>until Monday</b></p>')
    assert not r.error, r.text
    msg = sent(graph)["internalReplyMessage"]
    assert "<b>until Monday</b>" in msg
    assert "script" not in msg and "onclick" not in msg and "alert" not in msg and "&lt;" not in msg
    assert any("script" in s for s in r.data["html"]["sanitized"])


@pytest.mark.parametrize("text, kept", [("Write to <alex@example.com> instead.", "alex@example.com"),
                                        ("Write to <a@example.com> or <strong@example.com>.", "strong@example.com"),
                                        ("Ask <name> at the desk.", "&lt;name&gt;")])
async def test_angle_brackets_in_plain_text_are_not_taken_for_html(graph, invoke, text, kept):
    await invoke("save_automatic_reply", status="on", internal_message=text, user_confirmed=True)
    assert kept in sent(graph)["internalReplyMessage"]
