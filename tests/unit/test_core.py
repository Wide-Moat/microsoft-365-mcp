import logging
from zoneinfo import ZoneInfo

import pytest

from helpers import make_token
from m365_mcp.core.context import scope_granted
from m365_mcp.core.errors import M365Error
from m365_mcp.core.guard import Guard
from m365_mcp.core.time import from_graph, graph_dt, parse, window, zone
from m365_mcp.identity import person_from_token
from m365_mcp.observability import ContentFilter

LISBON = ZoneInfo("Europe/Lisbon")


# -- guard ------------------------------------------------------------------------------------

def test_guard_requires_confirmation():
    with pytest.raises(M365Error) as e:
        Guard().check("p", kind="mail", user_confirmed=False)
    assert e.value.key == "confirm"


def test_guard_allows_own_response_without_confirmation():
    assert Guard().check("p", kind="respond", user_confirmed=None, needs_confirmation=False)["kind"] == "respond"


def test_guard_caps_recipients_and_refuses_lists():
    with pytest.raises(M365Error) as e:
        Guard().check("p", kind="mail", user_confirmed=True, recipients=21)
    assert e.value.key == "too_many_recipients"
    with pytest.raises(M365Error) as e:
        Guard().check("p", kind="mail", user_confirmed=True, recipients=1, group_addresses=["all@x.test"])
    assert e.value.key == "distribution_list"


def test_bucket_bursts_ten_then_refills_one_per_six_seconds():
    now = {"t": 0.0}
    g = Guard(clock=lambda: now["t"])
    for _ in range(10):
        g.check("p", kind="mail", user_confirmed=True)
    with pytest.raises(M365Error) as e:
        g.check("p", kind="mail", user_confirmed=True)
    assert e.value.key == "rate" and e.value.params["seconds"] == 6
    now["t"] = 6.0
    g.check("p", kind="mail", user_confirmed=True)
    g.check("other-person", kind="mail", user_confirmed=True)  # buckets are per person


# -- time -------------------------------------------------------------------------------------

def test_local_time_uses_the_mailbox_zone():
    dt = parse("2026-10-23 15:00", zone("GMT Standard Time"))
    assert graph_dt(dt) == {"dateTime": "2026-10-23T14:00:00", "timeZone": "UTC"}


def test_explicit_offset_is_respected():
    assert graph_dt(parse("2026-10-23 15:00 +03:00", LISBON))["dateTime"] == "2026-10-23T12:00:00"


@pytest.mark.parametrize("bad", ["tomorrow", "23/10/2026", "2026-13-01", "2026-02-30 10:00"])
def test_unparseable_times_are_refused(bad):
    with pytest.raises(M365Error):
        parse(bad, LISBON)


def test_window_limits():
    with pytest.raises(M365Error):
        window("2026-10-10", "2026-10-01", LISBON, max_days=31)
    with pytest.raises(M365Error):
        window("2026-01-01", "2026-04-01", LISBON, max_days=62)
    assert window("2026-10-01", "2026-10-31", LISBON, max_days=62)


def test_graph_times_come_back_local_with_offset():
    assert from_graph({"dateTime": "2026-10-23T14:00:00.0000000", "timeZone": "UTC"}, LISBON) == "2026-10-23T15:00+01:00"
    assert from_graph("2026-07-01T10:00:00Z", LISBON) == "2026-07-01T11:00+01:00"


# -- identity and scopes ---------------------------------------------------------------------

def test_person_key_differs_by_person_and_hides_the_token():
    a = person_from_token(make_token(oid="a"))
    b = person_from_token(make_token(oid="b"))
    assert a.key != b.key
    assert a.token not in repr(a)


@pytest.mark.parametrize("required, granted, ok", [
    ("Mail.Read", "Mail.ReadWrite", True),
    ("Files.Read", "Files.ReadWrite.All", True),
    ("Files.Read.All", "Files.ReadWrite", False),
    ("Mail.Send", "Mail.ReadWrite", False),
    ("Calendars.ReadBasic", "Calendars.Read", True),
    ("Sites.ReadWrite.All", "Sites.Read.All", False),
])
def test_scope_satisfaction(required, granted, ok):
    assert scope_granted(required, frozenset(granted.split())) is ok


# -- errors and logs ---------------------------------------------------------------------------

def test_error_str_is_the_key_alone():
    e = M365Error("ambiguous", kind="person", query="Anna", candidates="Anna Berg <ab@x.test>")
    assert str(e) == "ambiguous"
    assert "ab@x.test" in e.text


def test_log_filter_drops_content_and_counts_it():
    f = ContentFilter()
    ok = logging.LogRecord("x", logging.INFO, "", 0, "tool list_messages ok", None, None)
    leak = logging.LogRecord("x", logging.INFO, "", 0, "searching for anna@contoso.test", None, None)
    jwt = logging.LogRecord("x", logging.INFO, "", 0, "token eyJhbGciOiJub25lIn0.eyJhdWQiOiJodHRwcyJ9.x", None, None)
    assert f.filter(ok) and not f.filter(leak) and not f.filter(jwt)


async def test_a_full_mail_flow_logs_no_content(graph, invoke, caplog):
    import httpx
    caplog.set_level(logging.DEBUG)
    graph.on("GET", "/me/people", lambda s: httpx.Response(200, json={"value": [
        {"displayName": "Secret Person", "scoredEmailAddresses": [{"address": "secret.person@contoso.test"}]}]}))
    graph.on("GET", "/me/messages", {"value": [{"id": "m1", "subject": "Confidential merger",
                                                "from": {"emailAddress": {"address": "secret.person@contoso.test"}}}]})
    graph.on("GET", "/me/messages/m1", {"error": {"code": "ErrorAccessDenied", "message": "secret.person@contoso.test"}}, status=403)
    await invoke("list_messages", query="merger plans")
    await invoke("get_message", id="m1")
    await invoke("save_draft", to=["Secret Person", "Other Person"], subject="Confidential merger")
    text = caplog.text
    for needle in ("secret.person", "Secret Person", "Confidential", "merger", "Other Person", "eyJ"):
        assert needle not in text, needle


async def test_an_unknown_windows_zone_is_learned_from_graph(graph, invoke):
    from m365_mcp.core import time as t
    t.LEARNED_ZONES.pop("Singapore Standard Time", None)
    graph.on("GET", "/me/mailboxSettings/timeZone", {"value": "Singapore Standard Time"})
    graph.on("GET", r"/me/outlook/supportedTimeZones\(TimeZoneStandard=microsoft.graph.timeZoneStandard'Windows'\)",
             {"value": [{"alias": "Singapore Standard Time", "displayName": "(UTC+08:00) Kuala Lumpur, Singapore"}]})
    graph.on("GET", r"/me/outlook/supportedTimeZones\(TimeZoneStandard=microsoft.graph.timeZoneStandard'Iana'\)",
             {"value": [{"alias": "Asia/Singapore", "displayName": "(UTC+08:00) Kuala Lumpur, Singapore"}]})  # gitleaks:allow
    graph.on("GET", "/me", {"displayName": "Ann"})
    graph.on("GET", "/me/mailboxSettings", {})
    r = await invoke("get_connection")
    assert r.data["time_zone"] == "Singapore Standard Time"
    assert t.zone("Singapore Standard Time").key == "Asia/Singapore"


async def test_requests_do_not_ask_graph_for_local_times(graph, invoke):
    graph.on("GET", "/me", {"displayName": "Ann"})
    graph.on("GET", "/me/mailboxSettings", {})
    await invoke("get_connection")
    assert all("outlook.timezone" not in s.headers.get("prefer", "") for s in graph.seen)


def test_a_server_failure_never_promises_nothing_happened():
    from m365_mcp.graph.errors import GraphError, to_person_error
    text = to_person_error(GraphError(502, "x", "rid")).text
    assert "may or may not have happened" in text and "nothing was changed" not in text
    assert "nothing was changed" in to_person_error(GraphError(400, "x", "rid")).text


async def test_a_failed_time_zone_read_is_not_guessed_as_utc(graph, invoke):
    graph.on("GET", "/me/mailboxSettings/timeZone", {"error": {"code": "x"}}, status=503)
    r = await invoke("list_events", start="2026-10-23 09:00", end="2026-10-24 09:00")
    assert r.error and not graph.calls("GET", "/me/calendarView")


async def test_a_mailbox_with_no_zone_set_reads_as_utc(graph, invoke):
    # Measured on a new mailbox: Graph answers 204 with no body until a zone is chosen in Outlook.
    graph.on("GET", "/me/mailboxSettings/timeZone", None, status=204)
    graph.on("GET", "/me", {"displayName": "Ann"})
    graph.on("GET", "/me/mailboxSettings", {})
    r = await invoke("get_connection")
    assert not r.error, r.text
    assert r.data["time_zone"] == "UTC"


def test_cursor_key_must_be_real():
    from m365_mcp.core.cursor import Cursors
    with pytest.raises(ValueError):
        Cursors(b"")


def test_a_receipt_always_shows_what_is_left_and_nothing_else():
    # Shown at 0-2 only, it came too late for an agent to pace a run of sends.
    from m365_mcp.core.guard import shown
    assert shown({"kind": "mail", "recipients": 2, "sends_left_now": 9}) == {"sends_left": 9}
    assert shown({"kind": "mail", "recipients": 2, "sends_left_now": 1}) == {"sends_left": 1}


def test_log_filter_drops_a_traceback_that_carries_content():
    import io
    import sys
    f = ContentFilter()
    try:
        raise ValueError("no mailbox for anna@contoso.test")
    except ValueError:
        exc = sys.exc_info()
    rec = logging.LogRecord("x", logging.ERROR, "", 0, "tool get_message crashed: ValueError", None, exc)
    assert f.filter(rec), "the line itself carries no content and is kept"
    out = io.StringIO()
    h = logging.StreamHandler(out)
    h.addFilter(f)
    h.handle(rec)
    assert "crashed" in out.getvalue() and "anna@contoso.test" not in out.getvalue()
    clean = logging.LogRecord("x", logging.ERROR, "", 0, "crashed", None, None)
    clean.stack_info = 'File "x.py", line 1, in f\n    send("anna@contoso.test")'
    assert f.filter(clean) and clean.stack_info is None


async def test_a_crash_carrying_content_logs_none_of_it_through_the_filter(graph, invoke, monkeypatch):
    import io

    from m365_mcp.observability import install_filter
    from m365_mcp.tools import mail
    out = io.StringIO()
    h = logging.StreamHandler(out)
    h.setFormatter(logging.Formatter("%(name)s %(message)s"))
    root = logging.getLogger()
    root.addHandler(h)
    old_level = root.level
    root.setLevel(logging.DEBUG)
    install_filter()

    def boom(*_a, **_k):
        raise RuntimeError("cannot read secret.person@contoso.test")
    monkeypatch.setattr(mail, "envelope", boom)
    graph.on("GET", "/me/messages/m1", {"id": "m1"})
    try:
        r = await invoke("get_message", id="m1")
        logging.getLogger("uvicorn.error").exception("ASGI app failed", exc_info=RuntimeError("x secret.person@contoso.test"))
    finally:
        root.removeHandler(h)
        root.setLevel(old_level)
    assert r.error
    assert "secret.person" not in out.getvalue()
