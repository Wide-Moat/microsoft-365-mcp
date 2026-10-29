from dataclasses import replace

from m365_mcp.registry import AREAS
from m365_mcp.server import M365Server

BIZ = "salon@x.test"
TOOLS = {"list_bookings", "save_booking", "cancel_booking"}


def business(graph):
    graph.on("GET", "/solutions/bookingBusinesses", {"value": [{"id": BIZ, "displayName": "Salon", "email": "salon@x.test"}]})
    graph.on("GET", f"/solutions/bookingBusinesses/{BIZ}/services", {"value": [
        {"id": "svc1", "displayName": "Haircut", "defaultDuration": "PT1H30M"}]})
    graph.on("GET", f"/solutions/bookingBusinesses/{BIZ}/staffMembers", {"value": [{"id": "st1", "displayName": "Ida"}]})


async def test_bookings_are_off_by_default(settings, token):
    assert AREAS["bookings"] is False
    default = M365Server(replace(settings, areas=frozenset(a for a, on in AREAS.items() if on)))
    assert not TOOLS & {t.name for t in await default.list_tools()}
    assert TOOLS <= {t.name for t in await M365Server(settings).list_tools()}


async def test_list_businesses_then_appointments_in_a_window(graph, invoke):
    business(graph)
    r = await invoke("list_bookings")
    assert r.data["items"] == [{"business": "Salon", "email": "salon@x.test"}]
    graph.on("GET", f"/solutions/bookingBusinesses/{BIZ}/calendarView", {"value": [{
        "id": "a1", "customerName": "Kim", "customerEmailAddress": "kim@x.test", "serviceName": "Haircut",
        "startDateTime": {"dateTime": "2026-10-23T14:00:00", "timeZone": "UTC"},
        "endDateTime": {"dateTime": "2026-10-23T15:30:00", "timeZone": "UTC"}, "staffMemberIds": ["st1"]}]})
    r = await invoke("list_bookings", business="salon", start="2026-10-23", end="2026-10-24", include_services=True, include_staff=True)
    assert not r.error, r.text
    a = r.data["items"][0]
    assert a["customer"] == "Kim" and a["start"] == "2026-10-23T14:00+00:00" and a["staff"] == 1
    assert r.data["services"] == [{"name": "Haircut", "duration": "PT1H30M"}] and r.data["staff"] == ["Ida"]
    q = graph.calls("GET", f"/solutions/bookingBusinesses/{BIZ}/calendarView")[0].params
    assert q == {"start": "2026-10-23T00:00:00Z", "end": "2026-10-24T00:00:00Z"}
    r = await invoke("list_bookings", business="salon", start="2026-10-23")
    assert r.error and "both start and end" in r.text


async def test_booking_without_confirmation_writes_nothing(graph, invoke):
    business(graph)
    r = await invoke("save_booking", business="Salon", service="Haircut", start="2026-10-23 15:00",
                     customer_name="Kim", customer_email="kim@x.test")
    assert r.error and "user_confirmed" in r.text and graph.writes() == []


async def test_booking_creates_the_appointment_with_the_service_length(graph, invoke):
    business(graph)
    graph.on("POST", f"/solutions/bookingBusinesses/{BIZ}/appointments", {"id": "a9"}, status=201)
    r = await invoke("save_booking", business="Salon", service="hair", start="2026-10-23 15:00", customer_name="Kim",
                     customer_email="kim@x.test", staff=["Ida"], notes="first visit", user_confirmed=True)
    assert not r.error and r.data["id"] == "a9" and r.data["kind"] == "booking", r.text
    assert graph.writes()[0].json == {
        "customerName": "Kim", "customerEmailAddress": "kim@x.test", "customerNotes": "first visit",
        "startDateTime": {"dateTime": "2026-10-23T15:00:00", "timeZone": "UTC"}, "serviceId": "svc1", "serviceName": "Haircut",
        "endDateTime": {"dateTime": "2026-10-23T16:30:00", "timeZone": "UTC"}, "staffMemberIds": ["st1"]}


async def test_a_new_booking_needs_service_and_customer(graph, invoke):
    business(graph)
    r = await invoke("save_booking", business="Salon", start="2026-10-23 15:00", user_confirmed=True)
    assert r.error and "needs service" in r.text and graph.writes() == []


async def test_changing_a_booking_is_guarded_too(graph, invoke):
    business(graph)
    existing(graph)
    r = await invoke("save_booking", business="Salon", id="a1", start="2026-10-24 10:00")
    assert r.error and graph.writes() == []
    r = await invoke("save_booking", business="Salon", id="a1", start="2026-10-24 10:00", user_confirmed=True)
    assert not r.error, r.text
    assert graph.writes()[0].method == "PATCH"


async def test_cancel_needs_confirmation_and_a_message(graph, invoke):
    business(graph)
    r = await invoke("cancel_booking", id="a1", business="Salon", message="Sorry")
    assert r.error and "user_confirmed" in r.text and graph.writes() == []
    r = await invoke("cancel_booking", id="a1", business="Salon", message="  ", user_confirmed=True)
    assert r.error and graph.writes() == []
    graph.on("POST", f"/solutions/bookingBusinesses/{BIZ}/appointments/a1/cancel", None, status=204)
    r = await invoke("cancel_booking", id="a1", business="Salon", message="Staff are ill", user_confirmed=True)
    assert not r.error and r.data["cancelled"] is True, r.text
    assert graph.writes()[0].json == {"cancellationMessage": "Staff are ill"}


APPT = f"/solutions/bookingBusinesses/{BIZ}/appointments/a1"


def existing(graph):
    graph.on("GET", APPT, {"id": "a1", "serviceName": "Haircut",
                           "startDateTime": {"dateTime": "2026-10-23T14:00:00.0000000", "timeZone": "UTC"},
                           "endDateTime": {"dateTime": "2026-10-23T16:00:00.0000000", "timeZone": "UTC"}})
    graph.on("PATCH", APPT, None, status=204)


async def test_moving_a_booking_keeps_its_length(graph, invoke):
    business(graph)
    existing(graph)
    r = await invoke("save_booking", business="Salon", id="a1", start="2026-10-24 10:00", user_confirmed=True)
    assert not r.error, r.text
    assert graph.writes()[0].json == {"startDateTime": {"dateTime": "2026-10-24T10:00:00", "timeZone": "UTC"},
                                      "endDateTime": {"dateTime": "2026-10-24T12:00:00", "timeZone": "UTC"}}


async def test_changing_notes_needs_no_start(graph, invoke):
    business(graph)
    existing(graph)
    r = await invoke("save_booking", business="Salon", id="a1", notes="allergic", user_confirmed=True)
    assert not r.error, r.text
    assert graph.writes()[0].json == {"customerNotes": "allergic"}


async def test_a_new_booking_needs_a_start(graph, invoke):
    business(graph)
    r = await invoke("save_booking", business="Salon", service="Haircut", customer_name="Kim",
                     customer_email="kim@x.test", user_confirmed=True)
    assert r.error and "start" in r.text and graph.writes() == []


async def test_list_bookings_accepts_either_read_permission(graph, token, invoke):
    from m365_mcp.identity import current_token
    from helpers import make_token
    business(graph)
    for scope in ("Bookings.Read.All", "BookingsAppointment.ReadWrite.All"):
        tok = current_token.set(make_token(scope))
        try:
            assert not (await invoke("list_bookings")).error, scope
        finally:
            current_token.reset(tok)
    tok = current_token.set(make_token("Mail.Read"))
    try:
        r = await invoke("list_bookings")
    finally:
        current_token.reset(tok)
    assert r.error and "Bookings.Read.All or BookingsAppointment.ReadWrite.All" in r.text
