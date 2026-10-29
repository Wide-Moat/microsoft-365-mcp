"""The refusals every tool shares: how a Graph failure reaches the person, through a real tool call."""


async def test_a_missing_permission_is_named_with_the_administrator_step(graph, invoke):
    graph.on("GET", "/me/messages/m1", {"error": {"code": "ErrorAccessDenied",
                                                 "message": "Missing scope Mail.ReadWrite for anna@x.test"}}, status=403)
    r = await invoke("get_message", id="m1")
    assert r.error and "Mail.ReadWrite" in r.text and "administrator" in r.text
    assert "anna@" not in r.text


async def test_throttling_after_retries_says_how_long_to_wait(graph, invoke):
    graph.on("GET", "/me/messages/m1", {"error": {"code": "ApplicationThrottled"}}, status=429,
             headers={"Retry-After": "7"})
    r = await invoke("get_message", id="m1")
    assert r.error and "Try again in 7 seconds" in r.text
    assert len(graph.calls("GET", "/me/messages/m1")) == 3


async def test_a_refused_write_says_nothing_changed_with_a_reference(graph, invoke):
    graph.on("POST", "/me/mailFolders", {"error": {"code": "ErrorInvalidRequest"}}, status=400,
             headers={"request-id": "rid-400"})
    r = await invoke("save_mail_folder", name="Projects")
    assert r.error and "nothing was changed" in r.text and "rid-400" in r.text


async def test_a_server_failure_on_a_write_never_claims_nothing_changed(graph, invoke):
    graph.on("POST", "/me/mailFolders", {"error": {"code": "InternalServerError"}}, status=500,
             headers={"request-id": "rid-500"})
    r = await invoke("save_mail_folder", name="Projects")
    assert r.error and "may or may not have happened" in r.text and "rid-500" in r.text
    assert "nothing was changed" not in r.text


# -- found on a real tenant ------------------------------------------------------------------

import pytest  # noqa: E402


@pytest.mark.parametrize("tool, args", [
    ("get_message", {"id": ""}),
    ("get_event", {"id": ""}),
    ("cancel_event", {"id": "", "user_confirmed": True}),
    ("delete_task", {"id": "", "source": "todo", "list": "Tasks", "user_confirmed": True}),
    ("delete_note_page", {"id": "", "user_confirmed": True}),
    ("update_messages", {"ids": [""], "action": "mark_read"}),
])
async def test_an_empty_id_is_refused_before_graph(graph, invoke, tool, args):
    # An empty id leaves "/me/events/", which Graph answers with the whole collection.
    graph.on("GET", "/me/todo/lists", {"value": [{"id": "L1", "displayName": "Tasks"}]})
    r = await invoke(tool, **args)
    text = r.text if r.error else str(r.data.get("failed"))
    assert "not valid" in text, text
    assert "went wrong on our side" not in text
    assert [s.path for s in graph.seen if s.path.endswith("/")] == []
    assert graph.writes() == []


@pytest.mark.parametrize("tool, args", [
    ("save_event", {"id": "", "subject": "S", "start": "2026-11-09 09:00", "duration_minutes": 30}),
    ("save_event", {"id": "  ", "subject": "S", "start": "2026-11-09 09:00", "duration_minutes": 30}),
    ("save_task", {"id": "", "source": "todo", "title": "T"}),
    ("save_task", {"id": "", "source": "planner", "title": "T", "plan": "P"}),
    ("save_note_page", {"id": "", "title": "T", "section": "S", "body": "x"}),
])
async def test_an_empty_id_on_a_save_is_refused_not_read_as_new(graph, invoke, tool, args):
    # Found by QA: save_event and save_task read id "" as no id and created a new event and task.
    graph.on("GET", "/me/todo/lists", {"value": [{"id": "L1", "displayName": "Tasks", "wellknownListName": "defaultList"}]})
    graph.on("POST", "/me/todo/lists/L1/tasks", {"id": "new"}, status=201)
    graph.on("POST", "/me/events", {"id": "new"}, status=201)
    graph.on("GET", "/me/onenote/sections", {"value": [{"id": "S1", "displayName": "S"}]})
    graph.on("POST", "/me/onenote/sections/S1/pages", {"id": "new"}, status=201)
    r = await invoke(tool, **args)
    assert r.error and "An id or path in this request is empty or not valid" in r.text, r.text
    assert graph.writes() == []


@pytest.mark.parametrize("tool, args", [
    ("get_event", {"id": "a/b"}),
    ("save_event", {"id": "a/b", "subject": "S"}),
    ("respond_to_event", {"id": "x/accept", "response": "accept"}),
    ("cancel_event", {"id": "a/b", "user_confirmed": True}),
    ("save_task", {"id": "a/b", "source": "todo", "title": "T"}),
    ("save_task", {"id": "a/b", "source": "planner", "title": "T"}),
    ("delete_task", {"id": "a/b", "source": "todo", "list": "Tasks", "user_confirmed": True}),
    ("delete_task", {"id": "a\\b", "source": "planner", "user_confirmed": True}),
    ("get_note_page", {"id": "a/b"}),
    ("save_note_page", {"id": "a/b", "action": "append", "body": "x"}),
    ("delete_note_page", {"id": "a/b", "user_confirmed": True}),
])
async def test_an_id_holding_a_separator_is_not_found_before_graph(graph, invoke, tool, args):
    # Found by QA: "a/b" became two path segments, and Graph's answer ("refused the request"), or the
    # client's own DELETE allowlist ("went wrong on our side"), reached the person instead of not found.
    graph.on("GET", "/me/todo/lists", {"value": [{"id": "L1", "displayName": "Tasks"}]})
    r = await invoke(tool, **args)
    assert r.error and "It does not exist, or it is not shared with you." in r.text, r.text
    assert [s.path for s in graph.seen if "/a/" in s.path or "/x/" in s.path or "\\" in s.path] == []
    assert graph.writes() == []


async def test_a_todo_delete_of_an_id_to_do_cannot_read_is_not_found(graph, invoke):
    # Measured: DELETE of task "x" answers 500 generalException ("not a valid Base-64 string"), which
    # read as "may or may not have happened"; GET of it answers 400 ErrorInvalidIdMalformed.
    graph.on("GET", "/me/todo/lists", {"value": [{"id": "L1", "displayName": "Tasks"}]})
    graph.on("GET", "/me/todo/lists/L1/tasks/x", {"error": {"code": "invalidRequest", "message": "Invalid request",
                                                            "innerError": {"code": "ErrorInvalidIdMalformed"}}}, status=400)
    graph.on("DELETE", "/me/todo/lists/L1/tasks/x", {"error": {"code": "generalException"}}, status=500)
    r = await invoke("delete_task", id="x", source="todo", list="Tasks", user_confirmed=True)
    assert r.error and "It does not exist, or it is not shared with you." in r.text and "may or may not" not in r.text
    assert graph.writes() == []


@pytest.mark.parametrize("status, error", [
    (400, {"code": "ErrorInvalidIdMalformed", "message": "Id is malformed."}),
    (400, {"code": "invalidRequest", "message": "Invalid request", "innerError": {"code": "ErrorInvalidIdMalformed"}}),
    (400, {"code": "20112", "message": "Invalid Entity ID specified."}),
    (400, {"code": "ErrorInvalidId", "message": "The Id is invalid."}),
])
async def test_a_malformed_id_is_not_found(graph, invoke, status, error):
    graph.on("GET", "/me/messages/bogus", {"error": error}, status=status, headers={"request-id": "rid"})
    r = await invoke("get_message", id="bogus")
    assert r.error and "does not exist, or it is not shared with you" in r.text


async def test_a_server_failure_on_a_read_says_nothing_changed(graph, invoke):
    graph.on("GET", "/me/messages/m1", {"error": {"code": "InternalServerError"}}, status=500, headers={"request-id": "rid-r"})
    r = await invoke("get_message", id="m1")
    assert r.error and "nothing was changed" in r.text and "rid-r" in r.text
    assert "may or may not" not in r.text


@pytest.mark.parametrize("tool, args, allowed", [
    ("list_people", {"query": "x", "source": "bogus"}, "'all', 'contacts' or 'directory'"),
    ("update_messages", {"ids": ["m1"], "action": "explode"}, "'move'"),
])
async def test_a_value_outside_an_enum_names_the_allowed_values(graph, invoke, tool, args, allowed):
    r = await invoke(tool, **args)
    assert r.error and allowed in r.text
    assert "pydantic" not in r.text and "validation error" not in r.text and "type=" not in r.text
    assert graph.seen == []


async def test_an_unprovisioned_onedrive_is_not_a_sign_in_problem(graph, invoke):
    graph.on("GET", "/me/drive/root/children", {"error": {
        "code": "notAllowed", "message": "You do not have access to create this personal site or you do not have a valid license",
        "innerError": {"code": "provisioningNotAllowed"}}}, status=403)
    r = await invoke("list_files")
    assert r.error and "OneDrive" in r.text and "not set up" in r.text
    assert "Reconnect" not in r.text


async def test_a_service_401_with_a_valid_token_does_not_say_reconnect(graph, invoke):
    # OneNote answers 401 with its own code when the person has no OneNote; the token is fine.
    graph.on("GET", "/me/onenote/notebooks", {"error": {
        "code": "40001", "message": "The request does not contain a valid authentication token."}}, status=401)
    r = await invoke("list_notes")
    assert r.error and "Reconnect" not in r.text and "OneNote" in r.text


async def test_a_chat_with_someone_teams_cannot_reach_is_not_a_sign_in_problem(graph, invoke):
    from people import me
    me(graph)
    graph.on("GET", "/me/people", {"value": []})
    graph.on("POST", "/chats", {"error": {"code": "Unauthorized", "message": "Access to tenant is denied.",
                                          "innerError": {"code": "TenantNotFound"}}}, status=401)
    r = await invoke("send_message", to=["client@example.invalid"], body="hi", user_confirmed=True)
    assert r.error and "Reconnect" not in r.text and "cannot reach" in r.text
    assert graph.calls("POST", "/chats/.*") == []


@pytest.mark.parametrize("code", ["unauthenticated", "InvalidAuthenticationToken", "invalidAuthenticationToken", ""])
async def test_a_token_401_in_any_spelling_says_reconnect(graph, invoke, code):
    graph.on("GET", "/me/drive/root/children", {"error": {"code": code, "message": "x"}}, status=401)
    r = await invoke("list_files")
    assert r.error and "Reconnect" in r.text


# -- every refusal names the next step ---------------------------------------

from m365_mcp.core.errors import M365Error  # noqa: E402


@pytest.mark.parametrize("kind, step", [
    ("person", "list_people"), ("folder", "list_messages kind=folders"), ("team", "list_conversations kind=channel"),
    ("channel", "list_conversations kind=channel"), ("calendar", "Check the spelling"),
])
def test_a_name_that_matches_nothing_says_where_to_look(kind, step):
    assert step in M365Error("unresolved", kind=kind, query="x").text


def test_the_shared_refusals_name_a_next_step():
    assert "find it again" in M365Error("not_found").text
    assert "Check the values" in M365Error("failed", request_id="r").text


def test_one_second_is_singular():
    assert "Wait 1 second." in M365Error("rate", seconds=1).text
    assert "Try again in 1 second." in M365Error("throttled", seconds=1).text
    assert "Wait 6 seconds." in M365Error("rate", seconds=6).text


async def test_a_draft_without_recipients_or_no_longer_a_draft_says_what_to_do(graph, invoke):
    from test_mail import OWN
    graph.on("GET", "/me/messages/d1", {"id": "d1", "isDraft": True, "subject": "S", **OWN})
    r = await invoke("send_draft", id="d1", user_confirmed=True)
    assert r.error and "save_draft" in r.text and graph.writes() == []
    graph.on("GET", "/me/messages/d1", {"id": "d1", "isDraft": False, "subject": "S", **OWN})
    r = await invoke("send_draft", id="d1", user_confirmed=True)
    assert r.error and "list_messages" in r.text and graph.writes() == []
