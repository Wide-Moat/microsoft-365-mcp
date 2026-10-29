"""A confirmation is the JSON boolean true and nothing else (rules 20, 21).

Pydantic's lax mode turns "yes", "true", 1 and "1" into True, and the SDK parses a JSON string
before validating; either way a model that wrote a word where a boolean belongs sent mail (QA
round 2, SN10). The list of tools is read from the registry, so a new tool with a confirmation
parameter is covered without editing this file.
"""

from __future__ import annotations

import inspect

import pytest

from m365_mcp.registry import all_tools
from m365_mcp.server import CONFIRMATION_PARAMS

from test_mail import OWN, person

NOT_TRUE = ["yes", "true", "True", 1, "1", 1.0, None, [True], {"value": True}]

CONFIRMING = sorted(
    (spec.name, p)
    for spec in all_tools().values()
    for p in inspect.signature(spec.fn).parameters
    if p in CONFIRMATION_PARAMS
)


def test_the_registry_has_the_confirming_tools_it_should():
    # Counted from the registry, so a rename that hid every parameter from the test below fails here.
    names = {t for t, _ in CONFIRMING}
    assert {"send_draft", "send_message", "save_event", "cancel_event", "share_file",
            "save_automatic_reply", "delete_contact"} <= names
    assert len(CONFIRMING) >= 14


@pytest.mark.parametrize("tool,param", CONFIRMING)
@pytest.mark.parametrize("value", NOT_TRUE, ids=repr)
async def test_a_confirmation_that_is_not_the_boolean_true_is_refused(graph, invoke, tool, param, value):
    r = await invoke(tool, **{param: value})
    assert r.error
    assert f"{param} must be the boolean true" in r.text, r.text
    assert graph.writes() == []


async def test_send_draft_with_the_word_yes_sends_nothing(graph, invoke):
    graph.on("GET", "/me/messages/d1", {"id": "d1", "isDraft": True, "subject": "S",
                                        "toRecipients": [{"emailAddress": {"address": "bo@x.test"}}], **OWN})
    graph.on("PATCH", "/me/messages/d1", {})
    graph.on("POST", "/me/messages/d1/send", None, status=202)
    person(graph, "Bo", "bo@x.test")
    r = await invoke("send_draft", id="d1", user_confirmed="yes")
    assert r.error and "user_confirmed must be the boolean true" in r.text
    assert graph.writes() == []
    ok = await invoke("send_draft", id="d1", user_confirmed=True)
    assert not ok.error, ok.text
    assert len(graph.calls("POST", "/me/messages/d1/send")) == 1


async def test_false_is_still_a_boolean_and_reaches_the_tool(graph, invoke):
    graph.on("GET", "/me/messages/d1", {"id": "d1", "isDraft": True, "subject": "S",
                                        "toRecipients": [{"emailAddress": {"address": "bo@x.test"}}], **OWN})
    r = await invoke("send_draft", id="d1", user_confirmed=False)
    assert r.error and "confirm" in r.text and "boolean true" not in r.text
    assert graph.writes() == []
