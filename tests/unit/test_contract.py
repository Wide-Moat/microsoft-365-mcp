"""The manifesto's rules on the tool catalogue, checked on every registered tool rather than on
a hand-picked few. A new tool that breaks one of them fails here before review sees it."""

import ast
import inspect
import re
from pathlib import Path

import pytest

from m365_mcp.budget import LIMIT_AREA, LIMIT_TOOL, LIMIT_TOTAL, measure
from m365_mcp.registry import ALLOWED_VERBS, AREAS, all_tools, enabled

SPECS = all_tools()
SRC = Path(__file__).resolve().parents[2] / "src" / "m365_mcp"


def test_the_catalogue_is_not_empty():
    # A broken import that registers nothing would make every test below pass vacuously.
    assert len(SPECS) >= 9


@pytest.mark.parametrize("name", sorted(SPECS))
def test_name_uses_an_allowed_verb_and_no_prefix(name):
    assert name.startswith(ALLOWED_VERBS), name
    assert not re.match(r"^(\w+_)?(m365|graph|mcp|ms|microsoft)_", name), name


@pytest.mark.parametrize("name", sorted(SPECS))
def test_description_reads_like_a_man_page(name):
    d = SPECS[name].description
    sentences = [s for s in re.split(r"(?<=[.!?])\s+", d) if s.strip()]
    assert len(sentences) >= 3, f"{name}: {len(sentences)} sentences"
    name_words = set(name.split("_"))
    first = [w.strip(".,:;").lower() for w in sentences[0].split()]
    assert len([w for w in first if w not in name_words]) >= 5, f"{name}: the first sentence only restates the name"
    assert not re.search(r"\b[A-Z]{4,}\b(?<!HTML)(?<!JSON)", re.sub(r"\b(URL|HTML|JSON|MIME|PDF|VTT|MP4|KQL|UTC|BCC|CSS)\b", "", d)), \
        f"{name}: shouted emphasis"


@pytest.mark.parametrize("name", sorted(SPECS))
def test_irreversible_tools_say_so(name):
    s = SPECS[name]
    if s.irreversible:
        assert re.search(r"cannot be undone|irreversible|not be undone", s.description, re.I), name


@pytest.mark.parametrize("name", sorted(SPECS))
def test_annotations_follow_behaviour(name):
    s = SPECS[name]
    if s.read_only:
        assert not s.notifies_others and not s.irreversible, name
    if s.notifies_others or s.irreversible:
        assert not s.idempotent, name


def _source(fn):
    return inspect.getsource(fn)


@pytest.mark.parametrize("name", sorted(n for n, s in SPECS.items() if s.notifies_others))
def test_tools_that_notify_others_go_through_the_guard(name):
    src = _source(SPECS[name].fn)
    assert "c.guard(" in src, f"{name} notifies others but never calls the guard"
    sig = inspect.signature(SPECS[name].fn)
    assert "user_confirmed" in sig.parameters, name


@pytest.mark.parametrize("name", sorted(n for n, s in SPECS.items() if s.read_only))
def test_read_only_tools_do_not_write(name):
    src = _source(SPECS[name].fn)
    for write in ("graph.post(", "graph.patch(", "graph.delete(", '"POST"', '"PATCH"', '"PUT"', '"DELETE"'):
        assert write not in src, f"{name} is read-only but calls {write}"


@pytest.mark.parametrize("name", sorted(SPECS))
def test_arguments_are_flat_and_described(name):
    sig = inspect.signature(SPECS[name].fn)
    for p in sig.parameters.values():
        ann = str(p.annotation)
        assert "dict[" not in ann, f"{name}.{p.name} is a nested object"
        self_explanatory = p.name in {
            "id", "ids", "query", "limit", "cursor", "subject", "body", "cc", "bcc", "name", "title", "start", "end",
            "unread_only", "has_attachments", "flagged", "include_body", "importance", "action", "body_format",
            "status", "comment", "message", "response", "notes", "role", "since", "until", "values", "fields",
        }
        if not self_explanatory:
            assert "Field(" in ann or "description=" in ann or "Literal[" in ann or ann in ("<class 'bool'>",) or \
                p.name.startswith("include_"), f"{name}.{p.name} needs a description"


@pytest.mark.parametrize("name", sorted(n for n, s in SPECS.items() if s.returns_list))
def test_lists_take_a_capped_limit(name):
    sig = inspect.signature(SPECS[name].fn)
    assert "limit" in sig.parameters, name
    assert "clamp_limit(" in _source(SPECS[name].fn), f"{name} does not cap its limit"


def test_every_area_is_known_and_bookings_is_off_by_default():
    assert {s.area for s in SPECS.values()} <= set(AREAS)
    assert AREAS["bookings"] is False


def test_a_switched_off_area_publishes_nothing():
    names = {s.name for s in enabled(frozenset({"account"}))}
    assert names == {n for n, s in SPECS.items() if s.area == "account"}


async def test_a_tool_of_an_area_switched_off_is_refused(settings, token):
    from dataclasses import replace

    from m365_mcp.server import M365Server
    server = M365Server(replace(settings, areas=frozenset({"account"})))
    names = {t.name for t in await server.list_tools()}
    assert "send_draft" not in names


async def test_per_person_area_header_hides_and_refuses(server, token):
    from mcp.types import CallToolRequestParams

    from m365_mcp.identity import current_areas
    current_areas.set(frozenset({"calendar"}))
    names = {t.name for t in await server.list_tools()}
    assert "list_messages" not in names and "get_connection" in names
    res = await server._handle_call_tool(None, CallToolRequestParams(name="list_messages", arguments={}))
    assert res.is_error and "switched off" in res.content[0].text


def test_every_area_publishes_a_tool_or_a_parameter():
    # An area whose last tool was merged away must still switch something off, or the switch lies.
    gated = {name for s in SPECS.values() for kind, name in s.gates.values() if kind == "area"}
    assert set(AREAS) == {s.area for s in SPECS.values()} | gated


@pytest.mark.parametrize("name", sorted(n for n, s in SPECS.items() if s.gates))
def test_gated_parameters_exist_and_name_a_known_area_or_switch(name):
    from m365_mcp.config import SWITCHES
    params = inspect.signature(SPECS[name].fn).parameters
    for param, (kind, gate) in SPECS[name].gates.items():
        assert param in params, f"{name}.{param}"
        assert params[param].default in (None, False), f"{name}.{param} must default to off"
        assert gate in (AREAS if kind == "area" else SWITCHES), f"{name}.{param}: {kind} {gate}"


async def _call(server, name, **args):
    from mcp.types import CallToolRequestParams
    res = await server._handle_call_tool(None, CallToolRequestParams(name=name, arguments=args))
    return res.is_error, "".join(getattr(c, "text", "") for c in res.content)


async def test_a_parameter_behind_an_off_switch_is_hidden_and_refused(settings, token, graph):
    from dataclasses import replace

    from m365_mcp.server import M365Server
    server = M365Server(replace(settings, switches=frozenset()))
    tools = {t.name: t.input_schema["properties"] for t in await server.list_tools()}
    assert "include_transcript" not in tools["get_event"] and "include_body" in tools["get_event"]
    assert "rooms" not in tools["find_meeting_times"] and "people" in tools["find_meeting_times"]
    error, text = await _call(server, "get_event", id="e1", include_transcript=True)
    assert error and "switched off" in text
    error, text = await _call(server, "find_meeting_times", people=["a@x.test"], duration_minutes=30, rooms=True)
    assert error and "switched off" in text
    assert graph.seen == []
    # Sent as off, it is no request for the feature, and the call goes on.
    graph.on("GET", "/me/events/e1", {"id": "e1", "subject": "S"})
    error, text = await _call(server, "get_event", id="e1", include_transcript=False)
    assert not error, text


async def test_a_parameter_of_an_area_off_for_the_person_is_hidden_and_refused(server, token, graph):
    from m365_mcp.identity import current_areas
    tools = {t.name: t.input_schema["properties"] for t in await server.list_tools()}
    assert "include_recording" in tools["get_event"]
    current_areas.set(frozenset({"calendar"}))
    tools = {t.name: t.input_schema["properties"] for t in await server.list_tools()}
    assert "include_recording" not in tools["get_event"] and "include_transcript" in tools["get_event"]
    error, text = await _call(server, "get_event", id="e1", include_recording=True)
    assert error and "recordings" in text and "switched off" in text
    assert graph.seen == []


async def test_a_parameter_of_an_area_off_for_the_installation_is_hidden(settings, token):
    from dataclasses import replace

    from m365_mcp.server import M365Server
    server = M365Server(replace(settings, areas=settings.areas - {"recordings"}))
    tools = {t.name: t.input_schema["properties"] for t in await server.list_tools()}
    assert "include_recording" not in tools["get_event"]
    error, text = await _call(server, "get_event", id="e1", include_recording=True)
    assert error and "recordings" in text


async def test_the_catalogue_fits_the_budget():
    m = await measure(all_on=False)
    assert m["total"] <= LIMIT_TOTAL, m["total"]
    for area, n in m["areas"].items():
        assert n <= LIMIT_AREA, (area, n)
    for r in m["tools"]:
        assert r["tokens"] <= LIMIT_TOOL, r


def test_only_the_graph_client_speaks_http():
    offenders = []
    for path in SRC.rglob("*.py"):
        rel = path.relative_to(SRC).as_posix()
        if rel.startswith("graph/"):
            continue
        tree = ast.parse(path.read_text())
        for node in ast.walk(tree):
            if isinstance(node, ast.Attribute) and isinstance(node.value, ast.Name) and node.value.id == "httpx" \
                    and node.attr in {"get", "post", "put", "patch", "delete", "request", "AsyncClient", "Client"}:
                offenders.append(f"{rel}:{node.lineno}")
    assert offenders == []


def test_no_mailbox_or_user_selector_anywhere():
    for path in (SRC / "tools").rglob("*.py"):
        src = path.read_text()
        assert not re.search(r"/users/\{[^}]+\}/(messages|mailFolders|events|calendar|drive|chats|contacts|todo|onenote)", src), path
        for fn in ast.walk(ast.parse(src)):
            if isinstance(fn, ast.AsyncFunctionDef):
                assert not {a.arg for a in fn.args.args + fn.args.kwonlyargs} & {"mailbox", "user_id", "tenant"}, \
                    f"{path.name}:{fn.name}"


def test_forbidden_capabilities_are_absent():
    banned = ["messageRules", "forwardingSmtpAddress", "/delegates", "calendarPermissions", "permanentDelete",
              '"anonymous"', "exportItems", "/delta"]
    for path in SRC.rglob("*.py"):
        src = path.read_text()
        for b in banned:
            assert b not in src, f"{path.name} mentions {b}"


def test_every_error_key_used_exists():
    from m365_mcp.core.errors import MESSAGES
    used = set()
    for path in SRC.rglob("*.py"):
        used |= set(re.findall(r'M365Error\("([a-z_]+)"', path.read_text()))
    assert used and used <= set(MESSAGES), used - set(MESSAGES)


async def test_an_argument_the_tool_does_not_have_is_refused_before_graph(server, token, graph):
    # QA: save_event(recurrence="weekly") made one single event and reported success; send_message
    # (attachments_upload=1) sent a message without the file the caller thought it attached.
    error, text = await _call(server, "save_event", subject="S", start="2026-10-30 10:00", duration_minutes=30,
                              recurrence="weekly")
    assert error and "recurrence" in text and "Nothing was done" in text
    error, text = await _call(server, "send_message", to=["19:chat"], body="hi", attachments_upload=1,
                              user_confirmed=True)
    assert error and "attachments_upload" in text
    assert graph.seen == []


async def test_a_limit_below_one_is_refused(server, token, graph):
    graph.on("GET", "/me/calendarView", {"value": []})
    graph.on("GET", "/me/todo/lists", {"value": []})
    for name, args in (("list_events", {"start": "2026-10-23", "end": "2026-10-24"}), ("list_tasks", {"query": "x"}),
                       ("list_conversations", {})):
        for bad in (0, -5):
            error, text = await _call(server, name, limit=bad, **args)
            assert error and "limit is 1 to 25" in text, (name, text)
    assert graph.writes() == [] and graph.calls("GET", "/me/calendarView") == []
