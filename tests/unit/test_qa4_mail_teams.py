"""QA round 4: what a model read in the mail, Teams, tasks and notes results, and what it did not understand."""

import pytest

from m365_mcp.core.html import html_to_text
from people import me

CHAT = "19:abc@thread.v2"


@pytest.fixture(autouse=True)
def _me(graph):
    me(graph)


# -- M19: a clamped limit says so in a sentence ----------------------------------------------

async def test_a_limit_above_the_cap_is_clamped_and_the_result_says_how_to_continue(graph, invoke):
    graph.on("GET", "/me/messages", {"value": [], "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skiptoken=abc"})
    r = await invoke("list_messages", limit=50)
    assert not r.error, r.text
    assert graph.calls("GET", "/me/messages")[0].params["$top"] == "25"
    assert r.data["limit_clamped_to"] == 25
    note = r.data["limit_note"]
    assert "25" in note and "cursor" in note


async def test_a_limit_within_the_cap_carries_no_note(graph, invoke):
    graph.on("GET", "/me/messages", {"value": []})
    r = await invoke("list_messages", limit=25)
    assert "limit_note" not in r.data and "limit_clamped_to" not in r.data


# -- M04/M09: thread_cut explains itself, and the quoted part can be had ---------------------

_REPLY = {
    "id": "m1", "from": {"emailAddress": {"address": "bo@x.test"}},
    "uniqueBody": {"content": "<p>New reply</p>"},
    "body": {"content": "<html><body><p>New reply</p><hr><p>QUOTED OLD</p></body></html>"}}


async def test_thread_cut_says_what_was_left_out_and_how_to_read_it(graph, invoke):
    graph.on("GET", "/me/messages/m1", _REPLY)
    r = await invoke("get_message", id="m1")
    assert "QUOTED OLD" not in r.data["body"]
    cut = r.data["thread_cut"]
    assert isinstance(cut, str) and "quoted" in cut and "include_quoted" in cut


async def test_include_quoted_keeps_the_history_in_the_body(graph, invoke):
    graph.on("GET", "/me/messages/m1", _REPLY)
    r = await invoke("get_message", id="m1", include_quoted=True)
    assert not r.error, r.text
    assert "New reply" in r.data["body"] and "QUOTED OLD" in r.data["body"]
    assert "thread_cut" not in r.data
    assert "with the quoted earlier messages" in r.data["body"].splitlines()[0]


# -- M04/M17: Graph's attachment size is not the file's size ---------------------------------

async def test_attachment_size_is_named_for_what_it_measures(graph, invoke):
    # Measured live (docs/graph-live-findings.md, attachment_size_vs_real_bytes): a 55-byte file
    # answers 317. The number includes the message's own wrapping, so it must not be called the size.
    graph.on("GET", "/me/messages/m1", {"id": "m1", "body": {"content": ""}})
    graph.on("GET", "/me/messages/m1/attachments", {"value": [
        {"@odata.type": "#microsoft.graph.fileAttachment", "id": "a1", "name": "n.txt", "size": 317}]})
    r = await invoke("get_message", id="m1", include_attachments=True, include_body=False)
    [a] = r.data["attachments"]
    assert a["size_incl_overhead"] == 317 and "approx_size" not in a and "size" not in a


async def test_a_drafts_attachment_size_is_named_the_same_way(graph, invoke):
    graph.on("POST", "/me/messages", {"id": "d1"})
    graph.on("GET", "/me/messages/d1", {"id": "d1", "subject": "S", "body": {"content": "<p>x</p>"}})
    graph.on("GET", "/me/messages/d1/attachments", {"value": [{"name": "n.txt", "size": 317}]})
    r = await invoke("save_draft", subject="S", body="x")
    assert r.data["attachments"] == [{"name": "n.txt", "size_incl_overhead": 317}]


# -- M07: a table keeps its rows ---------------------------------------------------------------

def test_a_table_reads_as_rows_with_closed_cells():
    html = ("<table><tr><th>Item</th><th>Qty</th></tr>"
            "<tr><td>Widgets</td><td>3</td></tr><tr><td>Bolts</td><td>12</td></tr></table>")
    assert html_to_text(html).splitlines() == [
        "| Item | Qty |", "| --- | --- |", "| Widgets | 3 |", "| Bolts | 12 |"]


def test_a_table_without_a_header_row_gets_no_separator():
    assert html_to_text("<table><tr><td>a</td><td>b</td></tr></table>") == "| a | b |"


def test_text_around_a_table_stays_apart_from_it():
    text = html_to_text("<p>Before</p><table><tr><td>a</td></tr></table><p>After</p>")
    assert [line for line in text.splitlines() if line] == ["Before", "| a |", "After"]


def test_a_hidden_cell_is_not_counted_in_the_row():
    html = '<table><tr><td>a</td><td style="display:none">secret</td><td>b</td></tr></table>'
    assert html_to_text(html) == "| a | b |"


# -- S05: the action error says permanent deletion is left out on purpose ---------------------

async def test_an_unknown_action_says_permanent_deletion_is_not_offered(graph, invoke):
    r = await invoke("update_messages", ids=["m1"], action="purge")
    assert r.error
    assert "one of" in r.text and "Permanent deletion is deliberately not offered" in r.text
    assert "Deleted Items" in r.text
    assert graph.writes() == []


async def test_other_argument_errors_do_not_carry_that_hint(graph, invoke):
    r = await invoke("update_messages", ids=["m1"], action="move", folder=123)
    assert r.error and "Permanent deletion" not in r.text


# -- T05/T06: an empty read says why, and the order is stated ---------------------------------

def _channels(graph):
    graph.on("GET", "/me/joinedTeams", {"value": [{"id": "t1", "displayName": "Ops"}]})
    graph.on("GET", "/teams/t1/channels", {"value": [{"id": "c1", "displayName": "General"}]})


async def test_an_empty_channel_says_no_posts_and_that_no_access_is_an_error_instead(graph, invoke):
    _channels(graph)
    graph.on("GET", "/teams/t1/channels/c1/messages", {"value": []})
    r = await invoke("get_conversation", team="Ops", channel="General")
    assert not r.error and r.data["messages"] == []
    assert "No posts" in r.data["empty"] and "error" in r.data["empty"]


async def test_an_empty_filtered_channel_names_the_filter_not_the_channel(graph, invoke):
    _channels(graph)
    graph.on("GET", "/teams/t1/channels/c1/messages", {"value": []})
    r = await invoke("get_conversation", team="Ops", channel="General", attachments_only=True)
    assert "files" in r.data["empty"] and "no posts" not in r.data["empty"]


async def test_an_empty_chat_says_so(graph, invoke):
    graph.on("GET", f"/chats/{CHAT}", {"id": CHAT, "chatType": "group", "topic": "T", "members": []})
    graph.on("GET", f"/chats/{CHAT}/messages", {"value": []})
    r = await invoke("get_conversation", chat=CHAT)
    assert "No messages" in r.data["empty"]


async def test_a_page_with_messages_has_no_empty_note(graph, invoke):
    graph.on("GET", f"/chats/{CHAT}", {"id": CHAT, "chatType": "group", "topic": "T", "members": []})
    graph.on("GET", f"/chats/{CHAT}/messages", {"value": [{
        "id": "x", "messageType": "message", "createdDateTime": "2026-10-20T10:00:00Z",
        "from": {"user": {"id": "u", "displayName": "Bo"}}, "body": {"contentType": "text", "content": "hi"}}]})
    r = await invoke("get_conversation", chat=CHAT)
    assert "empty" not in r.data


async def test_the_order_of_messages_is_stated_for_chats_and_channels(graph, invoke):
    graph.on("GET", f"/chats/{CHAT}", {"id": CHAT, "chatType": "group", "topic": "T", "members": []})
    graph.on("GET", f"/chats/{CHAT}/messages", {"value": []})
    chat = await invoke("get_conversation", chat=CHAT)
    assert "oldest first" in chat.data["order"] and "last" in chat.data["order"]
    _channels(graph)
    graph.on("GET", "/teams/t1/channels/c1/messages", {"value": []})
    channel = await invoke("get_conversation", team="Ops", channel="General")
    assert "latest activity first" in channel.data["order"]


async def test_a_read_that_stopped_short_with_more_behind_is_not_called_empty(graph, invoke):
    from test_teams import chain, msg

    _channels(graph)
    gone = {"deletedDateTime": "2026-10-23T01:00:00Z"}
    path = "/teams/t1/channels/c1/messages"
    pages = [[msg(f"d{i}", "", **gone)] for i in range(8)] + [[msg("p1", "alive")]]
    graph.on("GET", path, chain(*pages, path=path))
    r = await invoke("get_conversation", team="Ops", channel="General", messages=1)
    assert not r.error, r.text
    assert r.data["messages"] == [] and r.data["has_more"] is True
    assert "empty" not in r.data


async def test_a_channel_continued_by_cursor_is_not_told_nothing_older(graph, invoke):
    from test_teams import chain, msg

    _channels(graph)
    path = "/teams/t1/channels/c1/messages"
    graph.on("GET", path, chain([msg("p1", "a"), msg("p2", "b")], [], path=path))
    first = await invoke("get_conversation", team="Ops", channel="General", messages=1)
    assert first.data["has_more"] is True
    rest = await invoke("get_conversation", team="Ops", channel="General", messages=5, cursor=first.data["cursor"])
    assert not rest.error, rest.text
    assert "older" not in rest.data.get("empty", "")


async def test_a_chat_that_holds_only_system_events_does_not_claim_to_have_no_messages(graph, invoke):
    graph.on("GET", f"/chats/{CHAT}", {"id": CHAT, "chatType": "group", "topic": "T", "members": []})
    graph.on("GET", f"/chats/{CHAT}/messages", {"value": [
        {"id": "sys", "messageType": "systemEventMessage", "body": {"content": "<systemEventMessage/>"}}]})
    r = await invoke("get_conversation", chat=CHAT)
    assert r.data["messages"] == []
    assert "system events" in r.data["empty"] and "has none" in r.data["empty"]


async def test_find_meeting_times_with_a_high_limit_does_not_mention_a_cursor(graph, invoke):
    from test_calendar import SUGGESTION, many_people

    many_people(graph, [("Anna", "anna@x.test"), ("Bo", "bo@x.test")])
    graph.on("POST", "/me/findMeetingTimes", {"meetingTimeSuggestions": [SUGGESTION]})
    r = await invoke("find_meeting_times", people=["anna@x.test", "bo@x.test"], duration_minutes=60,
                     start="2026-10-23 09:00", end="2026-10-30 09:00", limit=50)
    assert not r.error, r.text
    assert r.data["limit_clamped_to"] == 25
    assert "cursor" not in r.text and "limit_note" not in r.data


async def test_a_list_that_ends_has_no_note_about_a_cursor(graph, invoke):
    graph.on("GET", "/me/messages", {"value": []})
    r = await invoke("list_messages", limit=50)
    assert r.data["limit_clamped_to"] == 25 and "limit_note" not in r.data


# -- N02/N06: the words a person uses lead to the argument -------------------------------------

async def _props(server, name):
    return next(t for t in await server.list_tools() if t.name == name).input_schema["properties"]


async def test_save_task_maps_remind_me_to_the_reminder_argument(server):
    assert "remind me" in (await _props(server, "save_task"))["reminder"]["description"].lower()


async def test_save_note_page_says_text_makes_no_lists(server):
    d = (await _props(server, "save_note_page"))["body_format"]["description"].lower()
    assert "html" in d and "list" in d


# -- M07 review: a cell stays on one line and one cell ---------------------------------------

def test_a_pipe_in_a_cell_is_escaped():
    assert html_to_text("<table><tr><td>a|b</td><td>c</td></tr></table>") == "| a\\|b | c |"


def test_blocks_and_breaks_in_a_cell_become_spaces_so_the_row_stays_on_one_line():
    html = "<table><tr><td><p>one</p><p>two</p></td><td>x<br>y</td></tr></table>"
    assert html_to_text(html) == "| one two | x y |"


def test_a_second_header_only_row_gets_no_second_separator():
    html = "<table><tr><th>A</th><th>B</th></tr><tr><th>C</th><th>D</th></tr><tr><td>1</td><td>2</td></tr></table>"
    text = html_to_text(html)
    assert text.count("---") == 2  # one separator row of two columns
    assert text.splitlines() == ["| A | B |", "| --- | --- |", "| C | D |", "| 1 | 2 |"]


def test_each_table_has_its_own_header_separator():
    one = "<table><tr><th>A</th></tr></table>"
    text = html_to_text(one + "<p>x</p>" + one)
    assert text.count("---") == 2


# -- Outlook's own table markup keeps its rows --------------------------------

def _outlook_cell(text, tag="td"):
    return (f'<{tag} width=200 valign=top style="width:150.0pt;border:solid windowtext 1.0pt;padding:0cm 5.4pt">'
            f'<p class=MsoNormal><span style="font-size:11.0pt;font-family:Calibri">{text}<o:p></o:p></span></p></{tag}>')


def _outlook_table(*rows):
    body = "".join("<tr style='height:15.0pt'>\n" + "\n".join(r) + "\n</tr>\n" for r in rows)
    return f'<div class=WordSection1><p class=MsoNormal>Plan<o:p></o:p></p>\n<table class=MsoTableGrid border=1>\n{body}</table>\n</div>'


def test_cells_with_paragraphs_spans_and_newlines_between_them_stay_on_their_row():
    html = _outlook_table([_outlook_cell("Task"), _outlook_cell("Owner")], [_outlook_cell("Ship"), _outlook_cell("Bo")])
    assert html_to_text(html).splitlines() == ["Plan", "", "| Task | Owner |", "| Ship | Bo |"]


def test_outlook_header_cells_get_the_separator_row():
    html = _outlook_table([_outlook_cell("Task", "th"), _outlook_cell("Owner", "th")], [_outlook_cell("Ship"), _outlook_cell("Bo")])
    assert html_to_text(html).splitlines()[2:] == ["| Task | Owner |", "| --- | --- |", "| Ship | Bo |"]


def test_a_nested_table_and_a_multi_paragraph_cell_stay_one_line_each():
    inner = "<table>\n<tr>\n<td>x</td>\n<td>y</td>\n</tr>\n</table>"
    html = ("<table>\n<tr>\n<td><p>One</p>\n<p>Two</p></td>\n<td>" + inner + "</td>\n</tr>\n</table>")
    lines = [line for line in html_to_text(html).splitlines() if line]
    assert lines[0].startswith("| One Two |")
    assert not any(line == "|" for line in lines)
