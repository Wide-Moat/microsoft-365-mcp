"""Ids and names from the model must never rewrite a Graph request. Each case here was a real
exploit against an earlier version: an id ending in "/send#" sent a draft from a reply call, ".."
deleted a whole To Do list or SharePoint list, a folder path walked into another user's drive."""

import pytest

from helpers import FakeGraph
from m365_mcp.graph.client import Graph, UnsafePath


@pytest.mark.parametrize("path", [
    "/me/messages/D/send#/createReply",
    "/me/messages/D/send?x=/createReply",
    "/me/todo/lists/L/tasks/..",
    "/sites/S/lists/L/items/..",
    "/me/messages/../mailFolders/F/move",
    "/me/drive/root:/../../../users/x/drive/root:/children",
    "/me/messages/%2e%2e/x",
    "/me/messages/a%2Fb",
    "/me/messages/a\\b",
    "/me/messages/a b",
])
async def test_unsafe_paths_never_reach_graph(path):
    import httpx
    fake = FakeGraph()
    g = Graph("t", "p", transport=httpx.MockTransport(fake))
    with pytest.raises(UnsafePath):
        await g.request("POST", path, json={})
    assert fake.seen == []
    await g.aclose()


async def test_reply_source_cannot_become_a_send(graph, invoke):
    r = await invoke("save_draft", mode="reply", source_id="DRAFT/send#", body="x")
    assert r.error and "not valid" in r.text
    assert not [s for s in graph.seen if s.path.endswith("/send")]
    assert graph.writes() == []


async def test_remove_attachment_cannot_delete_the_draft(graph, invoke):
    from m365_mcp.tools.mail import DRAFT_MARK
    graph.on("GET", "/me/messages/D#", {"id": "D", "isDraft": True,
                                        "singleValueExtendedProperties": [{"id": DRAFT_MARK, "value": "1"}]})
    r = await invoke("save_draft", id="D#", remove_attachments=["a.pdf"])
    assert r.error
    assert not [s for s in graph.seen if s.method == "DELETE"]


async def test_dotdot_task_id_cannot_delete_the_list(graph, invoke):
    r = await invoke("delete_task", id="..", source="todo", list="Groceries", user_confirmed=True)
    assert r.error
    assert not [s for s in graph.seen if s.method == "DELETE"]


async def test_dotdot_message_id_cannot_move_a_folder(graph, invoke):
    r = await invoke("update_messages", ids=["../mailFolders/F1"], action="delete")
    assert not [s for s in graph.seen if s.method == "POST"]
    assert r.data["failed"] if not r.error else True


async def test_folder_path_cannot_walk_into_another_drive(graph, invoke):
    r = await invoke("list_files", folder="../../../users/victim@x.test")
    assert r.error
    assert not [s for s in graph.seen if "/users/" in s.path]
