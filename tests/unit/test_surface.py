"""The tool surface — names, areas, annotations and input schemas — is pinned in a committed lock file,
so a change to what the model is offered shows up in review as a change to that file."""

import asyncio
import copy
import json
from pathlib import Path

from m365_mcp.surface import REGENERATE, drift, render

LOCK = Path(__file__).resolve().parents[2] / "tool-surface.lock.json"


def current():
    return asyncio.run(render())


def test_the_registry_matches_the_committed_lock():
    problems = drift(json.loads(LOCK.read_text()), current())
    assert not problems, "\n".join(problems) + f"\n{REGENERATE}"


def test_the_lock_is_the_exact_output_of_the_generator():
    assert LOCK.read_text() == json.dumps(current(), indent=2, sort_keys=True) + "\n", REGENERATE


def test_the_lock_is_not_empty():
    tools = current()["tools"]
    from m365_mcp.registry import all_tools
    assert len(tools) == len(all_tools()) > 30


def test_drift_names_each_kind_of_change():
    lock = current()
    bad = copy.deepcopy(lock)
    name = sorted(bad["tools"])[0]
    bad["tools"][name]["input_schema_sha256"] = "0" * 64
    bad["tools"][name]["annotations"]["readOnlyHint"] = not bad["tools"][name]["annotations"]["readOnlyHint"]
    bad["tools"][name]["gates"] = {"include_everything": "switch rooms"}
    bad["tools"]["list_everything"] = dict(bad["tools"][name])
    removed = sorted(bad["tools"])[-1]
    del bad["tools"][removed]
    problems = "\n".join(drift(lock, bad))
    assert f"{name}: input schema changed" in problems
    assert f"{name}: annotations changed" in problems
    assert f"{name}: gated parameters changed" in problems
    assert "list_everything: not in the lock" in problems
    assert f"{removed}: in the lock but no longer registered" in problems


def test_no_drift_is_no_problems():
    assert drift(current(), current()) == []
