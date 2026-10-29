"""Live suite wiring: everything here is marked `live`, and skipped unless two signed-in test users exist.

Run `python scripts/live_login.py` first (docs/live-testing.md), then `pytest -m live`. Without a
usable token file every test reports as skipped with the reason, never as a failure.
"""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from livekit import Live, load_tokens, token_problem

HERE = Path(__file__).parent


def _tokens_path(config: pytest.Config) -> Path:
    return Path(os.environ.get("M365_LIVE_TOKENS") or config.rootpath / ".live-tokens.json")


def _findings_path(config: pytest.Config) -> Path:
    return Path(os.environ.get("M365_LIVE_FINDINGS") or config.rootpath / "live-findings.json")


@pytest.hookimpl(tryfirst=True)
def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    path = _tokens_path(config)
    problem = token_problem(load_tokens(path))
    for item in items:
        if HERE not in Path(str(item.path)).parents:
            continue
        item.add_marker(pytest.mark.live)
        if problem:
            item.add_marker(pytest.mark.skip(
                reason=f"live tenant not available: {problem} ({path}); run scripts/live_login.py [--refresh]"))


@pytest.fixture(scope="session")
def live(request: pytest.FixtureRequest):
    config = request.config
    tokens = load_tokens(_tokens_path(config))
    problem = token_problem(tokens)
    if problem:  # only reachable when the tokens changed after collection
        pytest.skip(f"live tenant not available: {problem}")
    box = Live(tokens, _findings_path(config))
    yield box
    box.close()


@pytest.fixture(autouse=True)
def _live_runtime(live: Live) -> None:
    """Unit tests configure their own runtime; this puts the live one back for every live test."""
    live.install()


@pytest.fixture
def invoke(live: Live):
    return live.invoke


@pytest.fixture
def as_a(live: Live):
    return live.acting("a")


@pytest.fixture
def as_b(live: Live):
    return live.acting("b")
