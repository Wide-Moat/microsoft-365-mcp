"""The stores the contract suite runs against: this process's memory always, and real Redis, Valkey,
Valkey behind three Sentinels (passwords everywhere) and Redis behind one Sentinel (no password, the way
a cluster's operator often runs it) when scripts/state-contract.sh started them (never a fake: a fake would agree with
whatever the Lua scripts get wrong).

Without those servers the real-store cases are skipped locally; CI sets M365_TEST_STORES_REQUIRED=1, which
turns a missing server into a failure, so the suite cannot pass having tested memory alone.
"""

from __future__ import annotations

import os
import uuid
from collections.abc import AsyncIterator

import pytest

from m365_mcp.state.memory import MemoryStore
from m365_mcp.state.remote import RedisStore
from m365_mcp.state.store import Namer

NAMER_KEY = bytes(range(64, 96))
REAL = ("redis", "valkey", "sentinel", "redis-sentinel")


def _env(name: str) -> str | None:
    return os.environ.get(name) or None


# Which variable names each store's address (pairs, not a mapping: a scanner reads `"…key": "…"` as a secret).
SERVER_VARIABLES = (
    ("redis", "M365_TEST_REDIS_URL"),
    ("valkey", "M365_TEST_VALKEY_URL"),
    ("sentinel", "M365_TEST_SENTINELS"),
    ("redis-sentinel", "M365_TEST_REDIS_SENTINELS"),
)


def _available(kind: str) -> bool:
    return bool(_env(dict(SERVER_VARIABLES)[kind]))


def make(kind: str, prefix: str) -> RedisStore:
    namer = Namer(NAMER_KEY)
    password = _env("M365_TEST_STATE_PASSWORD")
    if kind == "redis":
        return RedisStore.standalone(os.environ["M365_TEST_REDIS_URL"], namer, prefix, password=password)
    if kind == "valkey":
        return RedisStore.standalone(os.environ["M365_TEST_VALKEY_URL"], namer, prefix, password=password)
    if kind == "redis-sentinel":
        h, _, p = os.environ["M365_TEST_REDIS_SENTINELS"].rpartition(":")
        return RedisStore.sentinels([(h, int(p))], "mymaster", namer, prefix)
    hosts = [(h, int(p)) for h, _, p in (x.rpartition(":") for x in os.environ["M365_TEST_SENTINELS"].split(","))]
    return RedisStore.sentinels(
        hosts,
        os.environ.get("M365_TEST_SENTINEL_MASTER", "mymaster"),
        namer,
        prefix,
        password=password,
        sentinel_password=_env("M365_TEST_SENTINEL_PASSWORD"),
    )


@pytest.fixture(params=("memory", *REAL))
async def replicas(request) -> AsyncIterator[tuple]:
    """Two replicas' views of one store. In memory both are the one process's store."""
    kind = request.param
    if kind == "memory":
        m = MemoryStore(Namer(NAMER_KEY))
        yield m, m
        return
    if not _available(kind):
        if os.environ.get("M365_TEST_STORES_REQUIRED") == "1":
            pytest.fail(f"M365_TEST_STORES_REQUIRED=1 and no {kind} server: run scripts/state-contract.sh")
        pytest.skip(f"no {kind} server; scripts/state-contract.sh starts one")
    prefix = f"contract-{uuid.uuid4().hex[:8]}:"
    a, b = make(kind, prefix), make(kind, prefix)
    try:
        yield a, b
    finally:
        await a.aclose()
        await b.aclose()


@pytest.fixture
def store(replicas):
    return replicas[0]
