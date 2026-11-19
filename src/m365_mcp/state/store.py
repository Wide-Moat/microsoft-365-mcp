"""The shared state store: its interface, the names it stores things under, and its instruments.

Authorisation is not here: the gateway forwards each person's Graph token, and this server keeps none.
What replicas must agree on is small, and all of it short-lived:

| What | Operation |
|---|---|
| the send allowance of each person | a token bucket, taken atomically |
| a short link's id -> its sealed value | stored with its expiry, at most so many per person and in all |
| an upload link receiving a body, or one that delivered | a busy mark with a deadline; a spent mark |
| the links that write into one draft | revoked together when the draft is sent |

Two implementations pass one contract suite (tests/state): `MemoryStore` (one process: tests, a single
replica) and `RedisStore` (Redis or Valkey, standalone or behind Sentinel).

No name in the store carries a person, an address or an id in clear: every identifier becomes
HMAC-SHA256 under a key derived from the sealing key for this purpose alone, truncated.
"""

from __future__ import annotations

import hashlib
import hmac
import time
from collections.abc import Awaitable, Callable
from typing import Protocol, TypeVar

from m365_mcp.observability import STATE_ERRORS, STATE_SECONDS, STATE_UP

T = TypeVar("T")


class StoreUnavailable(Exception):
    """The store did not answer in time, or refused. Callers fail closed: no send, no new link."""


class Namer:
    """Identifiers to store names: keyed, so nobody with the store alone can test a guess."""

    def __init__(self, key: bytes) -> None:
        if len(key) < 32:
            raise ValueError("the store's naming key must be at least 32 bytes")
        self._key = key

    def __call__(self, kind: str, raw: str) -> str:
        return hmac.new(self._key, f"{kind}\n{raw}".encode(), hashlib.sha256).hexdigest()[:32]


class StateStore(Protocol):
    kind: str  # "memory" or "redis"

    def view(self, namer: Namer) -> StateStore:
        """The same store, naming things under another key."""
        ...

    async def ping(self) -> None: ...

    async def take(self, person: str, *, capacity: float, refill_per_second: float) -> tuple[bool, float]:
        """Take one token from the person's bucket. Returns whether it was taken and the tokens left after
        (when refused, the tokens there are, below one)."""
        ...

    async def link_put(
        self, link_id: str, value: bytes, *, expires: float, owner: str, owner_cap: int, total_cap: int, tag: str
    ) -> bool:
        """Store `value` under `link_id` until `expires` (epoch seconds). False, storing nothing, while the
        owner already holds `owner_cap` live links or the store `total_cap`: a live link is never evicted."""
        ...

    async def link_get(self, link_id: str) -> bytes | None: ...

    async def link_claim(self, link_id: str, *, hold: float) -> bool:
        """Mark an upload link busy for at most `hold` seconds. False while another upload holds it or
        after one delivered."""
        ...

    async def link_release(self, link_id: str) -> None: ...

    async def link_consume(self, link_id: str, *, until: float) -> None:
        """The upload delivered: the link is spent until `until`, and its stored value goes."""
        ...

    async def link_spent(self, link_id: str) -> bool: ...

    async def link_busy_under(self, tag: str) -> bool: ...

    async def link_revoke(self, tag: str) -> int:
        """Spend every live link stored under `tag`, busy ones too. Returns how many."""
        ...

    async def aclose(self) -> None: ...


async def measured(op: str, call: Callable[[], Awaitable[T]], errors: tuple[type[BaseException], ...]) -> T:
    """Run one store operation with its duration and outcome recorded; `errors` become StoreUnavailable."""
    t0 = time.perf_counter()
    try:
        result = await call()
    except errors as e:
        STATE_ERRORS.labels(op).inc()
        STATE_UP.one().set(0)
        raise StoreUnavailable(op) from e
    finally:
        STATE_SECONDS.labels(op).observe(time.perf_counter() - t0)
    STATE_UP.one().set(1)
    return result
