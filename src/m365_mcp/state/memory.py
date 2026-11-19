"""The state store in this process's memory: one replica, and every test.

The same contract as RedisStore (tests/state), so behaviour does not change with the deployment. With
more than one replica it is wrong, not slow: each replica would grant its own send allowance and know only
the links it issued. The chart refuses that combination, and the server warns at start.
"""

from __future__ import annotations

import threading
import time
from collections.abc import Callable
from dataclasses import dataclass

from m365_mcp.state.store import Namer, measured


@dataclass
class _Bucket:
    tokens: float
    at: float
    full_at: float = 0.0


@dataclass
class _Link:
    value: bytes
    expires: float
    owner: str
    tag: str


class MemoryStore:
    kind = "memory"

    def __init__(self, namer: Namer | None = None, clock: Callable[[], float] = time.time) -> None:
        self._name = namer or Namer(b"\0" * 32)
        self._clock = clock
        self._buckets: dict[str, _Bucket] = {}
        self._links: dict[str, _Link] = {}
        # link name -> (held until, the tag it was issued under)
        self._busy: dict[str, tuple[float, str]] = {}
        self._spent: dict[str, float] = {}
        # Threads, not tasks: a test client serves the app in a thread of its own while the test calls in.
        self._lock = threading.Lock()

    def view(self, namer: Namer) -> MemoryStore:
        """The same store, naming things under another key (a previous sealing key, during a rotation)."""
        other = MemoryStore.__new__(MemoryStore)
        other.__dict__.update(self.__dict__)
        other._name = namer
        return other

    def bucket_count(self) -> int:
        """How many people have drawn from the allowance (tests: nothing was drawn)."""
        with self._lock:
            return len(self._buckets)

    async def ping(self) -> None:
        async def op() -> None:
            return None

        await measured("ping", op, ())

    async def take(
        self, person: str, *, capacity: float, refill_per_second: float, also: tuple[Namer, ...] = ()
    ) -> tuple[bool, float]:
        async def op() -> tuple[bool, float]:
            now = self._clock()
            keys = [self._name("bucket", person)] + [n("bucket", person) for n in also]
            with self._lock:
                # A bucket full again is the same as none: forget it, so the process does not keep every
                # person it ever saw. Each bucket knows when it is full under its own rate and capacity; a
                # bucket still refilling is kept, or forgetting it would refill it.
                for k in [k for k, v in self._buckets.items() if v.full_at <= now]:
                    del self._buckets[k]
                buckets = []
                for key in keys:
                    b = self._buckets.get(key)
                    if b is None:
                        b = self._buckets[key] = _Bucket(capacity, now)
                    b.tokens = min(capacity, b.tokens + max(0.0, now - b.at) * refill_per_second)
                    b.at = now
                    buckets.append(b)
                # From every bucket of the person (one per key during a rotation) or from none.
                fewest = min(b.tokens for b in buckets)
                taken = fewest >= 1
                for b in buckets:
                    b.tokens -= 1 if taken else 0
                    b.full_at = now + (capacity - b.tokens) / refill_per_second
                return taken, fewest - (1 if taken else 0)

        return await measured("take", op, ())

    def _prune(self, now: float) -> None:
        for k in [k for k, v in self._links.items() if v.expires <= now]:
            del self._links[k]
        for k in [k for k, (exp, _tag) in self._busy.items() if exp <= now]:
            del self._busy[k]
        for k in [k for k, exp in self._spent.items() if exp <= now]:
            del self._spent[k]

    async def link_put(
        self, link_id: str, value: bytes, *, expires: float, owner: str, owner_cap: int, total_cap: int, tag: str
    ) -> bool:
        async def op() -> bool:
            now = self._clock()
            key, own = self._name("link", link_id), self._name("owner", owner)
            group = self._name("tag", tag) if tag else ""
            if expires <= now:
                return False
            with self._lock:
                self._prune(now)
                mine = sum(1 for v in self._links.values() if v.owner == own)
                if len(self._links) >= total_cap or mine >= owner_cap or key in self._links:
                    return False
                self._links[key] = _Link(value, expires, own, group)
                return True

        return await measured("link_put", op, ())

    async def link_get(self, link_id: str) -> bytes | None:
        async def op() -> bytes | None:
            now = self._clock()
            key = self._name("link", link_id)
            with self._lock:
                hit = self._links.get(key)
                if hit is None:
                    return None
                if hit.expires <= now:
                    del self._links[key]
                    return None
                return hit.value

        return await measured("link_get", op, ())

    async def link_claim(self, link_id: str, *, hold: float) -> bool:
        async def op() -> bool:
            now = self._clock()
            key = self._name("link", link_id)
            with self._lock:
                self._prune(now)
                if key in self._busy or key in self._spent:
                    return False
                held = self._links.get(key)
                self._busy[key] = (now + hold, held.tag if held is not None else "")
                return True

        return await measured("link_claim", op, ())

    async def link_release(self, link_id: str) -> None:
        async def op() -> None:
            with self._lock:
                self._busy.pop(self._name("link", link_id), None)

        await measured("link_release", op, ())

    async def link_consume(self, link_id: str, *, until: float) -> None:
        async def op() -> None:
            key = self._name("link", link_id)
            with self._lock:
                self._busy.pop(key, None)
                if until > self._clock():
                    self._spent[key] = until
                self._links.pop(key, None)

        await measured("link_consume", op, ())

    async def link_spent(self, link_id: str) -> bool:
        async def op() -> bool:
            key = self._name("link", link_id)
            with self._lock:
                exp = self._spent.get(key)
                return exp is not None and exp > self._clock()

        return await measured("link_spent", op, ())

    async def link_busy_under(self, tag: str) -> bool:
        async def op() -> bool:
            if not tag:
                return False
            now = self._clock()
            group = self._name("tag", tag)
            with self._lock:
                self._prune(now)
                # The claim carries its tag: an upload admitted just before its link expired still runs.
                return any(t == group for _until, t in self._busy.values())

        return await measured("link_busy_under", op, ())

    async def link_revoke(self, tag: str) -> int:
        async def op() -> int:
            if not tag:
                return 0
            now = self._clock()
            group = self._name("tag", tag)
            with self._lock:
                self._prune(now)
                keys = [k for k, v in self._links.items() if v.tag == group]
                for k in keys:
                    self._spent[k] = self._links.pop(k).expires
                    self._busy.pop(k, None)
                # An upload still running past its link's expiry is spent too.
                for k in [k for k, (_until, t) in self._busy.items() if t == group]:
                    self._spent[k] = self._busy.pop(k)[0]
                return len(keys)

        return await measured("link_revoke", op, ())

    async def aclose(self) -> None:
        return None
