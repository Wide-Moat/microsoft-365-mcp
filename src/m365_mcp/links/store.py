"""Where a short link's signed value waits until it is used.

A signed value carries the person's sealed Graph token and runs to about 6,300 characters; every one
of them in a tool result costs the model's context, twice per link once text and structured copy both
hold it. The value now stays here, in this process, under a random id, and the link carries only the
id: about 80 characters.

The store is the process's memory, on purpose. The links live minutes (download 300 s, upload 180 s,
never past LINK_TTL_CAP), the server runs one replica (a deployment must not run more while links
are held here), and a restart loses links that the person then asks for again. Nothing about them is
worth a database. 5,000 live links at ~6.5 KB each is ~33 MB, inside the pod's 384 Mi; past that a new
link is refused rather than an old one dropped.
"""

from __future__ import annotations

import secrets
import threading
import time
from collections.abc import Callable


class LinksFull(Exception):
    """Every slot holds a link that has not expired: refuse a new one rather than drop one in use."""


class LinkStore:
    def __init__(self, max_entries: int = 5_000, per_owner: int = 300,
                 clock: Callable[[], float] = time.time) -> None:
        self._max = max_entries
        self._per_owner = per_owner  # one person looping a tool must not fill the store for everyone
        self._clock = clock
        self._items: dict[str, tuple[str, float, str, str]] = {}  # id -> (value, expiry, owner, tag)
        # Upload links move one file, once: a link receiving a body is busy, and a link that delivered
        # one is spent until it would have expired anyway.
        self._busy: set[str] = set()
        self._spent: dict[str, float] = {}
        self._lock = threading.Lock()

    def _prune(self, now: float) -> None:
        for key in [k for k, (_, exp, _o, _t) in self._items.items() if exp <= now]:
            del self._items[key]

    def put(self, value: str, *, expires: float, owner: str = "", tag: str = "") -> str:
        now = self._clock()
        with self._lock:
            mine = sum(1 for _, exp, o, _t in self._items.values() if o == owner and exp > now)
            if len(self._items) >= self._max or mine >= self._per_owner:
                self._prune(now)
            if len(self._items) >= self._max or mine >= self._per_owner:
                raise LinksFull()
            key = secrets.token_urlsafe(16)  # 128 random bits, 22 characters
            self._items[key] = (value, expires, owner, tag)
            return key

    def get(self, key: str, now: float | None = None) -> str | None:
        if not key or len(key) > 64:
            return None
        now = self._clock() if now is None else now
        with self._lock:
            hit = self._items.get(key)
            if hit is None:
                return None
            value, exp, _owner, _tag = hit
            if exp <= now:
                del self._items[key]
                return None
            return value

    def claim(self, key: str) -> bool:
        """Mark a link as receiving an upload. False when another upload holds it or it was used."""
        now = self._clock()
        with self._lock:
            for k in [k for k, exp in self._spent.items() if exp <= now]:
                del self._spent[k]
            if key in self._busy or key in self._spent:
                return False
            self._busy.add(key)
            return True

    def release(self, key: str) -> None:
        """The upload did not deliver: the link may be used again until it expires."""
        with self._lock:
            self._busy.discard(key)

    def consume(self, key: str, until: float) -> None:
        """The upload delivered: the link is spent, and a short id's stored value goes."""
        with self._lock:
            self._busy.discard(key)
            self._spent[key] = until
            self._items.pop(key, None)

    def spent(self, key: str) -> bool:
        with self._lock:
            exp = self._spent.get(key)
            return exp is not None and exp > self._clock()

    def busy_under(self, tag: str) -> bool:
        """Whether a link issued under `tag` is receiving an upload right now."""
        if not tag:
            return False
        with self._lock:
            return any(k in self._busy and t == tag for k, (_, _e, _o, t) in self._items.items())

    def revoke(self, tag: str) -> int:
        """Spend every link issued under `tag` (a draft that was sent): busy ones too, so an upload in
        progress cannot be followed by another. Returns how many were revoked."""
        if not tag:
            return 0
        with self._lock:
            keys = [k for k, (_, _e, _o, t) in self._items.items() if t == tag]
            for k in keys:
                self._spent[k] = self._items.pop(k)[1]
            return len(keys)
