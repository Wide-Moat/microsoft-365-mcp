"""Where a short link's signed value waits until it is used.

A signed value carries the person's sealed Graph token and runs to about 6,300 characters; every one
of them in a tool result costs the model's context, twice per link once text and structured copy both
hold it. The value stays in the shared state store under a random id, and the link carries only the id:
about 80 characters. A link issued by one replica is used through any other.

The store never holds a usable link, nor anything readable. The value is AES-GCM ciphertext under a key
derived from both the sealing key's store sub-key and the id, and stored under a keyed hash of the id:
someone who reads the store (one without a password, say) has neither the id nor the key to open what is
there, so cannot download or upload through it.

The links live minutes (download 300 s, upload 180 s, never past LINK_TTL_CAP). One person may hold 300
open links and the store 5,000 in all; past either, a new link is refused rather than an old one dropped.
Upload links move one file, once: a link receiving a body is busy, and a link that delivered one is spent
until it would have expired anyway. Every such mark is one atomic operation in the store, so two replicas
cannot both accept a body for one link.
"""

from __future__ import annotations

import os
import secrets

from cryptography.exceptions import InvalidTag
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.hkdf import HKDF

from m365_mcp.state.memory import MemoryStore
from m365_mcp.state.store import Namer, StateStore

# How long an upload may hold its link busy: the longest a body may take to arrive (links/routes.py,
# BODY_TOTAL_TIMEOUT) and then to reach Microsoft. A replica that dies mid-upload frees the link after it.
UPLOAD_HOLD = 1200.0


class LinksFull(Exception):
    """Every slot holds a link that has not expired: refuse a new one rather than drop one in use."""


def _box(secret: bytes, link_id: str) -> AESGCM:
    key = HKDF(algorithm=hashes.SHA256(), length=32, salt=secret, info=b"m365-mcp link value").derive(link_id.encode())
    return AESGCM(key)


class LinkStore:
    """Links in the shared store, under the names and the value key of the sealing key that issued them.

    `keys` are the store sub-keys of every line of the key file, the current one first. A link is put under
    the current key. Read back, it is looked for under each key in turn, so while pods roll after a key
    rotation (the previous key kept as the file's second line) a link one generation issued is found by the
    other. Every later mark on that link (busy, spent) is written under the generation that holds it."""

    def __init__(
        self,
        store: StateStore | None = None,
        max_entries: int = 5_000,
        per_owner: int = 300,
        *,
        key: bytes = b"\0" * 32,
        previous: tuple[bytes, ...] = (),
    ) -> None:
        if any(len(k) < 32 for k in (key, *previous)):
            raise ValueError("the link store's keys must be at least 32 bytes")
        base: StateStore = store if store is not None else MemoryStore()
        self.state = base.view(Namer(key))
        self._gens: list[tuple[StateStore, bytes]] = [(self.state, key)] + [(base.view(Namer(k)), k) for k in previous]
        self._max = max_entries
        self._per_owner = per_owner  # one person looping a tool must not fill the store for everyone

    async def put(self, value: str, *, expires: float, owner: str = "", tag: str = "") -> str:
        key = secrets.token_urlsafe(16)  # 128 random bits, 22 characters
        nonce = os.urandom(12)
        sealed = nonce + _box(self._gens[0][1], key).encrypt(nonce, value.encode(), b"m365-mcp link")
        if not await self.state.link_put(
            key, sealed, expires=expires, owner=owner, owner_cap=self._per_owner, total_cap=self._max, tag=tag
        ):
            raise LinksFull()
        return key

    async def _find(self, key: str) -> tuple[StateStore, bytes, bytes | None]:
        """The generation holding this link, and its sealed value; the current one when none does."""
        for state, secret in self._gens:
            sealed = await state.link_get(key)
            if sealed is not None:
                return state, secret, sealed
        return self._gens[0][0], self._gens[0][1], None

    async def get(self, key: str) -> str | None:
        if not key or len(key) > 64:
            return None
        _state, secret, sealed = await self._find(key)
        if sealed is None or len(sealed) < 13:
            return None
        try:
            return _box(secret, key).decrypt(sealed[:12], sealed[12:], b"m365-mcp link").decode()
        except (InvalidTag, UnicodeDecodeError):
            return None

    async def _holder(self, key: str, stored: bool) -> StateStore | None:
        """The generation whose marks govern this link: the one holding its value. None when a stored link's
        value is in no generation any more (delivered, revoked or expired): nothing may claim it then, or
        an upload verified under one generation could be claimed again under another."""
        if len(self._gens) == 1 or not stored:
            return self.state
        state, _secret, sealed = await self._find(key)
        return state if sealed is not None else None

    async def claim(self, key: str, *, stored: bool = True) -> bool:
        """Mark a link as receiving an upload. False when another upload holds it or it was used.
        `stored`: the key is a short id whose value this store holds (False for a full signed value)."""
        holder = await self._holder(key, stored)
        return holder is not None and await holder.link_claim(key, hold=UPLOAD_HOLD)

    async def release(self, key: str, *, stored: bool = True) -> None:
        """The upload did not deliver: the link may be used again until it expires."""
        holder = await self._holder(key, stored)
        if holder is not None:
            await holder.link_release(key)

    async def consume(self, key: str, until: float, *, stored: bool = True) -> None:
        """The upload delivered: the link is spent, and a short id's stored value goes."""
        await (await self._holder(key, stored) or self.state).link_consume(key, until=until)

    async def spent(self, key: str) -> bool:
        for state, _secret in self._gens:
            if await state.link_spent(key):
                return True
        return False

    async def busy_under(self, tag: str) -> bool:
        """Whether a link issued under `tag` is receiving an upload right now."""
        for state, _secret in self._gens:
            if await state.link_busy_under(tag):
                return True
        return False

    async def revoke(self, tag: str) -> int:
        """Spend every link issued under `tag` (a draft that was sent): busy ones too, so an upload in
        progress cannot be followed by another. Returns how many were revoked."""
        return sum([await state.link_revoke(tag) for state, _secret in self._gens])
