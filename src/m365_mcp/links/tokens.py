"""Signed, short-lived links that move bytes without a Microsoft token leaving the server.

The payload carries the operation, the kind of object, the target, the person, the expiry,
and the person's Graph token encrypted under a separate key. The signature covers all of it.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import time
from dataclasses import dataclass
from typing import Any

from cryptography.hazmat.primitives.ciphers.aead import AESGCM

from m365_mcp.config import LINK_TTL_CAP
from m365_mcp.links.store import LinkStore

KINDS = frozenset({
    "attachment", "drive", "recording", "transcript", "note_resource",
    "draft_body", "draft_attachment", "drive_upload", "note_body",
})
DOWNLOAD_KINDS = frozenset({"attachment", "drive", "recording", "transcript", "note_resource"})
UPLOAD_KINDS = KINDS - DOWNLOAD_KINDS


def draft_tag(person: str, kind: str, target: dict[str, Any]) -> str:
    """The links that write into one person's draft share a tag, so sending the draft can revoke them."""
    if kind not in ("draft_body", "draft_attachment") or not target.get("d"):
        return ""
    return f"draft:{person}:{target['d']}"


class LinkRefused(Exception):
    """Any reason a link is not honoured. Deliberately carries no detail."""


def _b64(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode()


def _unb64(s: str) -> bytes:
    return base64.urlsafe_b64decode(s + "=" * (-len(s) % 4))


@dataclass(frozen=True)
class Link:
    op: str
    kind: str
    target: dict[str, Any]
    person: str
    expires: int
    token: str
    max_bytes: int | None = None
    content_type: str | None = None


class LinkSigner:
    def __init__(self, sign_keys: tuple[bytes, ...], enc_keys: tuple[bytes, ...],
                 store: LinkStore | None = None) -> None:
        if not sign_keys or not enc_keys or any(len(k) < 32 for k in sign_keys + enc_keys):
            # Fail closed: a signer without keys is not a signer.
            raise ValueError("link keys are required")
        self._sign = sign_keys
        self._enc = enc_keys
        # With a store, issue() keeps the signed value here and hands out a short id;
        # verify() takes either. Without one, the signed value is the link, as before.
        self._store = store
        # Which upload links are busy or spent; the store when there is one, else a store of its own.
        self._uses = store if store is not None else LinkStore()

    def _mac(self, key: bytes, payload: bytes) -> str:
        return _b64(hmac.new(key, payload, hashlib.sha256).digest()[:16])

    def issue(self, *, op: str, kind: str, target: dict[str, Any], person: str, token: str,
              token_expires: int, ttl: int, max_bytes: int | None = None,
              content_type: str | None = None, now: float | None = None) -> str:
        if op not in ("d", "u") or kind not in KINDS:
            raise ValueError("bad link")
        if (op == "d") != (kind in DOWNLOAD_KINDS):
            raise ValueError("operation does not match kind")
        now = time.time() if now is None else now
        exp = int(min(now + min(ttl, LINK_TTL_CAP), token_expires))
        if exp <= now:
            raise ValueError("token already expired")
        nonce = os.urandom(12)
        sealed = nonce + AESGCM(self._enc[0][:32]).encrypt(nonce, token.encode(), person.encode())
        body: dict[str, Any] = {"o": op, "k": kind, "t": target, "p": person, "e": exp, "x": _b64(sealed)}
        if op == "u":
            body["m"] = {"max": max_bytes, "ct": content_type}
        payload = json.dumps(body, separators=(",", ":"), sort_keys=True).encode()
        value = f"{_b64(payload)}.{self._mac(self._sign[0], payload)}"
        if self._store is None:
            return value
        return self._store.put(value, expires=exp, owner=person, tag=draft_tag(person, kind, target))

    def revoke_draft(self, person: str, draft_id: str) -> int:
        """A draft that is being sent takes no more uploads: its body and attachment links are spent."""
        return self._uses.revoke(draft_tag(person, "draft_body", {"d": draft_id}))

    def draft_busy(self, person: str, draft_id: str) -> bool:
        """An upload into this draft (its body or a file) is running now."""
        return self._uses.busy_under(draft_tag(person, "draft_body", {"d": draft_id}))

    @staticmethod
    def _use_key(value: str) -> str:
        # A short id is its own key; a full signed value is long, so its digest stands for it.
        return value if "." not in value else hashlib.sha256(value.encode()).hexdigest()

    def claim(self, value: str) -> bool:
        """Start an upload on this link; False while another one runs on it or after one delivered."""
        return self._uses.claim(self._use_key(value))

    def release(self, value: str) -> None:
        self._uses.release(self._use_key(value))

    def consume(self, value: str, until: float) -> None:
        self._uses.consume(self._use_key(value), until)

    def verify(self, value: str, *, op: str, now: float | None = None) -> Link:
        if op == "u" and self._uses.spent(self._use_key(value)):
            raise LinkRefused()
        if "." not in value:
            # A short id: the signed value is the one this process stored, checked below as before.
            stored = self._store.get(value, now) if self._store is not None else None
            if stored is None:
                raise LinkRefused()
            value = stored
        try:
            p64, sig = value.split(".", 1)
            payload = _unb64(p64)
        except (ValueError, TypeError):
            raise LinkRefused() from None
        if not sig.isascii() or not any(hmac.compare_digest(self._mac(k, payload), sig) for k in self._sign):
            raise LinkRefused()
        try:
            body = json.loads(payload)
            if body["o"] != op or body["k"] not in KINDS:
                raise LinkRefused()
            if (op == "d") != (body["k"] in DOWNLOAD_KINDS):
                raise LinkRefused()
            now = time.time() if now is None else now
            if not isinstance(body["e"], int) or body["e"] <= now:
                raise LinkRefused()
            sealed = _unb64(body["x"])
            token = None
            for k in self._enc:
                try:
                    token = AESGCM(k[:32]).decrypt(sealed[:12], sealed[12:], body["p"].encode()).decode()
                    break
                except Exception:  # noqa: BLE001 - try the next key; any failure means "not this key"
                    continue
            if token is None:
                raise LinkRefused()
            m = body.get("m") or {}
            return Link(op=op, kind=body["k"], target=body["t"], person=body["p"], expires=body["e"],
                        token=token, max_bytes=m.get("max"), content_type=m.get("ct"))
        except LinkRefused:
            raise
        except (KeyError, TypeError, ValueError, AttributeError):
            raise LinkRefused() from None
