"""Opaque cursors. A cursor wraps Graph's next link, bound to one tool and one person and
signed, so it cannot be replayed elsewhere or edited to point the server at another URL."""

from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os

from m365_mcp.core.errors import invalid

_FALLBACK_KEY = os.urandom(32)  # only for configurations without link keys (account area alone)


def _b64(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode()


def _unb64(s: str) -> bytes:
    return base64.urlsafe_b64decode(s + "=" * (-len(s) % 4))


class Cursors:
    def __init__(self, key: bytes | None) -> None:
        if key is not None and len(key) < 32:
            raise ValueError("a cursor key must be at least 32 bytes")
        self._key = key if key is not None else _FALLBACK_KEY

    def _mac(self, payload: bytes) -> str:
        return _b64(hmac.new(self._key, b"cursor:" + payload, hashlib.sha256).digest()[:12])

    def make(self, tool: str, person: str, state: dict) -> str:
        payload = json.dumps({"t": tool, "p": person, "s": state}, separators=(",", ":")).encode()
        return f"{_b64(payload)}.{self._mac(payload)}"

    def read(self, tool: str, person: str, cursor: str) -> dict:
        try:
            p64, sig = cursor.split(".", 1)
            payload = _unb64(p64)
            if not hmac.compare_digest(self._mac(payload), sig):
                raise ValueError
            body = json.loads(payload)
            if body["t"] != tool or body["p"] != person:
                raise ValueError
            return body["s"]
        except (ValueError, KeyError, TypeError):
            raise invalid("This cursor does not belong to this list. Start the list again without a cursor.") from None
