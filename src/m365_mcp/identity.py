"""Who is calling: read from the Graph token the gateway forwards. Never stored, never logged."""

from __future__ import annotations

import base64
import hashlib
import json
import time
from contextvars import ContextVar
from dataclasses import dataclass

from m365_mcp.core.errors import M365Error

GRAPH_AUDIENCES = frozenset({"https://graph.microsoft.com", "00000003-0000-0000-c000-000000000000"})

# Set per request by the HTTP middleware; tests set them directly.
current_token: ContextVar[str | None] = ContextVar("current_token", default=None)
current_areas: ContextVar[frozenset[str] | None] = ContextVar("current_areas", default=None)


@dataclass(frozen=True)
class Person:
    tenant: str
    object_id: str
    scopes: frozenset[str]
    expires: int
    token: str  # the bearer token; repr hides it

    @property
    def key(self) -> str:
        return hashlib.sha256(f"{self.tenant}:{self.object_id}".encode()).hexdigest()[:16]

    def __repr__(self) -> str:
        return f"Person(key={self.key})"


def _claims(token: str) -> dict:
    parts = token.split(".")
    if len(parts) != 3:
        raise M365Error("reconnect")
    try:
        payload = base64.urlsafe_b64decode(parts[1] + "=" * (-len(parts[1]) % 4))
        claims = json.loads(payload)
    except (ValueError, UnicodeDecodeError):
        raise M365Error("reconnect") from None
    if not isinstance(claims, dict):
        raise M365Error("reconnect")
    return claims


def _client(claims: dict) -> str:
    # v2 tokens name the client in `azp`, v1 tokens (what Graph usually issues) in `appid`.
    value = claims.get("azp") or claims.get("appid")
    return value.strip().lower() if isinstance(value, str) else ""


def person_from_token(token: str | None, *, now: float | None = None,
                      clients: frozenset[str] = frozenset()) -> Person:
    """Shape, audience, expiry and, when the installation names its clients, the client the token was
    issued to. The signature is Graph's to check; a forged token gets a 401 there. A Graph token
    minted for some other application is still a valid Graph token, so only the client check keeps
    one that leaked from elsewhere from being replayed through this server."""
    if not token:
        raise M365Error("reconnect")
    c = _claims(token)
    now = time.time() if now is None else now
    exp = c.get("exp")
    aud = c.get("aud")
    if not isinstance(aud, str) or aud not in GRAPH_AUDIENCES or not isinstance(exp, int) or exp <= now:
        raise M365Error("reconnect")
    tid, oid = c.get("tid"), c.get("oid")
    if not (isinstance(tid, str) and tid and isinstance(oid, str) and oid):
        raise M365Error("reconnect")
    if clients and _client(c) not in clients:
        raise M365Error("reconnect")
    scopes = frozenset(str(c.get("scp", "")).split())
    return Person(tenant=tid, object_id=oid, scopes=scopes, expires=exp, token=token)


def current_person(clients: frozenset[str] = frozenset()) -> Person:
    return person_from_token(current_token.get(), clients=clients)
