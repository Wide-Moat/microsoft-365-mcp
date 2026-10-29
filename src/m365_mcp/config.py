"""Configuration, from the environment only. Nothing here names an installation."""

from __future__ import annotations

import base64
import os
from dataclasses import dataclass

from m365_mcp.registry import AREAS

# Switches inside areas. Each one needs a permission Microsoft lets only an administrator
# grant; they are on by default because the administrator consents once at installation.
SWITCHES: dict[str, bool] = {
    "teams_channels": True,   # ChannelMessage.Read.All, ChannelMessage.ReadWrite
    "transcripts": True,      # OnlineMeetingTranscript.Read.All
    "rooms": True,            # Place.Read.All
    "chat_admin": False,      # creating group chats and changing their members: people see it at once
}

LINK_TTL_DOWNLOAD = 300
LINK_TTL_UPLOAD = 180
LINK_TTL_CAP = 900


class ConfigError(Exception):
    pass


def _derive(secret: str, purpose: bytes) -> bytes:
    """One installation secret, two independent keys: HKDF with a purpose label for each."""
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.kdf.hkdf import HKDF

    if len(secret) < 32:
        raise ConfigError("M365_LINK_SECRET must be at least 32 characters")
    return HKDF(algorithm=hashes.SHA256(), length=32, salt=b"m365-mcp-links", info=purpose).derive(secret.encode())


def _cluster_local(url: str) -> bool:
    from urllib.parse import urlsplit

    host = urlsplit(url).hostname or ""
    return url.startswith("http://") and (host.endswith(".svc.cluster.local") or host.endswith(".svc"))


def _keys(name: str) -> tuple[bytes, ...]:
    raw = os.environ.get(name, "").strip()
    if not raw:
        return ()
    keys = []
    for part in raw.split(","):
        key = base64.urlsafe_b64decode(part.strip() + "=" * (-len(part.strip()) % 4))
        if len(key) < 32:
            raise ConfigError(f"{name}: every key must be at least 32 bytes")
        keys.append(key)
    return tuple(keys)


def _set(name: str, default: frozenset[str], known: set[str]) -> frozenset[str]:
    raw = os.environ.get(name)
    if raw is None:
        return default
    chosen = frozenset(p.strip() for p in raw.split(",") if p.strip())
    unknown = chosen - known
    if unknown:
        raise ConfigError(f"{name}: unknown {', '.join(sorted(unknown))}")
    return chosen


def _lower_list(name: str) -> frozenset[str]:
    return frozenset(p.strip().lower() for p in os.environ.get(name, "").split(",") if p.strip())


@dataclass(frozen=True)
class Settings:
    areas: frozenset[str]
    switches: frozenset[str]
    sign_keys: tuple[bytes, ...]
    enc_keys: tuple[bytes, ...]
    links_base_url: str
    # Client ids (`azp`/`appid`) a forwarded token must have been issued to. Empty: any client.
    expected_client_ids: frozenset[str] = frozenset()
    # Recipient domains anything that notifies other people may reach. Empty: no restriction.
    allowed_recipient_domains: frozenset[str] = frozenset()

    @classmethod
    def from_env(cls) -> Settings:
        areas = _set("M365_AREAS", frozenset(a for a, on in AREAS.items() if on), set(AREAS))
        switches = _set("M365_SWITCHES", frozenset(s for s, on in SWITCHES.items() if on), set(SWITCHES))
        sign, enc = _keys("M365_LINK_SIGNING_KEYS"), _keys("M365_LINK_ENCRYPTION_KEYS")
        secret = os.environ.get("M365_LINK_SECRET", "").strip()
        if secret:
            sign, enc = (_derive(secret, b"sign"),) + sign, (_derive(secret, b"encrypt"),) + enc
        s = cls(
            areas=areas,
            switches=switches,
            sign_keys=sign,
            enc_keys=enc,
            links_base_url=os.environ.get("M365_LINKS_BASE_URL", "").rstrip("/"),
            expected_client_ids=_lower_list("M365_EXPECTED_CLIENT_IDS"),
            allowed_recipient_domains=frozenset(d.lstrip("@.") for d in _lower_list("M365_ALLOWED_RECIPIENT_DOMAINS")),
        )
        s.validate()
        return s

    @property
    def links_ready(self) -> bool:
        # https in general; plain http only for a cluster-internal Service name, which a sandbox
        # reaches inside the cluster network.
        from urllib.parse import urlsplit

        base = self.links_base_url
        parts = urlsplit(base)
        well_formed = bool(parts.hostname) and not parts.query and not parts.fragment
        return bool(self.sign_keys and self.enc_keys and well_formed
                    and (parts.scheme == "https" or _cluster_local(base)))

    def validate(self) -> None:
        # Mail, files, recordings, notes and Teams all move bytes through links. A server that
        # starts without keys would publish tools that fail on their first attachment.
        byte_areas = {"mail", "files", "recordings", "notes", "teams", "excel", "contacts"} & set(self.areas)
        if byte_areas and not self.links_ready:
            raise ConfigError(
                "M365_LINK_SECRET (or M365_LINK_SIGNING_KEYS and M365_LINK_ENCRYPTION_KEYS) and an https or "
                "cluster-internal M365_LINKS_BASE_URL "
                f"are required while these areas are on: {', '.join(sorted(byte_areas))}"
            )
