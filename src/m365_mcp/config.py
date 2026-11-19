"""Configuration: settings from the environment, secrets from files. Nothing here names an installation.

The sealing key is a file (`M365_KEY_FILE`) the chart's init container stages, mode 0400, into a memory
volume. Keys in the environment (`M365_LINK_SECRET`, `M365_LINK_SIGNING_KEYS`,
`M365_LINK_ENCRYPTION_KEYS`) still work in 0.2.x with a warning at every start, and are refused from
0.3.0: an environment is readable by every process of the same user and by anyone who can read the pod
spec or its Secret reference.
"""

from __future__ import annotations

import base64
import os
from dataclasses import dataclass, field

from m365_mcp.keys.files import SecretFileError, read_key_file, read_password_file
from m365_mcp.registry import AREAS

# Switches inside areas. Each one needs a permission Microsoft lets only an administrator
# grant; they are on by default because the administrator consents once at installation.
SWITCHES: dict[str, bool] = {
    "teams_channels": True,  # ChannelMessage.Read.All, ChannelMessage.ReadWrite
    "transcripts": True,  # OnlineMeetingTranscript.Read.All
    "rooms": True,  # Place.Read.All
    "chat_admin": False,  # creating group chats and changing their members: people see it at once
}

LINK_TTL_DOWNLOAD = 300
LINK_TTL_UPLOAD = 180
LINK_TTL_CAP = 900


ENV_KEYS = ("M365_LINK_SECRET", "M365_LINK_SIGNING_KEYS", "M365_LINK_ENCRYPTION_KEYS")
ENV_KEYS_DEPRECATED = (
    "the link keys are read from the environment ({names}); this works in 0.2.x and is refused from 0.3.0. "
    "Move them to a key file (M365_KEY_FILE), staged by the chart's init container"
)
DEFAULT_PREFIX = "microsoft-365-mcp:"
DEFAULT_METRICS_PORT = 9090


class ConfigError(Exception):
    pass


def _hkdf(material: bytes, salt: bytes, purpose: bytes) -> bytes:
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.kdf.hkdf import HKDF

    return HKDF(algorithm=hashes.SHA256(), length=32, salt=salt, info=purpose).derive(material)


def _derive(secret: str, purpose: bytes) -> bytes:
    """One installation secret, independent keys: HKDF with a purpose label for each."""
    if len(secret) < 32:
        raise ConfigError("M365_LINK_SECRET must be at least 32 characters")
    return _hkdf(secret.encode(), b"m365-mcp-links", purpose)


def _file_keys(path: str) -> tuple[tuple[bytes, ...], tuple[bytes, ...], tuple[bytes, ...]]:
    """Each line of the key file gives a signing, an encryption and a store key. The first line signs,
    seals and names; every line opens and finds."""
    try:
        lines = read_key_file(path, "M365_KEY_FILE")
    except SecretFileError as e:
        raise ConfigError(str(e)) from None
    sign = tuple(_hkdf(k, b"m365-mcp-links", b"sign") for _, k in lines)
    enc = tuple(_hkdf(k, b"m365-mcp-links", b"encrypt") for _, k in lines)
    return sign, enc, tuple(_hkdf(k, b"m365-mcp-state", b"state") for _, k in lines)


@dataclass(frozen=True)
class StateConfig:
    """Where replicas keep what they share. Passwords come from files, never the environment."""

    url: str = ""
    sentinels: tuple[tuple[str, int], ...] = ()
    master: str = "mymaster"
    db: int | None = None
    tls: bool = False
    username: str = ""
    password: str | None = field(default=None, repr=False)
    sentinel_password: str | None = field(default=None, repr=False)
    prefix: str = DEFAULT_PREFIX

    @property
    def mode(self) -> str:
        return "sentinel" if self.sentinels else "standalone"


def _bool(name: str) -> bool:
    raw = os.environ.get(name, "").strip().lower()
    if raw in ("", "0", "false", "no"):
        return False
    if raw in ("1", "true", "yes"):
        return True
    raise ConfigError(f"{name}: true or false")


def _int(name: str, default: int) -> int:
    raw = os.environ.get(name, "").strip()
    if not raw:
        return default
    try:
        return int(raw)
    except ValueError:
        raise ConfigError(f"{name}: a whole number") from None


def _password(name: str) -> str | None:
    path = os.environ.get(name, "").strip()
    if not path:
        return None
    try:
        return read_password_file(path, name)
    except SecretFileError as e:
        raise ConfigError(str(e)) from None


def _state() -> StateConfig | None:
    """None: the store is this process's memory (one replica)."""
    from urllib.parse import urlsplit

    url = os.environ.get("M365_STATE_URL", "").strip()
    raw_sentinels = os.environ.get("M365_STATE_SENTINELS", "").strip()
    for secret in ("M365_STATE_PASSWORD", "M365_STATE_SENTINEL_PASSWORD"):
        if os.environ.get(secret):
            raise ConfigError(
                f"{secret} is never read from the environment: put the password in a file and name it in {secret}_FILE"
            )
    if url and raw_sentinels:
        raise ConfigError("M365_STATE_URL and M365_STATE_SENTINELS are two ways to reach one store: set one")
    if not url and not raw_sentinels:
        return None
    db = _int("M365_STATE_DB", -1)
    db = None if db == -1 and not os.environ.get("M365_STATE_DB", "").strip() else db
    if db is not None and not 0 <= db <= 15:
        raise ConfigError("M365_STATE_DB: 0 to 15")
    common = {
        "username": os.environ.get("M365_STATE_USERNAME", "").strip(),
        "password": _password("M365_STATE_PASSWORD_FILE"),
        "prefix": os.environ.get("M365_STATE_PREFIX", DEFAULT_PREFIX),
        "db": db,
    }
    if not common["prefix"]:
        raise ConfigError("M365_STATE_PREFIX: empty; every key needs this server's namespace")
    if url:
        parts = urlsplit(url)
        if parts.scheme not in ("redis", "rediss") or not parts.hostname:
            raise ConfigError("M365_STATE_URL: redis://host:6379/0 or rediss://host:6380/0")
        if parts.query or parts.fragment:
            # redis-py reads options from the query, and they win over ours: `?password=` would bypass the
            # file, `?ssl_cert_reqs=none` would turn certificate checks off.
            raise ConfigError("M365_STATE_URL: no query or fragment; options are settings of their own")
        if parts.password or parts.username:
            raise ConfigError(
                "M365_STATE_URL carries a user or a password; the password goes in a file named "
                "by M365_STATE_PASSWORD_FILE, the user in M365_STATE_USERNAME"
            )
        if _bool("M365_STATE_TLS") and parts.scheme != "rediss":
            raise ConfigError("M365_STATE_TLS=true with a redis:// URL; use rediss://")
        return StateConfig(url=url, tls=parts.scheme == "rediss", **common)
    hosts: list[tuple[str, int]] = []
    for part in raw_sentinels.split(","):
        host, sep, raw = part.strip().rpartition(":")
        try:
            port = int(raw)
        except ValueError:
            port = 0
        if not sep or not host or not 0 < port < 65536:
            raise ConfigError("M365_STATE_SENTINELS: host:port[,host:port…]")
        hosts.append((host.strip("[]"), port))
    master = os.environ.get("M365_STATE_SENTINEL_MASTER", "mymaster").strip()
    if not master:
        raise ConfigError("M365_STATE_SENTINEL_MASTER: empty")
    return StateConfig(
        sentinels=tuple(hosts),
        master=master,
        tls=_bool("M365_STATE_TLS"),
        sentinel_password=_password("M365_STATE_SENTINEL_PASSWORD_FILE"),
        **{**common, "db": db or 0},
    )


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
    # The key the store names things with (HMAC); derived from the sealing key, never used for anything else.
    state_key: bytes = field(default=b"", repr=False)
    # The store keys of the key file's later lines: links issued under them are still found (a rotation).
    previous_state_keys: tuple[bytes, ...] = field(default=(), repr=False)
    # None: this process's memory, which is right for one replica only.
    state: StateConfig | None = None
    # "file", "env" (deprecated) or "" (no key: the account area alone).
    key_source: str = ""
    metrics_port: int = DEFAULT_METRICS_PORT
    # Extra CAs (PEM) trusted for Microsoft Graph and the state store besides the system's.
    ca_bundle: str = ""
    # What the server says once at start, after logging is up: deprecations, a risky combination.
    warnings: tuple[str, ...] = ()

    @classmethod
    def from_env(cls) -> Settings:
        areas = _set("M365_AREAS", frozenset(a for a, on in AREAS.items() if on), set(AREAS))
        switches = _set("M365_SWITCHES", frozenset(s for s, on in SWITCHES.items() if on), set(SWITCHES))
        warnings: list[str] = []
        key_file = os.environ.get("M365_KEY_FILE", "").strip()
        in_env = [n for n in ENV_KEYS if os.environ.get(n, "").strip()]
        if key_file and in_env:
            raise ConfigError(
                f"M365_KEY_FILE is set, and so is {', '.join(in_env)}: the key comes from the file "
                "alone; remove the environment variables"
            )
        state_key = b""
        previous_state_keys: tuple[bytes, ...] = ()
        source = ""
        if key_file:
            sign, enc, state_keys = _file_keys(key_file)
            state_key, previous_state_keys = state_keys[0], state_keys[1:]
            source = "file"
        else:
            sign, enc = _keys("M365_LINK_SIGNING_KEYS"), _keys("M365_LINK_ENCRYPTION_KEYS")
            secret = os.environ.get("M365_LINK_SECRET", "").strip()
            if secret:
                sign, enc = (_derive(secret, b"sign"),) + sign, (_derive(secret, b"encrypt"),) + enc
                state_key = _derive(secret, b"state")
            elif sign:
                state_key = _hkdf(sign[0], b"m365-mcp-state", b"state")
            if in_env:
                source = "env"
                warnings.append(ENV_KEYS_DEPRECATED.format(names=", ".join(in_env)))
        metrics_port = _int("M365_METRICS_PORT", DEFAULT_METRICS_PORT)
        if not 0 < metrics_port < 65536:
            raise ConfigError("M365_METRICS_PORT: a port number")
        ca_bundle = os.environ.get("M365_CA_BUNDLE", "").strip()
        if ca_bundle:
            try:
                import ssl

                ssl.create_default_context().load_verify_locations(cafile=ca_bundle)
            except (OSError, ssl.SSLError):
                raise ConfigError("M365_CA_BUNDLE: not a readable file of PEM certificates") from None
        state = _state()
        if state is None:
            warnings.append(
                "the shared state is this process's memory: run one replica only, or set "
                "M365_STATE_URL or M365_STATE_SENTINELS"
            )
        elif not state_key:
            raise ConfigError(
                "a shared state store needs the sealing key (M365_KEY_FILE): every name in the store is derived from it"
            )
        s = cls(
            areas=areas,
            switches=switches,
            sign_keys=sign,
            enc_keys=enc,
            links_base_url=os.environ.get("M365_LINKS_BASE_URL", "").rstrip("/"),
            expected_client_ids=_lower_list("M365_EXPECTED_CLIENT_IDS"),
            allowed_recipient_domains=frozenset(d.lstrip("@.") for d in _lower_list("M365_ALLOWED_RECIPIENT_DOMAINS")),
            state_key=state_key,
            previous_state_keys=previous_state_keys,
            state=state,
            key_source=source,
            metrics_port=metrics_port,
            ca_bundle=ca_bundle,
            warnings=tuple(warnings),
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
        return bool(
            self.sign_keys and self.enc_keys and well_formed and (parts.scheme == "https" or _cluster_local(base))
        )

    def validate(self) -> None:
        # Mail, files, recordings, notes and Teams all move bytes through links. A server that
        # starts without keys would publish tools that fail on their first attachment.
        byte_areas = {"mail", "files", "recordings", "notes", "teams", "excel", "contacts"} & set(self.areas)
        if byte_areas and not self.links_ready:
            raise ConfigError(
                "a key file (M365_KEY_FILE) and an https or cluster-internal M365_LINKS_BASE_URL "
                f"are required while these areas are on: {', '.join(sorted(byte_areas))}"
            )
