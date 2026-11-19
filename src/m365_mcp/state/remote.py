"""The state store in Redis or Valkey: standalone (`M365_STATE_URL`) or behind Sentinel
(`M365_STATE_SENTINELS`), through redis-py's asyncio client.

Every operation that reads and writes together is one Lua script, so two replicas can never both take the
last token of a bucket, both claim one upload link, or both store a link past a person's cap. Time inside
a script is the server's (`TIME`), so replicas whose clocks drift still refill one bucket alike.

Timeouts are short (one second to connect, one to answer, three for a whole operation): a store that does
not answer makes a send fail closed within seconds rather than hang the call. On Sentinel, a failover is
found by redis-py's Sentinel client, which asks the sentinels for the master again after a failed
connection; the first operations during a failover fail closed, the next ones reach the new master.

Standalone and Sentinel only: the scripts touch keys they derive (a link's owner set), which Redis Cluster
would refuse across slots.
"""

from __future__ import annotations

import asyncio
import math
from collections.abc import Awaitable, Callable
from typing import Any, TypeVar

from redis.asyncio import Redis
from redis.asyncio.retry import Retry
from redis.asyncio.sentinel import Sentinel
from redis.backoff import NoBackoff
from redis.exceptions import RedisError

from m365_mcp.state.store import Namer, measured

T = TypeVar("T")

SOCKET_TIMEOUT = 1.0
CONNECT_TIMEOUT = 1.0
OPERATION_TIMEOUT = 3.0
ERRORS: tuple[type[BaseException], ...] = (RedisError, OSError, asyncio.TimeoutError)

# A script with the shebang is checked against maxmemory before it runs: with noeviction, a full store
# refuses the whole script, and the operation fails closed. Without it Redis lets a script write past
# the limit once it has written (measured by exchange-mcp on Redis 7.2 and Valkey 8.1). Redis or Valkey 7+.
_SHEBANG = "#!lua\n"

_NOW = "local t = redis.call('TIME') local now = tonumber(t[1]) + tonumber(t[2]) / 1000000\n"

# KEYS: the person's bucket under each key of the key file (one, or two during a rotation). ARGV: capacity,
# refill per second. A token is taken from every one or from none, so pods holding the key file in either
# order during a rotation draw from a bucket they share. Returns {taken, tokens after, the fewest}.
_TAKE = (
    _NOW
    + """
local cap = tonumber(ARGV[1]) local rate = tonumber(ARGV[2])
local tokens = {}
local fewest = cap
for i, k in ipairs(KEYS) do
  local b = redis.call('HMGET', k, 'tokens', 'at')
  local t = tonumber(b[1]) local at = tonumber(b[2])
  if t == nil or at == nil then t = cap at = now end
  t = math.min(cap, t + math.max(0, now - at) * rate)
  tokens[i] = t
  if t < fewest then fewest = t end
end
local taken = 0
if fewest >= 1 then taken = 1 end
for i, k in ipairs(KEYS) do
  local t = tokens[i] - taken
  redis.call('HSET', k, 'tokens', tostring(t), 'at', tostring(now))
  redis.call('PEXPIRE', k, math.ceil((cap - t) / rate * 1000) + 1000)
end
return {taken, tostring(fewest - taken)}
"""
)

# Keeps a sorted set alive as long as its longest-lived member.
_KEEP = """
local function keep(k)
  local top = redis.call('ZRANGE', k, -1, -1, 'WITHSCORES')
  if top[2] then redis.call('PEXPIRE', k, math.ceil((tonumber(top[2]) - now) * 1000) + 1000) end
end
"""

# KEYS: link, owner set, all set, tag set (or an unused name). ARGV: member, value, expires, owner cap,
# total cap, has tag. Returns 1 when stored.
_PUT = (
    _NOW
    + _KEEP
    + """
local exp = tonumber(ARGV[3])
if exp <= now then return 0 end
redis.call('ZREMRANGEBYSCORE', KEYS[2], '-inf', now)
redis.call('ZREMRANGEBYSCORE', KEYS[3], '-inf', now)
if redis.call('ZCARD', KEYS[2]) >= tonumber(ARGV[4]) then return 0 end
if redis.call('ZCARD', KEYS[3]) >= tonumber(ARGV[5]) then return 0 end
if redis.call('EXISTS', KEYS[1]) == 1 then return 0 end
local tag = ''
if ARGV[6] == '1' then tag = KEYS[4] end
redis.call('HSET', KEYS[1], 'v', ARGV[2], 'o', KEYS[2], 't', tag)
redis.call('PEXPIRE', KEYS[1], math.ceil((exp - now) * 1000))
redis.call('ZADD', KEYS[2], exp, ARGV[1]) keep(KEYS[2])
redis.call('ZADD', KEYS[3], exp, ARGV[1]) keep(KEYS[3])
if tag ~= '' then
  redis.call('ZREMRANGEBYSCORE', tag, '-inf', now)
  redis.call('ZADD', tag, exp, ARGV[1]) keep(tag)
end
return 1
"""
)

# KEYS: busy, spent, link. ARGV: hold in ms, member. A claim on a link issued under a tag (a draft) is
# also entered in the tag's busy set for as long as it holds: an upload admitted just before its link
# expired still counts as running into that draft while its body arrives.
_CLAIM = (
    _NOW
    + _KEEP
    + """
if redis.call('EXISTS', KEYS[2]) == 1 then return 0 end
if not redis.call('SET', KEYS[1], '1', 'NX', 'PX', ARGV[1]) then return 0 end
local tag = redis.call('HGET', KEYS[3], 't')
if tag and tag ~= '' then
  local busy = tag .. ':busy'
  redis.call('ZREMRANGEBYSCORE', busy, '-inf', now)
  redis.call('ZADD', busy, now + tonumber(ARGV[1]) / 1000, ARGV[2]) keep(busy)
end
return 1
"""
)

# Forgets a link's value and its place in the owner and tag sets. Needs `now`, the all set in `all`.
_FORGET = """
local function forget(link, member)
  local meta = redis.call('HMGET', link, 'o', 't')
  if meta[1] then redis.call('ZREM', meta[1], member) end
  if meta[2] and meta[2] ~= '' then redis.call('ZREM', meta[2], member) end
  redis.call('ZREM', all, member)
  redis.call('DEL', link)
end
"""

# KEYS: busy, spent, link, all set. ARGV: until (epoch seconds), member.
_CONSUME = (
    _NOW
    + "local all = KEYS[4]\n"
    + _FORGET
    + """
redis.call('DEL', KEYS[1])
local ms = math.floor((tonumber(ARGV[1]) - now) * 1000)
if ms > 0 then redis.call('SET', KEYS[2], '1', 'PX', ms) end
forget(KEYS[3], ARGV[2])
return 1
"""
)

# KEYS: tag set. ARGV: busy-key prefix. Busy: a claim entered under the tag, still held.
_BUSY_UNDER = (
    _NOW
    + """
local busy = KEYS[1] .. ':busy'
redis.call('ZREMRANGEBYSCORE', busy, '-inf', now)
for _, m in ipairs(redis.call('ZRANGE', busy, 0, -1)) do
  if redis.call('EXISTS', ARGV[1] .. m) == 1 then return 1 end
end
return 0
"""
)

# KEYS: tag set, all set. ARGV: link prefix, busy prefix, spent prefix.
_REVOKE = (
    _NOW
    + "local all = KEYS[2]\n"
    + _FORGET
    + """
redis.call('ZREMRANGEBYSCORE', KEYS[1], '-inf', now)
local members = redis.call('ZRANGE', KEYS[1], 0, -1, 'WITHSCORES')
local n = 0
for i = 1, #members, 2 do
  local m = members[i]
  local ms = math.ceil((tonumber(members[i + 1]) - now) * 1000)
  if ms > 0 then redis.call('SET', ARGV[3] .. m, '1', 'PX', ms) end
  forget(ARGV[1] .. m, m)
  redis.call('DEL', ARGV[2] .. m)
  n = n + 1
end
-- An upload still running past its link's expiry is spent too, so it cannot be followed by another.
local busy = KEYS[1] .. ':busy'
local running = redis.call('ZRANGE', busy, 0, -1, 'WITHSCORES')
for i = 1, #running, 2 do
  local m = running[i]
  local ms = math.ceil((tonumber(running[i + 1]) - now) * 1000)
  if ms > 0 and redis.call('EXISTS', ARGV[2] .. m) == 1 then
    redis.call('SET', ARGV[3] .. m, '1', 'PX', ms)
    redis.call('DEL', ARGV[2] .. m)
  end
end
redis.call('DEL', KEYS[1], busy)
return n
"""
)


def client_kwargs(
    username: str | None, password: str | None, tls: bool, db: int, ca_bundle: str = ""
) -> dict[str, Any]:
    extra = {"ssl_ca_certs": ca_bundle} if tls and ca_bundle else {}
    return {
        **extra,
        "username": username or None,
        "password": password,
        "ssl": tls,
        "db": db,
        "socket_timeout": SOCKET_TIMEOUT,
        "socket_connect_timeout": CONNECT_TIMEOUT,
        # No retry: a script whose reply was lost may have run, and running it again would take a second
        # token or find its own busy mark. The operation fails closed instead, and the next one connects
        # afresh (on Sentinel, to whichever master the sentinels name). A connection idle for 15 s is
        # checked with a PING before it is used, so a connection the server closed is not the one that fails.
        "retry": Retry(NoBackoff(), 0),
        "health_check_interval": 15,
    }


class RedisStore:
    kind = "redis"

    def __init__(self, client: Redis, namer: Namer, prefix: str, sentinel: Sentinel | None = None) -> None:
        self._r = client
        self._name = namer
        self._p = prefix
        self._sentinel = sentinel
        self._owner = True
        self._take = client.register_script(_SHEBANG + _TAKE)
        self._put = client.register_script(_SHEBANG + _PUT)
        self._claim = client.register_script(_SHEBANG + _CLAIM)
        self._consume = client.register_script(_SHEBANG + _CONSUME)
        self._busy = client.register_script(_SHEBANG + _BUSY_UNDER)
        self._revoke = client.register_script(_SHEBANG + _REVOKE)

    @classmethod
    def standalone(
        cls,
        url: str,
        namer: Namer,
        prefix: str,
        *,
        username: str | None = None,
        password: str | None = None,
        db: int | None = None,
        ca_bundle: str = "",
    ) -> RedisStore:
        kw = client_kwargs(username, password, url.startswith("rediss://"), 0, ca_bundle)
        kw.pop("ssl")  # the scheme says it: rediss:// is TLS
        kw.pop("db")
        if db is not None:
            kw["db"] = db
        return cls(Redis.from_url(url, **kw), namer, prefix)

    @classmethod
    def sentinels(
        cls,
        hosts: list[tuple[str, int]],
        master: str,
        namer: Namer,
        prefix: str,
        *,
        username: str | None = None,
        password: str | None = None,
        sentinel_password: str | None = None,
        tls: bool = False,
        db: int = 0,
        ca_bundle: str = "",
    ) -> RedisStore:
        sentinel_kw = {
            "password": sentinel_password,
            "ssl": tls,
            "socket_timeout": SOCKET_TIMEOUT,
            "socket_connect_timeout": CONNECT_TIMEOUT,
            "retry": Retry(NoBackoff(), 0),
            **({"ssl_ca_certs": ca_bundle} if tls and ca_bundle else {}),
        }
        s = Sentinel(hosts, sentinel_kwargs=sentinel_kw, **client_kwargs(username, password, tls, db, ca_bundle))
        return cls(s.master_for(master), namer, prefix, sentinel=s)

    def view(self, namer: Namer) -> RedisStore:
        """The same store and connection, naming things under another key (a previous sealing key, during a
        rotation). Only the store this one was made from closes the connection."""
        other = RedisStore.__new__(RedisStore)
        other.__dict__.update(self.__dict__)
        other._name = namer
        other._sentinel = None
        other._owner = False
        return other

    def _k(self, kind: str, raw: str) -> str:
        return f"{self._p}{kind}:{self._name(kind, raw)}"

    async def _op(self, op: str, call: Callable[[], Awaitable[T]]) -> T:
        async def bounded() -> T:
            return await asyncio.wait_for(call(), OPERATION_TIMEOUT)

        return await measured(op, bounded, ERRORS)

    async def ping(self) -> None:
        await self._op("ping", self._r.ping)

    async def take(
        self, person: str, *, capacity: float, refill_per_second: float, also: tuple[Namer, ...] = ()
    ) -> tuple[bool, float]:
        keys = [self._k("bucket", person)] + [f"{self._p}bucket:{n('bucket', person)}" for n in also]

        async def call() -> tuple[bool, float]:
            taken, tokens = await self._take(keys=keys, args=[capacity, refill_per_second])
            return bool(int(taken)), float(tokens)

        return await self._op("take", call)

    async def link_put(
        self, link_id: str, value: bytes, *, expires: float, owner: str, owner_cap: int, total_cap: int, tag: str
    ) -> bool:
        member = self._name("link", link_id)
        keys = [
            f"{self._p}link:{member}",
            self._k("owner", owner),
            f"{self._p}links",
            self._k("tag", tag) if tag else f"{self._p}tag:none",
        ]
        args = [member, value, repr(float(expires)), owner_cap, total_cap, "1" if tag else "0"]

        async def call() -> bool:
            return bool(int(await self._put(keys=keys, args=args)))

        return await self._op("link_put", call)

    async def link_get(self, link_id: str) -> bytes | None:
        key = f"{self._p}link:{self._name('link', link_id)}"
        return await self._op("link_get", lambda: self._r.hget(key, "v"))

    async def link_claim(self, link_id: str, *, hold: float) -> bool:
        member = self._name("link", link_id)

        async def call() -> bool:
            got = await self._claim(
                keys=[f"{self._p}busy:{member}", f"{self._p}spent:{member}", f"{self._p}link:{member}"],
                args=[max(1, math.ceil(hold * 1000)), member],
            )
            return bool(int(got))

        return await self._op("link_claim", call)

    async def link_release(self, link_id: str) -> None:
        key = f"{self._p}busy:{self._name('link', link_id)}"
        await self._op("link_release", lambda: self._r.delete(key))

    async def link_consume(self, link_id: str, *, until: float) -> None:
        m = self._name("link", link_id)
        keys = [f"{self._p}busy:{m}", f"{self._p}spent:{m}", f"{self._p}link:{m}", f"{self._p}links"]
        await self._op("link_consume", lambda: self._consume(keys=keys, args=[repr(float(until)), m]))

    async def link_spent(self, link_id: str) -> bool:
        key = f"{self._p}spent:{self._name('link', link_id)}"

        async def call() -> bool:
            return bool(await self._r.exists(key))

        return await self._op("link_spent", call)

    async def link_busy_under(self, tag: str) -> bool:
        if not tag:
            return False

        async def call() -> bool:
            return bool(int(await self._busy(keys=[self._k("tag", tag)], args=[f"{self._p}busy:"])))

        return await self._op("link_busy_under", call)

    async def link_revoke(self, tag: str) -> int:
        if not tag:
            return 0

        async def call() -> int:
            return int(
                await self._revoke(
                    keys=[self._k("tag", tag), f"{self._p}links"],
                    args=[f"{self._p}link:", f"{self._p}busy:", f"{self._p}spent:"],
                )
            )

        return await self._op("link_revoke", call)

    async def aclose(self) -> None:
        if not self._owner:
            return
        await self._r.aclose()
        if self._sentinel is not None:
            for s in self._sentinel.sentinels:
                await s.aclose()
