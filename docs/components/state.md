# Component: state

What replicas must agree on, where it lives, and what happens when it is gone.

## What is shared

| What | Operation | Lives |
|---|---|---|
| Each person's send allowance | token bucket, ten at once, one per six seconds; one Lua script | until full again (about a minute) |
| A short link's id → its sealed value | stored with its expiry; at most 300 per person and 5,000 in all, checked in the same script | the link's life: 300 s down, 180 s up |
| An upload running on a link | busy mark, `SET NX` with a deadline (1,200 s), so a dead replica frees it; entered under the link's draft for as long as it holds, past the link's own expiry | the upload |
| A link that delivered | spent mark | the link's life |
| The links that write into one draft | a set, revoked together when the draft is sent | the links' life |

Authorisation is not here: the gateway forwards each person's Graph token, and this server keeps no
session. Graph pacing and concurrency stay per process ([graph client](graph-client.md)).

## Two implementations, one contract

`state/store.py` is the interface; `MemoryStore` (one process) and `RedisStore` (redis-py's asyncio
client) implement it. `tests/state/` is one contract suite that runs against memory always, and in CI
(`scripts/state-contract.sh`) against real Redis 7.4 and Valkey 8.1 standalone, Valkey behind three
Sentinels with passwords, and Redis behind one Sentinel without any; never a fake. Each Lua script's
guard is removed in turn there (`mutation_check.py --stores`) and the suite must fail.

## Configuration

| Variable | Meaning |
|---|---|
| `M365_STATE_URL` | `redis://host:6379/0` or `rediss://…`: one server |
| `M365_STATE_SENTINELS` | `host:26379[,host:26379…]`: Sentinel; one host (a Service in front of them) is enough |
| `M365_STATE_SENTINEL_MASTER` | default `mymaster` |
| `M365_STATE_DB` | Sentinel mode; with a URL its path decides |
| `M365_STATE_TLS` | Sentinel mode; with a URL its scheme decides |
| `M365_STATE_USERNAME` | an ACL user |
| `M365_STATE_PASSWORD_FILE`, `M365_STATE_SENTINEL_PASSWORD_FILE` | passwords, from files only: regular, 0400, owned by the server's user, at most 64 KiB |
| `M365_STATE_PREFIX` | default `microsoft-365-mcp:`; every key is under it |

A URL and Sentinels together, a user or password inside the URL, or `M365_STATE_PASSWORD` in the
environment stop the server at start. Neither set: memory, with a warning at start; the chart refuses
more than one replica that way.

## What the store can see

Nothing it could use. Every name is HMAC-SHA256 of the identifier (person hash, link id, draft) under a
key derived from the sealing key for this purpose alone, truncated to 128 bits. A link's value is
AES-GCM ciphertext under a key derived from that key and the link id, which only the link's holder
has. A store without a password (common inside a cluster) leaks no token, no link and no person.

## Requirements on the store

- `maxmemory-policy noeviction`. Every key here has a TTL and all of it is small, but an evicting
  policy could drop a spent mark or a bucket and let a second upload or a send through; with
  `noeviction` a full store refuses the write, and the server fails closed.
- Redis Cluster is not supported: the scripts touch keys they derive.

## When it is gone

| Operation | Store down |
|---|---|
| Tools that only read | keep working |
| Anything that draws from the allowance (send, invite, share, …) | refused: "try again in a few seconds", nothing done |
| A new file link | refused the same way, no link made |
| `send_draft` checking that no upload into the draft runs | refused, the draft kept |
| A link route | 503 with `Retry-After: 5`; never 403, the link may be good |
| `/readyz` | 200, `"store": "down"`: readiness never waits for the store, or one outage would empty the Service |
| Metrics | `mcp_state_store_up 0`, `mcp_state_store_errors_total{op}` |

Timeouts: one second to connect, one to answer, three for a whole operation, one immediate retry. On
Sentinel, redis-py asks the sentinels for the master again after a failed connection, so the first
operations of a failover fail closed and the next reach the new master; the chart e2e measures the
whole outage, from killing the master to a send that works, under 30 seconds.

- No persistence is needed: everything here lives minutes. `save ""` and `appendonly no` keep the
  sealed link values off disk.
- Replication to a Sentinel replica is asynchronous. A master that dies right after a write can lose
  it in the failover: a few sends drawn in the last moment are given back, and an upload link spent in
  that moment could take one more file before it expires (at most 180 s). Accepted: the alternative,
  waiting for the replica on every write, stalls sends whenever the replica is away.

## Key rotation

The key file holds one key per line; the first issues, all open. Links are looked for under every
line's store key, and their marks are written under the line that holds them. A rotation is three
deploys (`scripts/new-sealing-key.sh`), so every pod serving during a rollout opens what any other
issues:

| Phase | Command | File | During the rollout |
|---|---|---|---|
| 1 | `new-sealing-key.sh <ns> <secret>` | old, new | pods with and without the new key all issue under the old one |
| 2 | `--promote` | new, old | new pods issue under the new key; pods of phase 1 already open it |
| 3 | `--drop-previous`, after the old key's links expired (minutes) | new | nothing issues under the old key any more |

A send draws from the person's bucket under every key the pod holds, all or none, so pods holding the
file in either order share one allowance through every phase. After a leak, run the three at once:
links issued in between are refused, and the old key is gone.
