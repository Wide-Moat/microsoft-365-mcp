# Changelog

## 0.2.0: many replicas, JSON logs, metrics on their own port, a Helm chart

Breaking changes are marked **breaking**; each says what to do.

### Runs as many replicas as the load needs

| What | Before | Now |
|---|---|---|
| File links (id → sealed value) | the issuing process's memory: one replica, a restart lost them | Redis or Valkey shared by every replica (`M365_STATE_URL`, or Sentinel: `M365_STATE_SENTINELS`); a link made on one pod works on any |
| Send allowance (10 at once, 1 per 6 s) | per process: N replicas, N allowances | one per person across replicas, one atomic script |
| Upload busy / spent marks, draft revocation | per process | shared, atomic: two replicas never accept one link's body twice |
| What the store holds | — | ciphertext under the sealing key and keyed hashes: no token, link, person or draft id readable |
| Rollout without a store | — | `Recreate` (one pod at a time, a few seconds without the server); with a store `RollingUpdate`, `maxUnavailable: 0` |
| Store down | — | reads keep working; sends, new links and `send_draft` fail closed ("try again in a few seconds"); link routes 503 with `Retry-After`, never 403 |
| No store configured | — | memory, with a warning at start; the chart refuses more than one replica that way |
| Graph pacing and concurrency | per process | still per process, documented ([graph client](docs/components/graph-client.md)) |

### The sealing key comes from a file

| What | Before | Now |
|---|---|---|
| Key source | `M365_LINK_SECRET` (or `M365_LINK_SIGNING_KEYS`/`M365_LINK_ENCRYPTION_KEYS`) in the environment | `M365_KEY_FILE`: one `<id>:<base64url 32 bytes>` per line, first issues, all open; regular file, 0400, the server's user, ≤ 64 KiB |
| Keys in the environment | the only way | **deprecated**: still read in 0.2.x with a warning at every start, **refused from 0.3.0**; set together with `M365_KEY_FILE` they stop the server |
| Process memory | readable by a shell of the same user | non-dumpable before any setting is read |
| Rotation | restart with a new secret, every open link refused | three deploys with `scripts/new-sealing-key.sh` (add, `--promote`, `--drop-previous`): every pod serving during a rollout opens every link |
| Store passwords | — | from files only (`M365_STATE_PASSWORD_FILE`, `M365_STATE_SENTINEL_PASSWORD_FILE`); a password in the environment or the URL stops the server |

### Logs: every line is JSON

| What | Before | Now |
|---|---|---|
| Format | `%(message)s`; only tool lines were JSON | one JSON object per line from every process: server, uvicorn, a refused start, an uncaught exception, the init container |
| Fields | — | `ts`, `level`, `logger`, `event`, `msg`, allowlisted fields, `exc` (frames only) |
| Level | INFO | `LOG_LEVEL` |

### Metrics: the common `mcp_*` names, on port 9090 — **breaking**

`/metrics` moved from the MCP port to `M365_METRICS_PORT` (9090); the MCP port answers it 404. Point
the scrape at 9090. Every metric carries `server="microsoft-365-mcp"`.

| Before | Now |
|---|---|
| `m365_tool_calls_total{tool,area,result}` | `mcp_tool_calls_total{tool,result}` |
| `m365_tool_duration_seconds{tool}` | `mcp_tool_duration_seconds{tool}` |
| `m365_guard_refusals_total{reason}` (never incremented) | `mcp_guard_refusals_total{reason}`, counted |
| `m365_link_transfers_total{op,kind,result}` | `mcp_link_transfers_total{op,result}` |
| `m365_link_bytes_total{op}` | `mcp_link_bytes_total{op}` |
| `m365_log_redactions_total` | `mcp_log_records_dropped_total` |
| — | `mcp_http_requests_in_flight{route_group}`, `mcp_http_requests_total`, `mcp_http_request_duration_seconds`, `mcp_upstream_requests_total`, `mcp_upstream_duration_seconds`, `mcp_auth_refusals_total`, `mcp_state_store_up`, `mcp_state_store_errors_total`, `mcp_state_store_duration_seconds`, `mcp_build_info`, `mcp_signin_total` (never moves here) |

### Health

| Endpoint | Before | Now |
|---|---|---|
| `/readyz` | 200 always | 200 once configured, with the store's state in the body; never waits for the store |
| Probes | MCP port | metrics port (also served on the MCP port) |

### Helm chart, published

`oci://ghcr.io/wide-moat/charts/microsoft-365-mcp`: init container staging the key and passwords into a
memory volume, three Service ports, RollingUpdate with `maxUnavailable: 0`, replicas spread across nodes,
PodDisruptionBudget, optional autoscaler, ServiceMonitor and links-only Ingress, no CPU limit, `/tmp` on
disk for the upload spool. [Chart README](charts/microsoft-365-mcp/README.md).

### Image — **breaking** for anyone pulling `latest`

| What | Before | Now |
|---|---|---|
| Tags | `v0.1.0`, `latest` | `0.2.0`, `0.2`; no `latest`: pin a version or the digest |
| Platforms | linux/amd64 | linux/amd64 and linux/arm64, one manifest list |
| Publication | on a tag, untested, unscanned | after every check in `validate.yml` on the tag, and a trivy gate (HIGH/CRITICAL with a fix) before any push |
| Build | `pip install .` (an unlocked build backend) | sources copied; `ARG BASE_IMAGE`, digest-pinned |
| Entry point | `python -m m365_mcp.app` | `python -m m365_mcp.main` (JSON logging before any import) |
| `M365_CA_BUNDLE` | — | extra CAs for Graph and the store |

### Checks added

| Check | What it proves |
|---|---|
| `tests/state` + `scripts/state-contract.sh` | the store contract on memory, Redis 7.4, Valkey 8.1, Valkey behind 3 Sentinels with passwords, Redis behind 1 Sentinel without |
| `scripts/chart-e2e.sh` in kind | 19 checks on 3 replicas and Valkey behind Sentinel, including a killed master and a paused mid-rotation rollout |
| `.github/workflows/known-bad.yml` | the same checks against 0.1.x, where they must fail |
| `tests/unit/test_logs_json.py` | every line of the real processes parses as JSON |
| `ruff format --check`, `ruff check` | in CI; the code base formatted once |
| mutations | 23 new guards, and 12 of the store's scripts against each real server |

### Fixed

- `Retry-After` that is not a number of seconds no longer reads as one (`isdigit()`).
- The live suite's expiry message named an undefined variable.
- `requirements.lock` carried css-inline 0.23.0 while `pyproject.toml` pins 0.23.1.

## 0.1.0: first public release

- 39 task-shaped tools over Microsoft Graph (43 with every area on): mail, calendar and scheduling,
  contacts, OneDrive and SharePoint files, SharePoint lists, Excel, Teams, To Do and Planner, OneNote,
  Bookings (off by default).
- Two-step sending, recipient limits, a send allowance and optional allowed recipient domains.
- Files through short-lived signed links; HTML mail cleaned both ways; text from other people marked
  as data.
- Tests: unit and contract suite, refusal evidence, mutation check, token budget, tool-surface lock,
  and a live suite against a real tenant.
