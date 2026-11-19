# Component: observability

What operators can see, and what they never can.

## Logs

Every line every process of the image writes is one JSON object: the server from its first line to
its last, uvicorn, a start refused for its settings, an uncaught exception, and the init container's
staging step. Fields: `ts` (RFC 3339, UTC, milliseconds), `level`, `logger`, `event` (a stable key),
`msg`, the allowlisted fields of the record, and `exc` for a traceback.

```json
{"ts":"2026-11-19T10:02:11.481Z","level":"info","logger":"m365_mcp.calls","event":"tool","msg":"tool call",
 "tool":"send_draft","area":"mail","ms":412,"result":"ok"}
```

| Event | When |
|---|---|
| `start` | the server listens: version, ports, store kind, key source |
| `config_error` | a setting or the key file is refused; the process exits 2 |
| `config_warning` | a deprecated key in the environment, or memory as the store |
| `tool` | one per tool call: tool, area, duration, result |
| `uncaught` | an exception nobody caught: frames and types only |
| `staged`, `stage_failed`, `stage_empty` | the init container copied the Secrets, or could not |
| `log` | anything else, from a library |

`LOG_LEVEL` (`DEBUG`, `INFO`, `WARNING`, `ERROR`) applies to every logger; the HTTP stack stays at
`WARNING` whatever it says, because its `INFO` lines carry Graph URLs.

**Forbidden in any line:** subjects, bodies, previews, addresses, names, file names, folder names,
search strings, tokens, link ids. A filter drops any record whose message or fields contain `@`
followed by a domain or a JWT-shaped string, and counts the drop. A traceback keeps its frames (file,
line, function) and exception types and loses every exception message. One `json.dumps` per record:
a quote, a newline or a traceback inside a value stays inside it.

## Metrics (Prometheus)

On their own port (`M365_METRICS_PORT`, 9090), path `/metrics`; never on the MCP or the links port.
The same names as exchange-mcp and vkteams-mcp, with a constant label `server="microsoft-365-mcp"`, so
one dashboard and one autoscaling rule fit all three. No label carries a person, an address or an
argument.

| Metric | Labels | Notes |
|---|---|---|
| `mcp_http_requests_in_flight` | `route_group` = mcp, links, health, other | the autoscaling signal |
| `mcp_http_requests_total` | `route_group`, `method`, `status_class` | |
| `mcp_http_request_duration_seconds` | `route_group` | histogram; a download counts until its last byte |
| `mcp_tool_calls_total` | `tool`, `result` = ok, error, crash | |
| `mcp_tool_duration_seconds` | `tool` | histogram |
| `mcp_upstream_requests_total` | `status_class` (or `error`: no answer) | Microsoft Graph and its file hosts |
| `mcp_upstream_duration_seconds` | | to the answer's headers |
| `mcp_auth_refusals_total` | `reason` = invalid_token, invalid_client | a forwarded token refused |
| `mcp_guard_refusals_total` | `reason` = confirm, rate, too_many_recipients, distribution_list, outside_domains, state_unavailable | |
| `mcp_link_transfers_total` | `op` = d, u; `result` | |
| `mcp_state_store_up` | | 1 when the store answered its last operation; pinged every 5 s |
| `mcp_state_store_errors_total` | `op` | |
| `mcp_state_store_duration_seconds` | `op` | histogram |
| `mcp_log_records_dropped_total` | | must stay 0: non-zero means code tried to log content |
| `mcp_build_info` | `version` | always 1 |
| `mcp_signin_total` | `result` | common to the three servers; never moves here, the gateway signs people in |
| `mcp_link_bytes_total` | `op` | this server only |

Python's process and runtime collectors are on.

Renamed in 0.2.0 (see CHANGELOG): every `m365_*` metric, and `m365_link_transfers_total` lost its
`kind` label, `m365_tool_calls_total` its `area` label.

## Tests

- A full run of every tool with recorded fixtures produces logs with only allowed keys, and none of
  the fixture's addresses, subjects or file names appear anywhere in captured output.
- The real processes are started (`tests/unit/test_logs_json.py`): the server, a refused start, an
  uncaught exception and the staging step, both ways; every line must parse as JSON.
- `tests/unit/test_app.py` reads the metrics from their own listener and finds none on the others.
- The chart e2e parses every log line of every container, the init container included, and scrapes
  every pod's metrics.
