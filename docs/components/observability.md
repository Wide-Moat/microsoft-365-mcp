# Component: observability

What operators can see, and what they never can.

## Logs

One JSON line per tool call and per link transfer:

```json
{"ts":"…","event":"tool","tool":"send_draft","area":"mail","person":"3f9c…","ms":412,
 "result":"ok","graph_calls":3,"graph_status":[200,200,202],"request_id":"…","recipients":2}
```

Allowed fields: tool, area, person hash, duration, result code, Graph status codes, Graph request
ids, counts (recipients, items, bytes). **Forbidden:** subjects, bodies, previews, addresses, names,
file names, folder names, search strings, tokens, link tokens. A logging filter drops any record
whose message contains `@` followed by a domain or a JWT-shaped string, and counts the drop as a
defect metric.

A traceback keeps its frames (file, line, function) and the exception types, and loses every
exception message, which can hold a subject or an address the pattern would not recognise; a stack
that looks like content is dropped. The filter sits on the root logger's handlers, installed after
they exist, and uvicorn runs with `log_config=None`, so its loggers keep no handlers of their own and
reach the same filtered handlers.

## Metrics (Prometheus)

- `m365_tool_calls_total{tool,area,result}`
- `m365_tool_duration_seconds{tool}` (histogram)
- `m365_graph_requests_total{area,status}`
- `m365_graph_throttled_total{area}`
- `m365_guard_refusals_total{reason}`
- `m365_link_transfers_total{op,kind,result}` and `m365_link_bytes_total{op}`
- `m365_log_redactions_total` — must stay 0; non-zero means code is trying to log content.

## Tests

- A full run of every tool with recorded fixtures produces logs with only allowed keys, and none of
  the fixture's addresses, subjects or file names appear anywhere in captured output.
- The redaction counter rises when a test deliberately logs an address.
