"""Logs and metrics that never carry content: no subjects, bodies, addresses, names, tokens.

Metric names are the ones exchange-mcp and vkteams-mcp use too: `mcp_*`, with a constant label
`server="microsoft-365-mcp"`, so one dashboard and one autoscaling rule fit all three servers. Labels
never carry a person, an address or a tool argument. They are served on their own port
(`M365_METRICS_PORT`, 9090), never on the MCP or the links port.
"""

from __future__ import annotations

import logging
import os
import re
import time
from collections.abc import Iterator
from contextlib import contextmanager
from typing import Any

from prometheus_client import Counter, Gauge, Histogram

from m365_mcp.logjson import JsonFormatter, bare_traceback
from m365_mcp.logjson import install as install_json

SERVER = "microsoft-365-mcp"

# Route groups of the HTTP metrics. A path that is none of them counts as "other".
ROUTE_GROUPS = ("mcp", "oauth", "signin", "links", "health", "other")


class _Server:
    """A metric whose first label is always `server`, given here once."""

    def __init__(self, metric: Any) -> None:
        self._m = metric

    def labels(self, *values: str) -> Any:
        return self._m.labels(SERVER, *values)

    def one(self) -> Any:
        return self._m.labels(SERVER)


def _counter(name: str, doc: str, labels: tuple[str, ...] = ()) -> _Server:
    return _Server(Counter(name, doc, ("server", *labels)))


def _gauge(name: str, doc: str, labels: tuple[str, ...] = ()) -> _Server:
    return _Server(Gauge(name, doc, ("server", *labels)))


def _histogram(name: str, doc: str, labels: tuple[str, ...] = (), buckets: tuple[float, ...] | None = None) -> _Server:
    kw = {"buckets": buckets} if buckets else {}
    return _Server(Histogram(name, doc, ("server", *labels), **kw))


# The common set.
HTTP_IN_FLIGHT = _gauge("mcp_http_requests_in_flight", "HTTP requests being served now", ("route_group",))
HTTP_REQUESTS = _counter("mcp_http_requests_total", "HTTP requests served", ("route_group", "method", "status_class"))
HTTP_SECONDS = _histogram("mcp_http_request_duration_seconds", "HTTP request duration", ("route_group",))
TOOL_CALLS = _counter("mcp_tool_calls_total", "Tool calls", ("tool", "result"))
TOOL_SECONDS = _histogram("mcp_tool_duration_seconds", "Tool call duration", ("tool",))
UPSTREAM_REQUESTS = _counter(
    "mcp_upstream_requests_total", "Requests to Microsoft Graph and its file hosts", ("status_class",)
)
UPSTREAM_SECONDS = _histogram("mcp_upstream_duration_seconds", "Microsoft Graph response time, to the headers")
SIGNIN = _counter("mcp_signin_total", "Sign-ins on this server (none here: the gateway signs people in)", ("result",))
AUTH_REFUSALS = _counter("mcp_auth_refusals_total", "Calls refused for their token", ("reason",))
GUARD_REFUSALS = _counter("mcp_guard_refusals_total", "Refusals by the send guard", ("reason",))
LINK_TRANSFERS = _counter("mcp_link_transfers_total", "Signed-link transfers", ("op", "result"))
STATE_UP = _gauge("mcp_state_store_up", "1 when the shared state store answered its last operation")
STATE_ERRORS = _counter("mcp_state_store_errors_total", "State store operations that failed", ("op",))
STATE_SECONDS = _histogram(
    "mcp_state_store_duration_seconds",
    "State store operation duration",
    ("op",),
    buckets=(0.001, 0.0025, 0.005, 0.01, 0.025, 0.05, 0.1, 0.25, 0.5, 1.0, 2.5),
)
LOG_DROPPED = _counter("mcp_log_records_dropped_total", "Log records dropped because they looked like content")
BUILD_INFO = _gauge("mcp_build_info", "Always 1; its labels say which build runs", ("version",))
# Specific to this server.
LINK_BYTES = _counter("mcp_link_bytes_total", "Bytes moved through signed links", ("op",))


def status_class(status: int) -> str:
    return f"{status // 100}xx" if 100 <= status < 600 else "error"


_log = logging.getLogger("m365_mcp.calls")
_CONTENT = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}|eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}")

# The reference implementation of a traceback without its messages; kept under its old name.
_bare_traceback = bare_traceback


class ContentFilter(logging.Filter):
    """Drops any record that looks like it carries an address or a token, and counts it. A traceback
    keeps its frames and loses every exception message; a stack that looks like content is dropped."""

    def filter(self, record: logging.LogRecord) -> bool:
        try:
            msg = record.getMessage()
        except Exception:  # noqa: BLE001 - a record that cannot render is not logged
            LOG_DROPPED.one().inc()
            return False
        fields = getattr(record, "fields", None)
        if _CONTENT.search(msg) or (isinstance(fields, dict) and _CONTENT.search(repr(fields))):
            LOG_DROPPED.one().inc()
            return False
        if record.exc_info and record.exc_info[0] is not None:
            record.exc_text = bare_traceback(record.exc_info)
            record.exc_info = None
        elif record.exc_text and _CONTENT.search(record.exc_text):
            LOG_DROPPED.one().inc()
            record.exc_text = "(traceback withheld: it looked like content)"
        if record.stack_info and _CONTENT.search(record.stack_info):
            LOG_DROPPED.one().inc()
            record.stack_info = None
        return True


def install_filter() -> None:
    """Puts the filter on the root logger's handlers, which every other logger reaches by propagation.
    Call it after the handlers exist (configure_logging does); a handler added later is not covered."""
    root = logging.getLogger()
    for h in root.handlers:
        if not any(isinstance(f, ContentFilter) for f in h.filters):
            h.addFilter(ContentFilter())
    if not any(isinstance(f, ContentFilter) for f in root.filters):
        root.addFilter(ContentFilter())


def log_level() -> int:
    """`LOG_LEVEL` (DEBUG, INFO, WARNING, ERROR), INFO when unset; it applies to every logger."""
    raw = os.environ.get("LOG_LEVEL", "").strip().upper() or "INFO"
    level = logging.getLevelName(raw)
    if raw not in ("DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL") or not isinstance(level, int):
        raise ValueError(f"LOG_LEVEL: one of DEBUG, INFO, WARNING, ERROR, CRITICAL, not {raw[:20]!r}")
    return level


def configure_logging() -> None:
    """The process's logging: one root handler writing one JSON object per line, then the filter on it.
    uvicorn is started with log_config=None so its loggers keep no handlers of their own and propagate
    here. The HTTP stack stays at WARNING whatever the level: its INFO lines carry Graph URLs."""
    install_json(logging.INFO)
    install_filter()
    logging.getLogger().setLevel(log_level())


def log_call(tool: str, **fields: Any) -> None:
    rec = {"tool": tool, **fields}
    _log.info("tool call", extra={"event": "tool", "fields": rec})


@contextmanager
def tool_timer(tool: str, area: str) -> Iterator[dict[str, Any]]:
    rec: dict[str, Any] = {"result": "ok"}
    t0 = time.monotonic()
    try:
        yield rec
    except Exception:
        rec["result"] = "crash"
        raise
    finally:
        ms = int((time.monotonic() - t0) * 1000)
        TOOL_CALLS.labels(tool, rec["result"]).inc()
        TOOL_SECONDS.labels(tool).observe(ms / 1000)
        log_call(tool, area=area, ms=ms, result=rec["result"])


__all__ = ["JsonFormatter", "ContentFilter", "configure_logging", "install_filter", "log_call", "tool_timer"]
