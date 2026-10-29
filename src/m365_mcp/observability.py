"""Logs and metrics that never carry content: no subjects, bodies, addresses, names, tokens."""

from __future__ import annotations

import json
import logging
import re
import time
import traceback
from collections.abc import Iterator
from contextlib import contextmanager
from typing import Any

from prometheus_client import Counter, Histogram

TOOL_CALLS = Counter("m365_tool_calls_total", "Tool calls", ["tool", "area", "result"])
TOOL_SECONDS = Histogram("m365_tool_duration_seconds", "Tool call duration", ["tool"])
GUARD_REFUSALS = Counter("m365_guard_refusals_total", "Refusals by the send guard", ["reason"])
LINK_TRANSFERS = Counter("m365_link_transfers_total", "Signed-link transfers", ["op", "kind", "result"])
LINK_BYTES = Counter("m365_link_bytes_total", "Bytes moved through signed links", ["op"])
REDACTIONS = Counter("m365_log_redactions_total", "Log records dropped because they looked like content")

_log = logging.getLogger("m365_mcp.calls")
_ALLOWED = {"event", "tool", "area", "person", "ms", "result", "graph_calls", "recipients", "bytes", "op", "kind"}
_CONTENT = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}|eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}")


def _bare_traceback(exc_info: Any) -> str:
    """Where it failed, without what it said: the frames and the exception types, never an exception's
    message or a local value, which can hold a subject or an address that the pattern cannot recognise."""
    part: traceback.TracebackException | None = traceback.TracebackException(*exc_info, capture_locals=False)
    lines: list[str] = []
    seen = 0
    while part is not None and seen < 5:
        lines.append("Traceback (most recent call last), messages withheld:")
        lines += [f'  File "{fr.filename}", line {fr.lineno}, in {fr.name}' for fr in part.stack]
        lines.append(getattr(part, "exc_type_str", None) or getattr(part.exc_type, "__qualname__", "Exception"))
        part = part.__cause__ or part.__context__
        seen += 1
    return "\n".join(lines)


class ContentFilter(logging.Filter):
    """Drops any record that looks like it carries an address or a token, and counts it. A traceback
    keeps its frames and loses every exception message; a stack that looks like content is dropped."""

    def filter(self, record: logging.LogRecord) -> bool:
        try:
            msg = record.getMessage()
        except Exception:  # noqa: BLE001 - a record that cannot render is not logged
            REDACTIONS.inc()
            return False
        if _CONTENT.search(msg):
            REDACTIONS.inc()
            return False
        if record.exc_info and record.exc_info[0] is not None:
            record.exc_text = _bare_traceback(record.exc_info)
            record.exc_info = None
        elif record.exc_text and _CONTENT.search(record.exc_text):
            REDACTIONS.inc()
            record.exc_text = "(traceback withheld: it looked like content)"
        if record.stack_info and _CONTENT.search(record.stack_info):
            REDACTIONS.inc()
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


def configure_logging() -> None:
    """The process's logging: one root handler, then the filter on it. uvicorn is started with
    log_config=None so its loggers keep no handlers of their own and propagate here."""
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    install_filter()


def log_call(tool: str, **fields: Any) -> None:
    rec = {"event": "tool", "tool": tool}
    rec.update({k: v for k, v in fields.items() if k in _ALLOWED})
    _log.info(json.dumps(rec, separators=(",", ":")))


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
        TOOL_CALLS.labels(tool, area, rec["result"]).inc()
        TOOL_SECONDS.labels(tool).observe(ms / 1000)
        log_call(tool, area=area, ms=ms, result=rec["result"])
