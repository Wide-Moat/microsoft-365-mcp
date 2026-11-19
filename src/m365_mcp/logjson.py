"""Every line this process writes is one JSON object. Standard library only, so the entry point can install
it before anything else is imported, and a failing import is logged as JSON too.

A line: `ts` (RFC 3339, UTC, milliseconds), `level`, `logger`, `event` (a stable key: what happened),
`msg` (a sentence for a person), the allowlisted fields a caller attached, and `exc` for a traceback:
its frames and exception types, never an exception's message, which can hold a subject or an address.
"""

from __future__ import annotations

import json
import logging
import sys
import threading
import traceback
from datetime import UTC, datetime
from typing import Any

# Fields a record may carry beside the fixed ones. A field outside this list is dropped, not written.
ALLOWED_FIELDS = frozenset(
    {
        "tool",
        "area",
        "person",
        "ms",
        "result",
        "graph_calls",
        "recipients",
        "bytes",
        "op",
        "kind",
        "port",
        "ports",
        "files",
        "names",
        "mode",
        "replicas",
        "status",
        "store",
        "setting",
        "reason",
        "version",
    }
)


def bare_traceback(exc_info: Any) -> str:
    """Where it failed, without what it said: the frames and the exception types, never an exception's
    message or a local value, which can hold a subject or an address that no pattern recognises."""
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


def _ts(created: float) -> str:
    return datetime.fromtimestamp(created, UTC).isoformat(timespec="milliseconds").replace("+00:00", "Z")


class JsonFormatter(logging.Formatter):
    """One `json.dumps` per record: a quote, a newline or a traceback inside a value stays inside it."""

    def format(self, record: logging.LogRecord) -> str:
        try:
            msg = record.getMessage()
        except Exception:  # noqa: BLE001 - a record whose arguments do not fit its template still gets a line
            msg = str(record.msg)
        out: dict[str, Any] = {
            "ts": _ts(record.created),
            "level": record.levelname.lower(),
            "logger": record.name,
            "event": getattr(record, "event", None) or "log",
            "msg": msg,
        }
        fields = getattr(record, "fields", None)
        if isinstance(fields, dict):
            out.update({k: v for k, v in fields.items() if k in ALLOWED_FIELDS and k not in out})
        if record.exc_info and record.exc_info[0] is not None:
            # No filter ran (or it did not cover this handler): the frames still go without the messages.
            out["exc"] = bare_traceback(record.exc_info)
        elif record.exc_text:
            out["exc"] = record.exc_text
        if record.stack_info:
            out["stack"] = record.stack_info
        return json.dumps(out, separators=(",", ":"), ensure_ascii=False, default=str)


def handler(stream: Any = None) -> logging.Handler:
    h = logging.StreamHandler(stream if stream is not None else sys.stderr)
    h.setFormatter(JsonFormatter())
    return h


def install(level: int = logging.INFO) -> None:
    """The root logger writes JSON to stderr through exactly one handler, and nothing reaches the streams
    any other way: uncaught exceptions in any thread and Python warnings go through it too."""
    root = logging.getLogger()
    for h in list(root.handlers):
        root.removeHandler(h)
    root.addHandler(handler())
    root.setLevel(level)
    logging.captureWarnings(True)

    def excepthook(kind: type[BaseException], value: BaseException, tb: Any) -> None:
        if issubclass(kind, KeyboardInterrupt):
            return
        logging.getLogger("m365_mcp").critical(
            "the process stopped on an uncaught exception", exc_info=(kind, value, tb), extra={"event": "uncaught"}
        )

    def thread_hook(args: threading.ExceptHookArgs) -> None:
        if args.exc_type is SystemExit:
            return
        logging.getLogger("m365_mcp").error(
            "a thread stopped on an uncaught exception",
            exc_info=(args.exc_type, args.exc_value, args.exc_traceback),
            extra={"event": "uncaught_thread"},
        )

    sys.excepthook = excepthook
    threading.excepthook = thread_hook


def emit(event: str, msg: str, *, level: int = logging.INFO, logger: str = "m365_mcp", **fields: Any) -> None:
    """One line, from a process that may not have installed logging (the init container's staging step)."""
    rec = logging.LogRecord(logger, level, "", 0, msg, None, None)
    rec.event = event  # type: ignore[attr-defined]
    rec.fields = fields  # type: ignore[attr-defined]
    stream = sys.stderr if level >= logging.WARNING else sys.stdout
    stream.write(JsonFormatter().format(rec) + "\n")
    stream.flush()
