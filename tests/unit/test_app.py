"""The assembled application, the way the container runs it. Every other test calls the server object
directly, so a mistake in wiring it into HTTP would pass the whole suite and crash the pod at start."""

from __future__ import annotations

from starlette.testclient import TestClient

from m365_mcp.app import build
from m365_mcp.config import Settings


def test_the_application_starts_and_answers_health(settings: Settings) -> None:
    main, _links = build(settings)
    with TestClient(main) as http:
        assert http.get("/healthz").json() == {"ok": True}


def test_the_mcp_endpoint_lists_tools_over_http(settings: Settings) -> None:
    main, _links = build(settings)
    with TestClient(main) as http:
        r = http.post("/mcp", json={"jsonrpc": "2.0", "id": 1, "method": "tools/list", "params": {}},
                      headers={"Accept": "application/json, text/event-stream"})
    assert r.status_code == 200
    names = {t["name"] for t in r.json()["result"]["tools"]}
    assert {"get_connection", "list_messages", "send_draft"} <= names


def test_uvicorn_logs_through_the_root_handlers_and_their_content_filter() -> None:
    import logging

    from m365_mcp.app import uvicorn_config
    from m365_mcp.observability import ContentFilter, configure_logging
    cfg = uvicorn_config(object(), 8001)
    assert cfg.log_config is None, "uvicorn's own dictConfig would give its loggers handlers outside the filter"
    for name in ("uvicorn", "uvicorn.error"):
        lg = logging.getLogger(name)
        assert not lg.handlers and lg.propagate, name
    assert not logging.getLogger("uvicorn.access").handlers and cfg.access_log is False, "no access log at all"
    configure_logging()
    handlers = logging.getLogger().handlers
    assert handlers and all(any(isinstance(f, ContentFilter) for f in h.filters) for h in handlers)
