"""The assembled application, the way the container runs it. Every other test calls the server object
directly, so a mistake in wiring it into HTTP would pass the whole suite and crash the pod at start."""

from __future__ import annotations

from helpers import DownStore
from starlette.testclient import TestClient

from m365_mcp.app import build
from m365_mcp.config import Settings


def test_the_application_starts_and_answers_health(settings: Settings) -> None:
    apps = build(settings)
    with TestClient(apps.main) as http:
        assert http.get("/healthz").json() == {"ok": True}
        ready = http.get("/readyz")
    assert ready.status_code == 200 and ready.json()["store"] == "up" and ready.json()["state"] == "memory"


def test_the_mcp_endpoint_lists_tools_over_http(settings: Settings) -> None:
    apps = build(settings)
    with TestClient(apps.main) as http:
        r = http.post(
            "/mcp",
            json={"jsonrpc": "2.0", "id": 1, "method": "tools/list", "params": {}},
            headers={"Accept": "application/json, text/event-stream"},
        )
    assert r.status_code == 200
    names = {t["name"] for t in r.json()["result"]["tools"]}
    assert {"get_connection", "list_messages", "send_draft"} <= names


def test_metrics_are_served_on_their_own_listener_only(settings: Settings) -> None:
    apps = build(settings)
    with TestClient(apps.main) as main, TestClient(apps.links) as links, TestClient(apps.metrics) as metrics:
        main.get("/healthz")
        assert main.get("/metrics").status_code == 404, "the MCP listener serves metrics"
        assert links.get("/metrics").status_code == 404, "the links listener serves metrics"
        body = metrics.get("/metrics").text
    assert 'mcp_http_requests_in_flight{route_group="health",server="microsoft-365-mcp"}' in body
    assert (
        'mcp_http_requests_total{method="GET",route_group="health",server="microsoft-365-mcp",status_class="2xx"}'
        in body
    )
    assert 'mcp_state_store_up{server="microsoft-365-mcp"} 1.0' in body


def test_readiness_never_depends_on_the_store(settings: Settings) -> None:
    # A store outage must not take every pod out of the Service: reads keep serving, sends fail closed.
    apps = build(settings, store=DownStore())  # type: ignore[arg-type]
    with TestClient(apps.main) as http, TestClient(apps.metrics) as probes:
        assert http.get("/healthz").status_code == 200
        r = http.get("/readyz")
        assert probes.get("/healthz").status_code == 200, "the probes are on the metrics port too"
        on_metrics_port = probes.get("/readyz")
    assert r.status_code == 200 and r.json()["store"] == "down" and r.json()["ok"] is True
    assert on_metrics_port.status_code == 200 and on_metrics_port.json()["store"] == "down"


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
