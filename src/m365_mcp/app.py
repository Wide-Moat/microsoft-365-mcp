"""Three listeners: the MCP endpoint (reachable from the gateway only), the link routes (the one thing
reachable from sandboxes), and the metrics (reachable from Prometheus only). Health and readiness on the
MCP listener; nothing on it or on the links listener serves metrics."""

from __future__ import annotations

import asyncio
import contextlib
import logging
import os
import time
from collections.abc import AsyncIterator, Callable
from dataclasses import dataclass
from typing import Any

from prometheus_client import CONTENT_TYPE_LATEST, generate_latest
from starlette.applications import Starlette
from starlette.requests import Request
from starlette.responses import JSONResponse, Response
from starlette.routing import Mount, Route
from starlette.types import ASGIApp, Message, Receive, Scope, Send

from m365_mcp import __version__
from m365_mcp.config import ConfigError, Settings
from m365_mcp.core.context import Runtime, configure
from m365_mcp.core.cursor import Cursors
from m365_mcp.core.guard import Guard
from m365_mcp.identity import current_areas, current_token
from m365_mcp.keys.harden import non_dumpable
from m365_mcp.links.routes import links_app
from m365_mcp.links.store import LinkStore
from m365_mcp.links.tokens import LinkSigner
from m365_mcp.observability import (
    BUILD_INFO,
    HTTP_IN_FLIGHT,
    HTTP_REQUESTS,
    HTTP_SECONDS,
    configure_logging,
    install_filter,
    status_class,
)
from m365_mcp.registry import AREAS
from m365_mcp.server import M365Server
from m365_mcp.state.memory import MemoryStore
from m365_mcp.state.remote import RedisStore
from m365_mcp.state.store import Namer, StateStore, StoreUnavailable

log = logging.getLogger("m365_mcp")

AREAS_HEADER = b"x-m365-areas"


class Identity:
    """Puts the bearer token and the optional per-person area list into context variables."""

    def __init__(self, app: ASGIApp) -> None:
        self.app = app

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return
        token = None
        areas = None
        for k, v in scope.get("headers", []):
            if k == b"authorization":
                raw = v.decode("latin-1")
                if raw[:7].lower() == "bearer ":
                    token = raw[7:].strip()
            elif k == AREAS_HEADER:
                areas = frozenset(a.strip() for a in v.decode("latin-1").split(",") if a.strip() in AREAS)
        t1 = current_token.set(token)
        t2 = current_areas.set(areas)
        try:
            await self.app(scope, receive, send)
        finally:
            current_token.reset(t1)
            current_areas.reset(t2)


def make_store(settings: Settings) -> StateStore:
    """The store every replica shares (Redis or Valkey), or this process's memory when none is configured."""
    namer = Namer(settings.state_key or os.urandom(32))
    cfg = settings.state
    if cfg is None:
        return MemoryStore(namer)
    if cfg.sentinels:
        return RedisStore.sentinels(
            list(cfg.sentinels),
            cfg.master,
            namer,
            cfg.prefix,
            username=cfg.username or None,
            password=cfg.password,
            sentinel_password=cfg.sentinel_password,
            tls=cfg.tls,
            db=cfg.db or 0,
            ca_bundle=settings.ca_bundle,
        )
    return RedisStore.standalone(
        cfg.url,
        namer,
        cfg.prefix,
        username=cfg.username or None,
        password=cfg.password,
        db=cfg.db,
        ca_bundle=settings.ca_bundle,
    )


def make_signer(settings: Settings, store: StateStore) -> LinkSigner | None:
    if not settings.links_ready:
        return None
    return LinkSigner(
        settings.sign_keys,
        settings.enc_keys,
        LinkStore(store, key=settings.state_key or os.urandom(32), previous=settings.previous_state_keys),
    )


_METHODS = frozenset({"GET", "HEAD", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"})


def main_group(path: str) -> str:
    if path == "/mcp" or path.startswith("/mcp/"):
        return "mcp"
    if path in ("/healthz", "/readyz"):
        return "health"
    return "other"


def links_group(path: str) -> str:
    return "links" if path.startswith(("/d/", "/u/")) else "other"


class HttpMetrics:
    """Requests in flight (the autoscaling signal), counted and timed, by route group. Never the path itself:
    a link's path is its credential."""

    def __init__(self, app: ASGIApp, group_of: Callable[[str], str]) -> None:
        self.app = app
        self.group_of = group_of

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return
        group = self.group_of(scope.get("path", ""))
        method = scope.get("method", "GET") if scope.get("method") in _METHODS else "other"
        status = 0

        async def recorded(message: Message) -> None:
            nonlocal status
            if message["type"] == "http.response.start":
                status = message["status"]
            await send(message)

        flight = HTTP_IN_FLIGHT.labels(group)
        flight.inc()
        t0 = time.perf_counter()
        try:
            await self.app(scope, receive, recorded)
        finally:
            flight.dec()
            HTTP_SECONDS.labels(group).observe(time.perf_counter() - t0)
            HTTP_REQUESTS.labels(group, method, status_class(status or 500)).inc()


@dataclass
class Apps:
    main: Starlette
    links: Starlette
    metrics: Starlette
    store: StateStore


PING_EVERY = 5.0


async def _ping(store: StateStore) -> bool:
    try:
        await store.ping()
        return True
    except StoreUnavailable:
        return False


def build(settings: Settings, store: StateStore | None = None) -> Apps:
    install_filter()
    store = store if store is not None else make_store(settings)
    configure(
        Runtime(
            settings=settings,
            signer=make_signer(settings, store),
            cursors=Cursors(settings.sign_keys[0] if settings.sign_keys else None),
            guard=Guard(store, previous=tuple(Namer(k) for k in settings.previous_state_keys)),
        )
    )
    server = M365Server(settings)
    mcp_app = server.streamable_http_app(
        streamable_http_path="/mcp", json_response=True, stateless_http=True, host="0.0.0.0"
    )

    async def healthz(_: Request) -> Response:
        return JSONResponse({"ok": True})

    async def readyz(_: Request) -> Response:
        # Ready means configured: the settings parsed and the key loaded, or the process would not be running.
        # The shared store is reported, never required: with the store down, reads keep serving and sends
        # and new links fail closed in code; a readiness tied to the store would take every pod out of the
        # Service at once (mcp_state_store_up and its alert say the store is down).
        up = await _ping(store)
        return JSONResponse(
            {
                "ok": True,
                "areas": sorted(settings.areas),
                "links": settings.links_ready,
                "key": settings.key_source or "none",
                "state": store.kind,
                "store": "up" if up else "down",
            }
        )

    async def metrics(_: Request) -> Response:
        return Response(generate_latest(), media_type=CONTENT_TYPE_LATEST)

    @contextlib.asynccontextmanager
    async def lifespan(app: Starlette) -> AsyncIterator[None]:
        async def keep_watch() -> None:
            # mcp_state_store_up stays true between calls too, without waiting for a probe or a send.
            while True:
                await _ping(store)
                await asyncio.sleep(PING_EVERY)

        watcher = asyncio.create_task(keep_watch())
        try:
            async with mcp_app.router.lifespan_context(app):
                yield
        finally:
            watcher.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await watcher
            await store.aclose()

    main = Starlette(
        routes=[
            Route("/healthz", healthz),
            Route("/readyz", readyz),
            Mount("/", app=Identity(mcp_app)),
        ],
        lifespan=lifespan,
    )
    main.add_middleware(HttpMetrics, group_of=main_group)
    # Every route group's in-flight gauge exists from the start, at 0: an autoscaler reading it finds a
    # value on a pod that has served nothing yet, not a missing metric.
    for group in ("mcp", "links", "health", "other"):
        HTTP_IN_FLIGHT.labels(group)
    links = links_app()
    links.add_middleware(HttpMetrics, group_of=links_group)
    # The probes are here too, so the kubelet never needs the MCP port.
    metrics_app = Starlette(routes=[Route("/metrics", metrics), Route("/healthz", healthz), Route("/readyz", readyz)])
    return Apps(main=main, links=links, metrics=metrics_app, store=store)


def uvicorn_config(app: object, port: int) -> Any:
    """uvicorn's default dictConfig gives its loggers handlers of their own, which the content filter on
    the root handlers never sees; without it they propagate to those handlers."""
    import uvicorn

    return uvicorn.Config(app, host="0.0.0.0", port=port, log_level="warning", access_log=False, log_config=None)


def _port(name: str, default: int) -> int:
    raw = os.environ.get(name, "").strip()
    try:
        port = int(raw) if raw else default
    except ValueError:
        port = 0
    if not 0 < port < 65536:
        log.critical(f"{name}: a port number", extra={"event": "config_error"})
        raise SystemExit(2)
    return port


def run() -> None:
    import uvicorn

    # Before any setting is read: nothing of this process's memory is readable by another of its user.
    hardened = non_dumpable()
    # Handlers first, then the filter on them: build() alone would find no handler to filter.
    try:
        configure_logging()
        settings = Settings.from_env()
    except (ConfigError, ValueError) as e:
        # Our own sentence, naming a setting and never a value.
        log.critical(str(e), extra={"event": "config_error"})
        raise SystemExit(2) from None
    for warning in settings.warnings:
        log.warning(warning, extra={"event": "config_warning"})
    BUILD_INFO.labels(__version__).set(1)
    apps = build(settings)
    port, links_port = _port("PORT", 8000), _port("LINKS_PORT", 8001)
    ports = {port, links_port, settings.metrics_port}
    if len(ports) != 3:
        log.critical("PORT, LINKS_PORT and M365_METRICS_PORT must differ", extra={"event": "config_error"})
        raise SystemExit(2)
    log.info(
        "starting",
        extra={
            "event": "start",
            "fields": {
                "version": __version__,
                "ports": [port, links_port, settings.metrics_port],
                "store": apps.store.kind,
                "mode": settings.key_source or "none",
                "status": "non-dumpable" if hardened else "dumpable",
            },
        },
    )

    async def serve() -> None:
        servers = [
            uvicorn.Server(uvicorn_config(apps.main, port)),
            uvicorn.Server(uvicorn_config(apps.links, links_port)),
            uvicorn.Server(uvicorn_config(apps.metrics, settings.metrics_port)),
        ]
        await asyncio.gather(*(s.serve() for s in servers))

    asyncio.run(serve())


if __name__ == "__main__":
    run()
