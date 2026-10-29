"""Two listeners: the MCP endpoint (reachable from the gateway only) and the link routes
(the one thing reachable from sandboxes). Health endpoints on the MCP listener."""

from __future__ import annotations

import contextlib
import logging
from collections.abc import AsyncIterator
from typing import Any

from prometheus_client import CONTENT_TYPE_LATEST, generate_latest
from starlette.applications import Starlette
from starlette.requests import Request
from starlette.responses import JSONResponse, Response
from starlette.routing import Mount, Route
from starlette.types import ASGIApp, Receive, Scope, Send

from m365_mcp.config import Settings
from m365_mcp.core.context import Runtime, configure
from m365_mcp.core.cursor import Cursors
from m365_mcp.identity import current_areas, current_token
from m365_mcp.links.routes import links_app
from m365_mcp.links.store import LinkStore
from m365_mcp.links.tokens import LinkSigner
from m365_mcp.observability import configure_logging, install_filter
from m365_mcp.registry import AREAS
from m365_mcp.server import M365Server

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


def build(settings: Settings) -> tuple[Starlette, Starlette]:
    install_filter()
    configure(Runtime(
        settings=settings,
        signer=LinkSigner(settings.sign_keys, settings.enc_keys, LinkStore()) if settings.links_ready else None,
        cursors=Cursors(settings.sign_keys[0] if settings.sign_keys else None),
    ))
    server = M365Server(settings)
    mcp_app = server.streamable_http_app(streamable_http_path="/mcp", json_response=True, stateless_http=True,
                                         host="0.0.0.0")

    async def healthz(_: Request) -> Response:
        return JSONResponse({"ok": True})

    async def readyz(_: Request) -> Response:
        return JSONResponse({"ok": True, "areas": sorted(settings.areas), "links": settings.links_ready})

    async def metrics(_: Request) -> Response:
        return Response(generate_latest(), media_type=CONTENT_TYPE_LATEST)

    @contextlib.asynccontextmanager
    async def lifespan(app: Starlette) -> AsyncIterator[None]:
        async with mcp_app.router.lifespan_context(app):
            yield

    main = Starlette(routes=[
        Route("/healthz", healthz), Route("/readyz", readyz), Route("/metrics", metrics),
        Mount("/", app=Identity(mcp_app)),
    ], lifespan=lifespan)
    return main, links_app()


def uvicorn_config(app: object, port: int) -> Any:
    """uvicorn's default dictConfig gives its loggers handlers of their own, which the content filter on
    the root handlers never sees; without it they propagate to those handlers."""
    import uvicorn

    return uvicorn.Config(app, host="0.0.0.0", port=port, log_level="warning", access_log=False, log_config=None)


def run() -> None:
    import asyncio
    import os

    import uvicorn

    # Handlers first, then the filter on them: build() alone would find no handler to filter.
    configure_logging()
    settings = Settings.from_env()
    main, links = build(settings)
    port, links_port = int(os.environ.get("PORT", "8000")), int(os.environ.get("LINKS_PORT", "8001"))

    async def serve() -> None:
        servers = [uvicorn.Server(uvicorn_config(main, port)), uvicorn.Server(uvicorn_config(links, links_port))]
        await asyncio.gather(*(s.serve() for s in servers))

    asyncio.run(serve())


if __name__ == "__main__":
    run()
