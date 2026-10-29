"""What the live tests stand on: two signed-in people, a configured server, links, cleanup, findings.

Nothing here is a fake. Tool calls go through the real server handler and the real Graph client
over the network; links are served by the real links app, in process.
"""

from __future__ import annotations

import asyncio
import base64
import datetime as dt
import json
import os
import re
import sys
import time
import uuid
import warnings
from collections.abc import Awaitable, Callable
from contextvars import Token
from pathlib import Path
from typing import Any, NamedTuple

import httpx
from mcp.types import CallToolRequestParams
from starlette.testclient import TestClient

from m365_mcp.config import Settings
from m365_mcp.core.context import Runtime, configure
from m365_mcp.core.cursor import Cursors
from m365_mcp.core.guard import Guard
from m365_mcp.identity import current_areas, current_token
from m365_mcp.links.routes import links_app
from m365_mcp.links.store import LinkStore
from m365_mcp.links.tokens import LinkSigner
from m365_mcp.registry import AREAS
from m365_mcp.server import M365Server

LINKS_BASE = "https://links.live.invalid"
GRAPH = "https://graph.microsoft.com/v1.0"
MIN_TOKEN_LIFE = 600  # seconds a token must still have when the run starts
_LINK = re.compile(r"https://links\.live\.invalid/[du]/[^'\s\"]+")


class Result(NamedTuple):
    error: bool
    text: str
    data: Any


def token_problem(tokens: dict | None, now: float | None = None) -> str | None:
    """Why the tokens cannot be used, or None."""
    now = time.time() if now is None else now
    if not tokens:
        return "no tokens"
    for who in ("a", "b"):
        t = tokens.get(who)
        if not isinstance(t, dict) or not t.get("access_token"):
            return f"user {who} is missing"
        try:
            expires = int(t.get("expires_at") or 0)
        except (TypeError, ValueError):
            return f"the {label} token in the tokens file has no usable expiry; run scripts/live_login.py again"
        if expires - now < MIN_TOKEN_LIFE:
            return f"the token of user {who} has expired or expires within {MIN_TOKEN_LIFE // 60} minutes"
    return None


def load_tokens(path: Path) -> dict | None:
    try:
        data = json.loads(path.read_text())
    except (OSError, ValueError):
        return None
    return data if isinstance(data, dict) else None


def claims(token: str) -> dict:
    payload = token.split(".")[1]
    return json.loads(base64.urlsafe_b64decode(payload + "=" * (-len(payload) % 4)))


class Acting:
    """`async with as_a:` makes the tools in the block act as user A."""

    def __init__(self, live: Live, user: str) -> None:
        self.live, self.user = live, user
        self._tokens: list[Token] = []

    async def __aenter__(self) -> Acting:
        self._tokens.append(current_token.set(self.live.tokens[self.user]["access_token"]))
        self._tokens.append(current_areas.set(None))
        return self

    async def __aexit__(self, *exc: object) -> None:
        current_areas.reset(self._tokens.pop())
        current_token.reset(self._tokens.pop())


class Live:
    def __init__(self, tokens: dict, findings_path: Path) -> None:
        self.tokens = tokens
        self.findings_path = findings_path
        self.run = f"m365-live-{uuid.uuid4().hex[:8]}"
        self.tag = self.run.replace("-", "")  # one word, for searches that split on hyphens
        self.state: dict[str, Any] = {}
        self.cleanup_failures: list[str] = []
        self._cleanup: list[tuple[str, Callable[[], Awaitable[Any]]]] = []
        self._addresses: dict[str, str] = {}
        sign, enc = os.urandom(32), os.urandom(32)
        self.settings = Settings(
            areas=frozenset(a for a, on in AREAS.items() if on),
            switches=frozenset({"teams_channels", "transcripts", "rooms"}),
            sign_keys=(sign,), enc_keys=(enc,), links_base_url=LINKS_BASE)
        # The bucket's clock runs 60 times faster, so a run is not throttled by the limit that unit tests
        # already cover. Everything else is the production runtime: real network, real sleep.
        self.rt = Runtime(settings=self.settings, signer=LinkSigner((sign,), (enc,), LinkStore()), cursors=Cursors(sign),
                          guard=Guard(clock=lambda: time.monotonic() * 60), transport=None, sleep=None)
        self.server = M365Server(self.settings)
        self.install()

    def install(self) -> None:
        configure(self.rt)

    def acting(self, user: str) -> Acting:
        return Acting(self, user)

    # -- tools ---------------------------------------------------------------------------------

    async def invoke(self, user: str, _tool: str, **args: Any) -> Result:
        """Call a tool as `user` the way the MCP handler does: (is_error, text, structured data)."""
        params = CallToolRequestParams(name=_tool, arguments=args)
        async with self.acting(user):
            res = await self.server._handle_call_tool(None, params)
        text = "".join(getattr(c, "text", "") for c in res.content)
        data = res.structured_content
        if isinstance(data, dict) and set(data) == {"result"}:
            data = data["result"]
        if not res.is_error:
            # As in the unit fixture: the copy Claude Code shows is the one the text carries.
            assert json.loads(text) == data, f"{_tool}: the text and the structured content differ"
        return Result(bool(res.is_error), text, data)

    async def ok(self, user: str, _tool: str, **args: Any) -> Any:
        """Like invoke, but a refusal fails the test with the sentence the person would read."""
        error, text, data = await self.invoke(user, _tool, **args)
        assert not error, f"{_tool} as {user} was refused: {text}"
        return data

    async def address(self, user: str) -> str:
        if user not in self._addresses:
            data = await self.ok(user, "get_connection")
            self._addresses[user] = data["address"]
        return self._addresses[user]

    # -- Graph, directly (for what the tools deliberately do not do) ---------------------------

    async def refresh_with(self, user: str, scopes: list[str]) -> str:
        """An access token from `user`'s refresh token, asking for only `scopes`. The token file is not
        changed. The client id is the one the current token was issued to."""
        entry = self.tokens[user]
        tok = claims(entry["access_token"])
        async with httpx.AsyncClient(timeout=60) as client:
            resp = await client.post(f"https://login.microsoftonline.com/{tok['tid']}/oauth2/v2.0/token", data={
                "client_id": tok.get("appid") or tok["azp"], "grant_type": "refresh_token",
                "refresh_token": entry["refresh_token"],
                "scope": " ".join(f"https://graph.microsoft.com/{s}" for s in scopes)})
        assert resp.status_code == 200, f"refresh asking for {scopes}: {resp.status_code}"
        return resp.json()["access_token"]

    async def raw(self, user: str, method: str, path: str, *, prefer: str | None = None, **kw: Any) -> httpx.Response:
        headers = {"Authorization": f"Bearer {self.tokens[user]['access_token']}", "Accept": "application/json",
                   "Prefer": 'IdType="ImmutableId"' + (f", {prefer}" if prefer else "")}
        headers.update(kw.pop("headers", {}))
        url = path if path.startswith("https://") else GRAPH + path
        async with httpx.AsyncClient(timeout=httpx.Timeout(120.0, connect=15.0), follow_redirects=False) as client:
            return await client.request(method, url, headers=headers, **kw)

    @staticmethod
    def error_of(resp: httpx.Response) -> str:
        """`status code: message` of a Graph failure, for the findings report."""
        try:
            err = resp.json().get("error", {})
            return f"{resp.status_code} {err.get('code', '')}: {str(err.get('message', ''))[:200]}"
        except ValueError:
            return str(resp.status_code)

    # -- links, in process -----------------------------------------------------------------------

    @staticmethod
    def link_in(text: str) -> str:
        """The links-app URL inside a curl command or any text."""
        m = _LINK.search(text)
        assert m, f"no link in: {text[:200]}"
        return m.group(0)

    @staticmethod
    def _path(url: str) -> str:
        assert url.startswith(LINKS_BASE), url
        return url[len(LINKS_BASE):]

    @staticmethod
    def _client() -> TestClient:
        return TestClient(links_app(), raise_server_exceptions=False)

    async def fetch(self, url: str, headers: dict[str, str] | None = None) -> httpx.Response:
        """GET a download link through the real links app."""
        return await asyncio.to_thread(lambda: self._client().get(self._path(url), headers=headers or {}))

    async def upload(self, url: str, data: bytes, headers: dict[str, str] | None = None) -> httpx.Response:
        """PUT bytes to an upload link through the real links app."""
        return await asyncio.to_thread(lambda: self._client().put(self._path(url), content=data, headers=headers or {}))

    # -- waiting ---------------------------------------------------------------------------------

    @staticmethod
    async def poll(check: Callable[[], Awaitable[Any]], *, timeout: float, every: float = 3.0, what: str = "") -> Any:
        """Await `check()` until it returns something truthy; fail with `what` when time runs out."""
        deadline = time.monotonic() + timeout
        while True:
            got = await check()
            if got:
                return got
            if time.monotonic() + every > deadline:
                raise AssertionError(f"not seen within {int(timeout)} s: {what}")
            await asyncio.sleep(every)

    # -- findings --------------------------------------------------------------------------------

    def finding(self, name: str, observed: Any) -> None:
        """Record what a live measurement saw: name -> {observed, date} in the findings report."""
        report: dict[str, Any] = {}
        try:
            report = json.loads(self.findings_path.read_text())
        except (OSError, ValueError):
            pass
        report[name] = {"observed": observed, "date": dt.date.today().isoformat()}
        self.findings_path.write_text(json.dumps(report, indent=2, sort_keys=True, default=str) + "\n")

    # -- shared scenario setup -------------------------------------------------------------------

    async def files_folder(self) -> str:
        """A folder in A's OneDrive named for the run; deleted with everything in it at the end."""
        if "folder" not in self.state:
            made = await self.ok("a", "update_file", action="create_folder", new_name=self.run)
            self.later_tool("a", "update_file", action="delete", ids=[made["id"]], user_confirmed=True)
            self.state["folder"] = self.run
        return self.state["folder"]

    async def put_file(self, folder: str, name: str, data: bytes) -> dict:
        """Upload through the upload_file link, in process; returns the link route's JSON."""
        plan = await self.ok("a", "upload_file", folder=folder, name=name)
        res = await self.upload(self.link_in(plan["upload"]), data, {"Content-Type": "application/octet-stream"})
        assert res.status_code == 200, res.text[:300]
        assert res.json()["size"] == len(data)
        return res.json()

    async def notes_target(self) -> tuple[str, str]:
        """(notebook, section) to write in: the first notebook with a section, or one created for the run."""
        if "notes" in self.state:
            return self.state["notes"]
        for item in (await self.ok("a", "list_notes"))["items"]:
            if item.get("notebook") and item.get("sections"):
                self.state["notes"] = (item["notebook"], item["sections"][0])
                return self.state["notes"]
        from m365_mcp.graph.client import Graph
        from m365_mcp.identity import current_person

        name = f"{self.run} notebook"
        async with self.acting("a"):
            person = current_person()
            async with Graph(person.token, person.key) as graph:
                book = await graph.post("/me/onenote/notebooks", {"displayName": name})
                await graph.post(f"/me/onenote/notebooks/{book['id']}/sections", {"displayName": "Live"})

        async def ready() -> bool:
            error, _, data = await self.invoke("a", "list_notes", notebook=name)
            return not error and data["count"] > 0
        await self.poll(ready, timeout=60, what="the new notebook and section to be listed")

        async def left_behind() -> None:
            raise RuntimeError(f"notebook '{name}' cannot be deleted through Graph; delete it in OneNote")
        self.later("notebook", left_behind)
        self.state["notes"] = (name, "Live")
        return self.state["notes"]

    # -- cleanup ---------------------------------------------------------------------------------

    def later(self, label: str, action: Callable[[], Awaitable[Any]]) -> None:
        self._cleanup.append((label, action))

    def later_tool(self, user: str, _tool: str, **args: Any) -> None:
        async def run() -> None:
            error, text, _ = await self.invoke(user, _tool, **args)
            # Already removed by the scenario itself is the normal case, not a failure.
            if error and not text.startswith(("It does not exist", "No ", "This contact was deleted")):
                raise RuntimeError(text[:200])
        self.later(f"{_tool} as {user}", run)

    async def _sweep_mail(self, user: str) -> None:
        cursor = None
        for _ in range(4):
            args: dict[str, Any] = {"query": self.run, "limit": 25}
            if cursor:
                args = {"query": self.run, "limit": 25, "cursor": cursor}
            error, text, data = await self.invoke(user, "list_messages", **args)
            if error:
                raise RuntimeError(text[:200])
            ids = [m["id"] for m in data["items"]]
            if ids:
                error, text, _ = await self.invoke(user, "update_messages", ids=ids, action="delete")
                if error:
                    raise RuntimeError(text[:200])
            cursor = data.get("cursor")
            if not cursor:
                break

    async def _sweep_events(self, user: str) -> None:
        r = await self.raw(user, "GET", "/me/events", params={
            "$filter": f"contains(subject,'{self.run}')", "$select": "id", "$top": "50"})
        if r.status_code != 200:
            raise RuntimeError(self.error_of(r))
        for e in r.json().get("value", []):
            d = await self.raw(user, "DELETE", f"/me/events/{e['id']}")
            if d.status_code not in (200, 204, 404):
                raise RuntimeError(self.error_of(d))

    async def finish(self) -> None:
        """Remove what the run created. A step that fails is reported at the end, never raised."""
        steps: list[tuple[str, Callable[[], Awaitable[Any]]]] = list(reversed(self._cleanup))
        for user in ("a", "b"):
            steps.append((f"mail of {user}", lambda u=user: self._sweep_mail(u)))
            steps.append((f"events of {user}", lambda u=user: self._sweep_events(u)))
        for label, action in steps:
            try:
                await action()
            except BaseException as e:  # noqa: BLE001 - cleanup must not stop at the first failure
                if isinstance(e, (KeyboardInterrupt, SystemExit)):
                    raise
                self.cleanup_failures.append(f"{label}: {type(e).__name__}: {str(e)[:200]}")

    def close(self) -> None:
        asyncio.run(self.finish())
        if self.cleanup_failures:
            message = "live cleanup left things behind:\n  " + "\n  ".join(self.cleanup_failures)
            print(message, file=sys.stderr)
            warnings.warn(message, stacklevel=1)
        else:
            print(f"live cleanup: run {self.run} removed", file=sys.stderr)
