"""Run inside a server container (kubectl exec -i … python - <args> < scripts/e2e-probe.py) by chart-e2e.sh.

It uses the product's own code and the container's own settings, key file and store: what it does is what
the server process does for a tool call, without a Microsoft 365 sign-in. Each command prints one line.

    link                 issue a download link the way a tool does; prints its id
    upload-link          issue an upload link; prints its id
    claim <id>           start an upload on a link, as the links route does; prints True or False
    guard <person> <n>   draw n sends from a person's allowance; prints each outcome (ok, rate, state_unavailable)
    get <url>            GET a URL; prints the status and the start of the body
    put <url>            PUT a small body to a URL; prints the status and the start of the body

It also runs against an image from before the shared store (0.1.x), to show the checks fail there: that
code keeps its links and its allowance in the memory of the process that made them.
"""

import asyncio
import base64
import json
import sys
import time
import urllib.error
import urllib.request


def forged_token() -> tuple[str, int]:
    """A token shaped like Graph's. Graph rejects it (it is not signed), which is the point: a link that
    reaches Graph with it was found, opened and verified by the replica that served it."""
    exp = int(time.time()) + 3600

    def part(d: dict) -> str:
        return base64.urlsafe_b64encode(json.dumps(d).encode()).rstrip(b"=").decode()

    claims = {"aud": "https://graph.microsoft.com", "tid": "e2e-tenant", "oid": "e2e-person", "exp": exp, "scp": ""}
    return f"{part({'alg': 'none'})}.{part(claims)}.x", exp


async def runtime():
    from m365_mcp.config import Settings

    settings = Settings.from_env()
    try:
        from m365_mcp.app import make_signer, make_store
        from m365_mcp.core.guard import Guard
    except ImportError:  # 0.1.x: everything in this process's memory
        from m365_mcp.core.guard import GUARD
        from m365_mcp.links.store import LinkStore
        from m365_mcp.links.tokens import LinkSigner

        return settings, LinkSigner(settings.sign_keys, settings.enc_keys, LinkStore()), GUARD, None
    store = make_store(settings)
    return settings, make_signer(settings, store), Guard(store), store


async def maybe(value):
    return await value if asyncio.iscoroutine(value) else value


def http(method: str, url: str) -> str:
    req = urllib.request.Request(
        url,
        method=method,
        data=b"e2e" if method == "PUT" else None,
        headers={"Content-Type": "application/octet-stream", "X-File-Name": "e2e.txt"},
    )
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            return f"{r.status} {r.read(200)!r}"
    except urllib.error.HTTPError as e:
        return f"{e.code} {e.read(200)!r}"


async def main(args: list[str]) -> None:
    if args[0] in ("get", "put"):
        print(http(args[0].upper(), args[1]))
        return
    settings, signer, guard, store = await runtime()
    try:
        token, exp = forged_token()
        from m365_mcp.identity import person_from_token

        person = person_from_token(token).key
        if args[0] in ("link", "upload-link"):
            op, kind, target = (
                ("d", "drive", {"i": "e2e"})
                if args[0] == "link"
                else ("u", "drive_upload", {"p": "e2e", "n": "e2e.txt", "c": "fail"})
            )
            print(
                await maybe(
                    signer.issue(
                        op=op,
                        kind=kind,
                        target=target,
                        person=person,
                        token=token,
                        token_expires=exp,
                        ttl=300 if op == "d" else 180,
                    )
                )
            )
        elif args[0] == "claim":
            print(await maybe(signer.claim(args[1])))
        elif args[0] == "guard":
            from m365_mcp.core.errors import M365Error

            out = []
            for _ in range(int(args[2])):
                try:
                    await maybe(guard.check(args[1], kind="mail", user_confirmed=True, recipients=1))
                    out.append("ok")
                except M365Error as e:
                    out.append(e.key)
            print(" ".join(out))
        else:
            raise SystemExit(f"unknown command {args[0]}")
    finally:
        if store is not None:
            await store.aclose()


asyncio.run(main(sys.argv[1:]))
