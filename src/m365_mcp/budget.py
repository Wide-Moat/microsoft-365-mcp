"""What the catalogue costs the model, measured on the exact tools/list payload.

    python -m m365_mcp.budget            # table for the default configuration
    python -m m365_mcp.budget --all      # every area and switch on
    python -m m365_mcp.budget --check    # exit 1 above the manifesto's limits
"""

from __future__ import annotations

import argparse
import asyncio
import json
import os
import sys
from pathlib import Path

LIMIT_TOTAL = 16_000
LIMIT_AREA = 3_500
LIMIT_TOOL = 600


def _encoder():
    import tiktoken

    cache = Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "tiktoken"
    if cache.is_dir():
        os.environ.setdefault("TIKTOKEN_CACHE_DIR", str(cache))
    return tiktoken.get_encoding("o200k_base")


def catalogue_server(all_on: bool = False):
    """The server as an installation with the default areas (or every area and switch) would publish it."""
    from m365_mcp.config import SWITCHES, Settings
    from m365_mcp.core.context import Runtime, configure
    from m365_mcp.core.cursor import Cursors
    from m365_mcp.registry import AREAS
    from m365_mcp.server import M365Server

    areas = frozenset(AREAS) if all_on else frozenset(a for a, on in AREAS.items() if on)
    switches = frozenset(SWITCHES) if all_on else frozenset(s for s, on in SWITCHES.items() if on)
    settings = Settings(areas=areas, switches=switches, sign_keys=(b"k" * 32,), enc_keys=(b"e" * 32,),
                        links_base_url="https://links.invalid")
    configure(Runtime(settings=settings, signer=None, cursors=Cursors(None)))
    return M365Server(settings)


async def measure(all_on: bool = False) -> dict:
    server = catalogue_server(all_on)
    enc = _encoder()
    rows = []
    for t in await server.list_tools():
        payload = json.dumps(t.model_dump(by_alias=True, exclude_none=True, mode="json"), separators=(",", ":"))
        rows.append({"tool": t.name, "area": server.specs[t.name].area, "tokens": len(enc.encode(payload))})
    # The total is the whole tools/list payload encoded at once, as the model receives it, not a sum.
    payload_all = json.dumps([t.model_dump(by_alias=True, exclude_none=True, mode="json") for t in await server.list_tools()],
                             separators=(",", ":"))
    by_area: dict[str, int] = {}
    for r in rows:
        by_area[r["area"]] = by_area.get(r["area"], 0) + r["tokens"]
    return {"tools": rows, "areas": by_area, "total": len(enc.encode(payload_all))}


def violations(m: dict, *, total: bool = True) -> list[str]:
    out = []
    if total and m["total"] > LIMIT_TOTAL:
        out.append(f"catalogue {m['total']} > {LIMIT_TOTAL}")
    out += [f"area {a} {n} > {LIMIT_AREA}" for a, n in m["areas"].items() if n > LIMIT_AREA]
    out += [f"tool {r['tool']} {r['tokens']} > {LIMIT_TOOL}" for r in m["tools"] if r["tokens"] > LIMIT_TOOL]
    return out


def table(m: dict) -> str:
    lines = ["| Area | Tool | Tokens |", "|---|---|---|"]
    for r in sorted(m["tools"], key=lambda r: (r["area"], r["tool"])):
        lines.append(f"| {r['area']} | `{r['tool']}` | {r['tokens']} |")
    lines += ["", "| Area | Tokens |", "|---|---|"]
    lines += [f"| {a} | {n} |" for a, n in sorted(m["areas"].items())]
    lines += ["", f"**Total: {m['total']} tokens, {len(m['tools'])} tools** "
                  f"(limits: {LIMIT_TOTAL} total, {LIMIT_AREA} per area, {LIMIT_TOOL} per tool)"]
    return "\n".join(lines)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--check", action="store_true", help="fail above any limit")
    ap.add_argument("--check-tools", action="store_true",
                    help="fail above the per-area and per-tool limits only (the total limit is for the default set)")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()
    m = asyncio.run(measure(all_on=a.all))
    print(json.dumps(m) if a.json else table(m))
    bad = violations(m, total=not a.check_tools)
    if bad:
        print("\nOVER BUDGET: " + "; ".join(bad), file=sys.stderr)
    return 1 if ((a.check or a.check_tools) and bad) else 0


if __name__ == "__main__":
    sys.exit(main())
