"""The tool surface, pinned: every tool's name, area, switch, gated parameters, annotations and a hash
of its input schema, with every area and switch on. The lock is committed, so any change to what the model is
offered — a renamed argument, a new enum value, a flipped hint — is a visible change to one file in
review instead of a line buried in a diff.

    python -m m365_mcp.surface > tool-surface.lock.json   # regenerate after an intended change
    python -m m365_mcp.surface --check                    # exit 1 when the registry drifted
"""

from __future__ import annotations

import argparse
import asyncio
import hashlib
import json
import sys
from pathlib import Path

LOCK = Path(__file__).resolve().parents[2] / "tool-surface.lock.json"
REGENERATE = ("If the change to the tool surface is intended, regenerate the lock and commit it: "
              "python -m m365_mcp.surface > tool-surface.lock.json")


def _sha(value: object) -> str:
    return hashlib.sha256(json.dumps(value, sort_keys=True, separators=(",", ":")).encode()).hexdigest()


async def render() -> dict:
    from m365_mcp.budget import catalogue_server

    server = catalogue_server(all_on=True)
    tools = {}
    for t in await server.list_tools():
        spec = server.specs[t.name]
        wire = t.model_dump(by_alias=True, exclude_none=True, mode="json")  # what tools/list sends
        tools[t.name] = {
            "area": spec.area,
            "switch": spec.switch,
            "gates": {p: f"{kind} {name}" for p, (kind, name) in sorted(spec.gates.items())},
            "annotations": wire.get("annotations", {}),
            "input_schema_sha256": _sha(wire["inputSchema"]),
        }
    return {"tools": tools}


def drift(lock: dict, now: dict) -> list[str]:
    """What differs between the committed lock and the registry, one line per tool and field."""
    old, new = lock.get("tools", {}), now.get("tools", {})
    out = [f"{n}: not in the lock" for n in sorted(set(new) - set(old))]
    out += [f"{n}: in the lock but no longer registered" for n in sorted(set(old) - set(new))]
    labels = {"area": "area", "switch": "switch", "gates": "gated parameters", "annotations": "annotations",
              "input_schema_sha256": "input schema"}
    for n in sorted(set(old) & set(new)):
        out += [f"{n}: {label} changed" for key, label in labels.items() if old[n].get(key) != new[n].get(key)]
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true", help="compare with the committed lock; exit 1 on drift")
    a = ap.parse_args()
    now = asyncio.run(render())
    if not a.check:
        print(json.dumps(now, indent=2, sort_keys=True))
        return 0
    problems = drift(json.loads(LOCK.read_text()), now)
    for p in problems:
        print(p)
    if problems:
        print(REGENERATE, file=sys.stderr)
    print(f"{len(now['tools'])} tools checked against {LOCK.name}, {len(problems)} differences")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
