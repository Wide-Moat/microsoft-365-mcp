"""Every rule of MANIFESTO.md names what holds it, and that holder exists.

Rule 36 says every rule has a test that fails when the rule is broken. docs/manifesto-map.md says
which; this script fails on a rule the map does not list, a rule listed twice or with no holder, a
number the manifesto does not have, a test id pytest does not collect, and a `ci:` holder that is
neither a job nor a step name in .github/workflows/validate.yml — and on reading no rules at all,
which is what a broken parser would report as success.

    python scripts/manifesto_map.py          # print the map; exit 1 on any problem
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from refusal_map import _ID, collected  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
_RULE = re.compile(r"^(\d+)\. \*\*", re.M)
_ROW = re.compile(r"^\|\s*(\d+)\s*\|(.*)\|\s*$")


def rules(manifesto: str) -> list[int]:
    return [int(n) for n in _RULE.findall(manifesto)]


def holders(table: str) -> tuple[dict[int, list[str]], list[str]]:
    out: dict[int, list[str]] = {}
    problems = []
    for n, line in enumerate(table.splitlines(), 1):
        m = _ROW.match(line)
        if not m:
            continue
        rule = int(m.group(1))
        if rule in out:
            problems.append(f"docs/manifesto-map.md:{n}: rule {rule} listed twice")
        out[rule] = re.findall(r"`([^`]+)`", m.group(2))
    return out, problems


def ci_names(workflow: str) -> set[str]:
    # Job keys only: the block under `on:` has the same indentation and names events, not checks.
    block = workflow.split("\njobs:\n", 1)[1] if "\njobs:\n" in workflow else ""
    jobs = set(re.findall(r"^  ([a-z][a-z0-9_-]*):\s*$", block, re.M))
    steps = {s.strip().strip("\"'") for s in re.findall(r"^\s+- name:\s*(.+)$", workflow, re.M)}
    return jobs | steps


def check(numbers: list[int], held: dict[int, list[str]], known_tests: set[str], known_ci: set[str]) -> list[str]:
    problems = []
    if len(numbers) < 10:
        problems.append(f"{len(numbers)} rules read from MANIFESTO.md: the parser is broken")
    if numbers != list(range(1, len(numbers) + 1)):
        problems.append("MANIFESTO.md rules are not numbered 1..N in order")
    for rule in numbers:
        if rule not in held:
            problems.append(f"rule {rule}: not in docs/manifesto-map.md")
        elif not held[rule]:
            problems.append(f"rule {rule}: no holder named")
    for rule, names in held.items():
        if rule not in numbers:
            problems.append(f"rule {rule}: the manifesto has no such rule")
        for h in names:
            if h.startswith("ci:"):
                if h[3:].strip() not in known_ci:
                    problems.append(f"rule {rule}: {h!r} is no job or step in validate.yml")
            elif not _ID.match(h):
                problems.append(f"rule {rule}: {h!r} is not a test id like test_mail.py::test_name or ci: <step>")
            elif h not in known_tests:
                problems.append(f"rule {rule}: {h} is not in the collected suite")
    return problems


def main() -> int:
    numbers = rules((ROOT / "MANIFESTO.md").read_text())
    held, problems = holders((ROOT / "docs" / "manifesto-map.md").read_text())
    known_ci = ci_names((ROOT / ".github" / "workflows" / "validate.yml").read_text())
    problems += check(numbers, held, collected(), known_ci)
    for rule in numbers:
        print(f"{rule:>3}  " + "  ".join(held.get(rule, [])))
    tests = sum(1 for names in held.values() for h in names if not h.startswith("ci:"))
    print(f"\n{len(numbers)} rules, {len(held)} mapped, {tests} tests, {len(problems)} problems")
    for p in problems:
        print(" -", p)
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
