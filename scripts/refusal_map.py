"""Every "Must refuse" row in the documents names the test that holds it, and that test exists.

A refusal written down without a test is a promise nobody checks. The documents' Must refuse
tables carry a Test column of pytest ids (relative to tests/unit/); this script reads every such
table in docs/requirements.md and docs/components/*.md and fails on a row without a test, an id
that is malformed, named twice, or not in the collected suite — and on finding no tables at all,
which is what a broken parser would report as success.

    python scripts/refusal_map.py          # print the map; exit 1 on any problem
"""

from __future__ import annotations

import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TEST_ROOT = "tests/unit"
# Markdown lets a table start up to three spaces in; both matches and the row loop allow it.
_HEADER = re.compile(r"^ {0,3}\|\s*Must refuse\s*\|")
_ANY_HEADER = re.compile(r"^\s*\|.*must refuse", re.I)  # deliberately looser than the parser
_ID = re.compile(r"^test_[a-z0-9_]+\.py::test_[a-z0-9_]+$")


def documents(root: Path = ROOT) -> list[Path]:
    return [root / "docs" / "requirements.md", *sorted((root / "docs" / "components").glob("*.md"))]


@dataclass(frozen=True)
class Row:
    doc: str
    line: int
    refusal: str
    tests: tuple[str, ...]


@dataclass
class Parsed:
    rows: list[Row]
    tables: int
    headers_seen: int  # every table line that mentions "must refuse", however it is written
    problems: list[str]


def _cells(line: str) -> list[str]:
    return [c.strip() for c in line.strip().strip("|").split("|")]


def parse(paths: list[Path], root: Path = ROOT) -> Parsed:
    out = Parsed(rows=[], tables=0, headers_seen=0, problems=[])
    for path in paths:
        rel = str(path.relative_to(root)) if path.is_relative_to(root) else str(path)
        lines = path.read_text().splitlines()
        i = 0
        while i < len(lines):
            line = lines[i]
            if _ANY_HEADER.match(line):
                out.headers_seen += 1
            if not _HEADER.match(line):
                i += 1
                continue
            out.tables += 1
            header = _cells(line)
            col = header.index("Test") if "Test" in header else None
            if col is None:
                out.problems.append(f"{rel}:{i + 1}: a Must refuse table without a Test column")
            i += 2  # the header and its |---| separator
            rows_here = 0
            while i < len(lines) and lines[i].lstrip(" ").startswith("|"):
                cells = _cells(lines[i])
                rows_here += 1
                cell = cells[col] if col is not None and col < len(cells) else ""
                ids = tuple(re.findall(r"`([^`]+)`", cell))
                out.rows.append(Row(rel, i + 1, cells[0], ids))
                i += 1
            if rows_here == 0:
                out.problems.append(f"{rel}:{i}: a Must refuse table with no rows")
    return out


def collected(root: Path = ROOT) -> set[str]:
    """Test ids pytest collects under tests/unit, without parameters, relative to tests/unit."""
    proc = subprocess.run([sys.executable, "-m", "pytest", "--collect-only", "-q", "-p", "no:cacheprovider", TEST_ROOT],
                          cwd=root, capture_output=True, text=True)
    ids = {line.split("[", 1)[0].removeprefix(TEST_ROOT + "/") for line in proc.stdout.splitlines() if "::" in line}
    if proc.returncode != 0 or not ids:
        raise RuntimeError(f"pytest collected nothing usable (exit {proc.returncode})")
    return ids


def check(parsed: Parsed, known: set[str]) -> list[str]:
    problems = list(parsed.problems)
    if parsed.tables == 0 or not parsed.rows:
        problems.append("no Must refuse rows were found: the documents moved or the parser is broken")
    if parsed.headers_seen != parsed.tables:
        problems.append(f"{parsed.headers_seen} table lines mention Must refuse but {parsed.tables} were parsed: "
                        "a header is written in a form the parser does not read")
    owner: dict[str, Row] = {}
    for r in parsed.rows:
        where = f"{r.doc}:{r.line} ({r.refusal[:60]})"
        if not r.tests:
            problems.append(f"{where}: no test named")
        for t in r.tests:
            if not _ID.match(t):
                problems.append(f"{where}: {t!r} is not a test id like test_mail.py::test_name")
            elif t not in known:
                problems.append(f"{where}: {t} is not in the collected suite")
            if t in owner and owner[t] != r:
                problems.append(f"{where}: {t} already holds {owner[t].doc}:{owner[t].line}; one test per row")
            owner.setdefault(t, r)
    return problems


def mapped_tests(parsed: Parsed) -> list[str]:
    return [t for r in parsed.rows for t in r.tests]


def main() -> int:
    parsed = parse(documents())
    problems = check(parsed, collected())
    for r in parsed.rows:
        print(f"{r.doc}:{r.line}  {r.refusal[:70]}")
        for t in r.tests:
            print(f"    {t}")
    docs = len({r.doc for r in parsed.rows})
    print(f"\n{len(parsed.rows)} rows in {parsed.tables} tables across {docs} documents, "
          f"{len(mapped_tests(parsed))} tests, {len(problems)} problems")
    for p in problems:
        print(" -", p)
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
