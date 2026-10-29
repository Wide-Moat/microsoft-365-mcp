"""Manifesto rules that no tool-level test held: each test here reads the whole repository or the
whole published catalogue, so a new file or a new tool that breaks the rule fails here."""

import json
import re
import shutil
import subprocess
import sys
import unicodedata
from pathlib import Path

from m365_mcp.registry import all_tools
from m365_mcp.tools.account import AREA_SCOPES

ROOT = Path(__file__).resolve().parents[2]
SRC = ROOT / "src" / "m365_mcp"
SPECS = all_tools()


def _walk(node, path=""):
    if isinstance(node, dict):
        yield path, node
        for k, v in node.items():
            yield from _walk(v, f"{path}/{k}")
    elif isinstance(node, list):
        for i, v in enumerate(node):
            yield from _walk(v, f"{path}/{i}")


async def test_actions_are_closed_enums_inline_in_the_schema(monkeypatch):
    # Rule 3: an `action: str` lets the model invent verbs; an enum behind a $ref is one it cannot see.
    from m365_mcp.budget import catalogue_server
    from m365_mcp.core import context

    monkeypatch.setattr(context, "_RUNTIME", context._RUNTIME)  # catalogue_server configures a stub runtime
    tools = await catalogue_server(all_on=True).list_tools()
    actions = 0
    for t in tools:
        schema = t.model_dump(by_alias=True, exclude_none=True, mode="json")["inputSchema"]
        for where, node in _walk(schema):
            assert "$ref" not in node and "$defs" not in node, f"{t.name}{where} hides a type behind a reference"
        action = schema.get("properties", {}).get("action")
        if action is not None:
            actions += 1
            variants = action.get("anyOf", [action])
            assert any(v.get("enum") for v in variants), f"{t.name}.action is not a closed enum"
            assert all(v.get("enum") or v.get("type") == "null" for v in variants), f"{t.name}.action accepts free text"
    assert actions >= 5, f"only {actions} tools take an action: the schemas are not what this test thinks"


def test_no_app_only_credentials_or_token_store():
    # Rule 15: the person's own bearer, forwarded by the gateway, is the only identity this server has.
    banned = ["client_credentials", "grant_type", "oauth2/v2.0/token", "msal", "ConfidentialClient", "client_secret"]
    files = list(SRC.rglob("*.py"))
    assert len(files) > 20
    for path in files:
        src = path.read_text()
        for b in banned:
            assert b not in src, f"{path.relative_to(ROOT)} mentions {b}"


def test_every_area_scope_is_needed_by_a_tool_of_that_area():
    # Rule 18: get_connection tells the person which permissions an area needs; one that no tool of
    # the area declares is a scope asked for without a tool that needs it.
    for area, scopes in AREA_SCOPES.items():
        needed = {s for spec in SPECS.values()
                  if spec.area == area or ("area", area) in spec.gates.values() for s in spec.scopes}
        assert needed, f"area {area} has no tool"
        assert set(scopes) <= needed, f"area {area} lists {sorted(set(scopes) - needed)} that no tool of it needs"


def _tracked_text_files() -> list[Path]:
    if shutil.which("git") and (ROOT / ".git").exists():
        out = subprocess.run(["git", "ls-files", "-z"], cwd=ROOT, capture_output=True, check=True).stdout
        paths = [ROOT / p for p in out.decode().split("\0") if p]
    else:
        skip = {".git", ".venv", "venv", "__pycache__", "build", "dist", "node_modules", ".pytest_cache"}
        paths = [p for p in ROOT.rglob("*") if p.is_file() and not skip & set(p.relative_to(ROOT).parts)]
    texts = []
    for p in paths:
        if not p.is_file():
            continue
        try:
            p.read_text(encoding="utf-8")
        except (UnicodeDecodeError, OSError):
            continue
        texts.append(p)
    return texts


# Scripts a passage in another language is written in. Greek stays out: a symbol such as Ω or µ in a
# unit is not a passage; signs such as º have no script name at all.
_SCRIPTS = ("CYRILLIC", "ARABIC", "HEBREW", "CJK", "HIRAGANA", "KATAKANA", "HANGUL", "THAI", "DEVANAGARI",
            "ARMENIAN", "GEORGIAN", "BENGALI", "TAMIL")


def test_the_repository_is_written_in_english():
    # Rule 35: a letter from another language's script in code, docs or tests is a passage not in English.
    files = _tracked_text_files()
    assert len(files) > 50, f"{len(files)} text files found: the walk is broken"
    for p in files:
        for n, line in enumerate(p.read_text(encoding="utf-8").splitlines(), 1):
            for ch in line:
                if unicodedata.category(ch).startswith("L") and unicodedata.name(ch, "").startswith(_SCRIPTS):
                    raise AssertionError(f"{p.relative_to(ROOT)}:{n}: {unicodedata.name(ch)}")


def test_every_live_finding_is_measured_by_a_probe():
    # Rule 40: a finding is what the tenant answered, so each one in the catalogue is written by a
    # live probe; the only exception is one whose entry says why there is no probe.
    sys.path.insert(0, str(ROOT / "scripts"))
    try:
        from live_findings_to_md import CATALOGUE
    finally:
        sys.path.remove(str(ROOT / "scripts"))
    probes = "".join(p.read_text() for p in (ROOT / "tests" / "live").glob("*.py"))
    recorded = set(re.findall(r'finding\(\s*"([a-z0-9_]+)"', probes))
    assert len(CATALOGUE) >= 20 and recorded
    for name, (_, does) in CATALOGUE.items():
        assert name in recorded or "No probe:" in does, f"finding {name} has no live probe and says no reason"
    assert recorded <= set(CATALOGUE), f"probes record {sorted(recorded - set(CATALOGUE))} the document never lists"
    doc = (ROOT / "docs" / "graph-live-findings.md").read_text()
    for name in CATALOGUE:
        assert f"| `{name}` |" in doc, f"docs/graph-live-findings.md has no row for {name}"


def test_the_review_checklist_is_wired_to_the_manifesto():
    # Rule 42: review runs against docs/code-review.md through .opencodereview/rule.json; each blocking
    # item names a manifesto rule that exists, so the checklist cannot drift from the rules it guards.
    rules = json.loads((ROOT / ".opencodereview" / "rule.json").read_text())["rules"]
    assert any(r["path"].startswith("src/") for r in rules), "no review rule covers the source"
    manifesto = (ROOT / "MANIFESTO.md").read_text()
    numbers = {int(n) for n in re.findall(r"^(\d+)\. \*\*", manifesto, re.M)}
    checklist = (ROOT / "docs" / "code-review.md").read_text()
    items = re.findall(r"^\| (B\d+) \|.*\| ([\d,– ]+) \|$", checklist, re.M)
    rows = re.findall(r"^\| B\d+ \|", checklist, re.M)
    assert len(items) >= 10 and len(items) == len(rows), f"{len(items)} of {len(rows)} blocking items parsed"
    for item, cited in items:
        for part in cited.split(","):
            lo, _, hi = part.strip().partition("–")
            for n in range(int(lo), int(hi or lo) + 1):
                assert n in numbers, f"{item} cites manifesto rule {n}, which does not exist"
