"""Every manifesto rule names a holder that exists, and the checker refuses the ways a map goes wrong."""

import importlib.util
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]


def _load(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / "scripts" / f"{name}.py")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


manifesto_map = _load("manifesto_map")

RULES = "".join(f"{n}. **Rule {n}.** Text.\n" for n in range(1, 11))
TESTS = {"test_a.py::test_one", "test_a.py::test_two"}
CI = {"purity", "the tests pass"}


def _table(rows):
    return "| Rule | Held by |\n|---|---|\n" + "".join(f"| {r} | {h} |\n" for r, h in rows)


def _check(rows, rules=RULES):
    held, problems = manifesto_map.holders(_table(rows))
    return problems + manifesto_map.check(manifesto_map.rules(rules), held, TESTS, CI)


GOOD = [(n, "`test_a.py::test_one`") for n in range(1, 10)] + [(10, "`ci: purity`")]


def test_every_rule_names_a_collected_holder():
    numbers = manifesto_map.rules((ROOT / "MANIFESTO.md").read_text())
    held, problems = manifesto_map.holders((ROOT / "docs" / "manifesto-map.md").read_text())
    known_ci = manifesto_map.ci_names((ROOT / ".github" / "workflows" / "validate.yml").read_text())
    assert len(numbers) >= 40 and len(known_ci) >= 10
    assert problems + manifesto_map.check(numbers, held, manifesto_map.collected(), known_ci) == []


def test_a_good_map_passes():
    assert _check(GOOD) == []


def test_a_rule_missing_from_the_map_is_a_problem():
    assert any("rule 10: not in" in p for p in _check(GOOD[:-1]))


def test_a_rule_with_no_holder_is_a_problem():
    assert any("rule 1: no holder" in p for p in _check([(1, "nothing")] + GOOD[1:]))


def test_a_rule_listed_twice_is_a_problem():
    assert any("listed twice" in p for p in _check(GOOD + [(3, "`test_a.py::test_two`")]))


def test_a_rule_the_manifesto_does_not_have_is_a_problem():
    assert any("rule 11: the manifesto has no such rule" in p for p in _check(GOOD + [(11, "`test_a.py::test_one`")]))


@pytest.mark.parametrize("holder, needle", [
    ("`test_a.py::test_gone`", "not in the collected suite"),
    ("`tests/unit/test_a.py::test_one`", "not a test id"),
    ("`ci: a step nobody wrote`", "no job or step"),
])
def test_a_holder_that_does_not_exist_is_a_problem(holder, needle):
    assert any(needle in p for p in _check([(1, holder)] + GOOD[1:]))


def test_reading_no_rules_is_a_problem():
    assert any("parser is broken" in p for p in _check([], rules="# Manifesto\n\nNo numbered rules.\n"))


def test_ci_names_are_jobs_and_steps_not_events():
    workflow = ("name: validate\non:\n  push:\n    branches: [main]\n  pull_request:\n"
                "jobs:\n  purity:\n    uses: x\n  tests:\n    steps:\n      - name: the tests pass\n        run: pytest\n")
    assert manifesto_map.ci_names(workflow) == {"purity", "tests", "the tests pass"}
