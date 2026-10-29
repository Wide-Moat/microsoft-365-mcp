"""Every Must refuse row names a collected test, and every such test has a guard to remove."""

import importlib.util
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]


def _load(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / "scripts" / f"{name}.py")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod  # dataclasses look their module up while the class is built
    spec.loader.exec_module(mod)
    return mod


refusal_map = _load("refusal_map")


@pytest.fixture(scope="module")
def known():
    return refusal_map.collected()


@pytest.fixture(scope="module")
def parsed():
    return refusal_map.parse(refusal_map.documents())


def test_every_must_refuse_row_names_a_collected_test(parsed, known):
    assert refusal_map.check(parsed, known) == []


def test_the_documents_hold_a_plausible_number_of_rows(parsed):
    # Derived, not chosen: every table counted by a looser match than the parser's was parsed, and
    # each holds at least one row.
    assert parsed.tables > 0 and parsed.headers_seen == parsed.tables
    assert len(parsed.rows) >= parsed.tables
    assert {r.doc for r in parsed.rows} >= {"docs/requirements.md", "docs/components/identity.md",
                                           "docs/components/guard.md"}


def test_every_mapped_test_has_exactly_one_guard_mutation(parsed):
    evidence = _load("refusal_evidence")
    mapped = set(refusal_map.mapped_tests(parsed))
    assert sorted(mapped - set(evidence.MUTATIONS)) == [], "rows whose guard has no mutation"
    assert sorted(set(evidence.MUTATIONS) - mapped) == [], "mutations for tests no row names"


# -- the checker refuses what it should ------------------------------------------------------

def _doc(tmp_path, text):
    p = tmp_path / "requirements.md"
    p.write_text(text)
    return refusal_map.parse([p], root=tmp_path)


GOOD = "| Must refuse | Told | Test |\n|---|---|---|\n| X | y | `test_a.py::test_x` |\n"


def test_a_good_table_passes(tmp_path):
    assert refusal_map.check(_doc(tmp_path, GOOD), {"test_a.py::test_x"}) == []


def test_a_row_without_a_test_is_a_problem(tmp_path):
    problems = refusal_map.check(_doc(tmp_path, GOOD + "| Y | z | |\n"), {"test_a.py::test_x"})
    assert any("no test named" in p for p in problems)


def test_a_dangling_test_id_is_a_problem(tmp_path):
    problems = refusal_map.check(_doc(tmp_path, GOOD), {"test_a.py::test_other"})
    assert any("not in the collected suite" in p for p in problems)


def test_one_test_for_two_rows_is_a_problem(tmp_path):
    problems = refusal_map.check(_doc(tmp_path, GOOD + "| Y | z | `test_a.py::test_x` |\n"), {"test_a.py::test_x"})
    assert any("one test per row" in p for p in problems)


def test_a_table_without_a_test_column_is_a_problem(tmp_path):
    problems = refusal_map.check(_doc(tmp_path, "| Must refuse | |\n|---|---|\n| X | y |\n"), set())
    assert any("without a Test column" in p for p in problems)


def test_no_tables_at_all_is_a_problem(tmp_path):
    problems = refusal_map.check(_doc(tmp_path, "# Nothing here\n"), set())
    assert any("no Must refuse rows" in p for p in problems)


def test_a_header_the_parser_cannot_read_is_a_problem(tmp_path):
    problems = refusal_map.check(_doc(tmp_path, GOOD + "\n| MUST REFUSE | Told | Test |\n|---|---|---|\n| Q | r | |\n"),
                                 {"test_a.py::test_x"})
    assert any("form the parser does not read" in p for p in problems)


def test_an_indented_table_is_read_like_any_other(tmp_path):
    indented = "\n".join("  " + line for line in "| Must refuse | Told | Test |\n|---|---|---|\n| Y | z | |".split("\n"))
    parsed = _doc(tmp_path, GOOD + "\n" + indented + "\n")
    assert parsed.tables == 2
    assert any("no test named" in p for p in refusal_map.check(parsed, {"test_a.py::test_x"}))


def test_the_evidence_command_refuses_a_broken_map(tmp_path, monkeypatch):
    evidence = _load("refusal_evidence")
    doc = tmp_path / "docs" / "requirements.md"
    doc.parent.mkdir()
    doc.write_text(GOOD + "| Y | z | |\n")
    monkeypatch.setattr(evidence, "documents", lambda: [doc])
    monkeypatch.setattr(evidence, "ROOT", tmp_path)
    monkeypatch.setattr(evidence, "collected", lambda: {"test_a.py::test_x"})
    monkeypatch.setattr(evidence, "run", lambda *a: 0)
    monkeypatch.setattr(evidence, "MUTATIONS", {})
    assert evidence.main([]) == 1
