# Working in m365-mcp

Read MANIFESTO.md before changing anything; docs/code-review.md lists what blocks a merge.

## Things that break silently

- **Logs carry content unless stopped.** httpx logs every request URL at INFO, and Graph URLs hold
  search strings and addresses; `graph/client.py` silences it. The SDK logs `str()` of a tool
  error; `M365Error.__str__` is the key alone and the sentence goes to the model separately
  (`server.py`). `tests/unit/test_core.py::test_a_full_mail_flow_logs_no_content` guards both.
- **A guard that two checks hold is invisible to a one-line mutation.** `scripts/mutation_check.py`
  removes each safety rule and requires a test to fail; when a rule is held twice, mutate both.
- **The budget limit is for the default set.** `budget --check` (default areas) enforces the total;
  `budget --all --check-tools` enforces per-area and per-tool limits with everything on.
- **`docs/tools.md` is generated.** Edit descriptions in code, then `python -m m365_mcp.docs > docs/tools.md`.
- **`tool-surface.lock.json` is generated too.** A changed argument, enum or annotation fails
  `tests/unit/test_surface.py` until `python -m m365_mcp.surface > tool-surface.lock.json` is run and
  the lock committed, so the change is visible in review.
- **A Must refuse row needs two things besides its test:** the test id in the row's Test column
  (`scripts/refusal_map.py`), and a mutation removing its guard in `scripts/refusal_evidence.py`,
  which requires that test to fail. A test that asserts only `isError` usually fails this: with the
  guard gone it errors later, for an unrelated reason, and passes.
- **Graph ids:** every request sends `Prefer: IdType="ImmutableId"`; never build a request that
  bypasses `Graph._headers`.

## Process

Test first; every refusal in docs/requirements.md has a test asserting `isError`, the text, and that
nothing was written (`graph.writes() == []`). English in the repository.
