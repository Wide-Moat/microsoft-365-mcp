# Contributing

This server acts with a person's own Microsoft 365 rights, so every change is held to the rules in [`MANIFESTO.md`](MANIFESTO.md); [`docs/code-review.md`](docs/code-review.md) lists
what blocks a merge.

## Before you start

Open an issue for anything larger than a small fix, so we can agree on the shape first. A new tool
must fit the token budget and the one-question-one-tool rule.

## The loop

```sh
uv venv -p 3.12 .venv && uv pip install -p .venv -r requirements-test.lock && uv pip install -p .venv --no-deps -e .
.venv/bin/pytest -q                            # unit and contract tests
.venv/bin/python -m m365_mcp.budget            # what the catalogue costs, per tool and area
.venv/bin/python scripts/mutation_check.py     # every safety rule fails a test when removed
.venv/bin/python scripts/refusal_map.py        # every Must refuse row names an existing test
.venv/bin/python scripts/refusal_evidence.py   # and that test fails with the row's guard removed
```

- **Test first.** A behaviour change comes with a test that fails without it.
- **A refusal is a row and a test.** Every "must refuse" in [`docs/requirements.md`](docs/requirements.md)
  names its test, and `scripts/refusal_evidence.py` removes the guard and requires the test to fail.
- **Regenerate generated files.** After changing a tool's description, arguments or annotations, run
  `python -m m365_mcp.docs > docs/tools.md` and `python -m m365_mcp.surface > tool-surface.lock.json`.
  Never edit them by hand.
- **No content in logs.** No subject, address or body ever reaches a log line; a test checks it.
- **English** in code, comments and documents.

CI runs every check on a pull request. The full list is in `.github/workflows/validate.yml`.

## Licence

By contributing you agree that your contribution is licensed under the
[Apache License 2.0](LICENSE).
