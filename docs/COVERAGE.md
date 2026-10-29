# Test coverage

[Documentation index](README.md) · [Development](DEVELOPMENT.md) · [CI](CI.md)

CI requires **90% combined statement and branch coverage**. The threshold is in
`pyproject.toml` under `[tool.coverage.report]`. Passing it does not establish that
a particular change is tested well: a small untested addition can still leave the
overall percentage above the floor.

## Reproduce the CI measurement

After [development setup](BUILDING.md#development-and-checks), run with the
project environment's Python:

```console
python ci_pytest.py
```

The runner executes:

```console
python -m pytest --cov=. --cov-report=xml --cov-report=term-missing --junitxml=test-results.xml
```

Pytest-cov applies the configured branch measurement, omissions and 90% threshold.
The runner retains the reports and original exit code under `.ci-diagnostics/tests/`.
CI uploads these together with interpreter/runtime/dependency diagnostics.

To inspect the resulting coverage locally:

```console
python -m coverage report --sort=cover
python -m coverage html
```

Open `htmlcov/index.html` for missing lines and branch paths. These commands can
also return a failure when the measured total is below the configured threshold.

## Measurement scope

There are two scopes to distinguish:

- **CI:** `--cov=.` measures the repository's Python code, including root scripts
  such as `bootstrap.py`, `ci_runtime.py` and `main.py`, subject to omissions.
- **Bare `coverage run`:** `[tool.coverage.run] source = ["edf_bill_fetcher"]`
  selects the package. This is useful for package-only investigation but is not
  the same denominator as the CI command.

The configured omissions include tests, bytecode caches, build/dist contents and
spec files. Branch measurement and report exclusions are also configured in
`pyproject.toml`. Do not compare percentages from different scopes or platforms
as if they were the same measurement.

## Investigate a gap

1. Read the missing lines/branches and identify the behavior that should be
   exercised: successful parsing, invalid input, unavailable evidence, cancellation
   or a publication failure, for example.
2. Add a small synthetic example with an independently known expected result.
   Prefer exercising real parsing/calculation paths over mocking their results.
3. Run the focused regression, then the full CI-equivalent measurement when
   verifying the overall gate.

For a focused coverage view without imposing the full-suite threshold:

```console
python -m pytest tests/test_excel_review.py --cov=. --cov-report=term-missing --cov-fail-under=0
```

The override is for a deliberately incomplete local subset only. It does not
replace `python ci_pytest.py` or change the CI requirement. Coverage runs normally
replace the existing local measurement, so run the full suite again before
reporting a full-suite percentage.

If the full gate fails, inspect the missing behavior and execution reports.
Changes in exercised paths, dependencies or platform behavior can matter as well
as newly added code. Do not assume that merely adding more assertions improves
coverage or that covering a line verifies its financial meaning.

## Coverage exclusions

Every new `# pragma: no cover` must have a concise justification and be reviewed.
Use it only for a path that cannot reasonably execute under the supported test
contract. A type-checking-only block can qualify; an inconvenient error path does
not.

Before excluding a branch, check whether synthetic input, a temporary file,
fault injection or a boundary mock can exercise it. Missing dependencies and
filesystem errors are often testable. An OS-specific branch is not automatically
unreachable: CI includes Windows, macOS and Linux. A type annotation alone also
does not make invalid runtime input impossible.

Do not exclude code, skip regression tests or lower the gate to hide a failure.
Tests should assert meaningful outcomes, including preservation of earlier files
and clear failure status when an operation fails.

## Keep measurements comparable

Before comparing two percentages, record the commit, Python version, OS,
dependency set and exact command. Confirm that both are full-suite runs with the
same source/omit configuration and branch measurement. JUnit skip/failure counts
matter: an interrupted suite can report a valid-looking percentage over only the
paths it reached.

Use `COVERAGE_FILE` if you need to preserve a full-run measurement while exploring
a subset. For example, Bash accepts
`COVERAGE_FILE=.coverage-focused python -m pytest tests/test_excel_review.py --cov=. --cov-fail-under=0`.
In PowerShell set `$env:COVERAGE_FILE = ".coverage-focused"` before that command,
then remove the temporary variable when finished. The report command must use
the same variable to inspect that file. Do not accidentally publish the focused
measurement as the CI-equivalent result.

When choosing the next test, prefer an uncovered financial sign/identity case,
source failure or cancellation path over assertions that merely mirror a helper's
implementation. A high percentage can still miss double counting, stale review
bindings, wrong page links and platform-specific decoding. Those need explicit
behavioral regressions as described in [development](DEVELOPMENT.md#write-a-useful-regression).

## Find the relevant regression suite

| Behavior | Tests and expected outcome |
| --- | --- |
| Report links to original evidence | `tests/test_source_link_integrity.py`: a recorded source identity never falls back to a namesake; missing, changed, unsafe or unreadable files stay unlinked without disabling other valid sources. Repeated invoice aliases of the same original remain linkable. |
| Shared Excel and desktop corrections | `tests/test_record_review_validation.py`: invalid dates, precision, transaction signs, reasons and row bindings are rejected; corrections and exclusions preserve the original record and nested provenance. |
| Excel review and regeneration | `tests/test_excel_review.py`: real workbook edits rebuild the analysis, preserve source files and review history, and remain lossless when no corrections are made. |
| Saved analysis integrity | `tests/test_saved_result_validation.py`: corrupt fields and nonexistent evidence references are rejected instead of being accepted as an earlier result. |
| Report output contracts | `tests/test_run_output_regressions.py`: synthetic examples exercise financial identity, ledger boundaries and actual Word/PDF hyperlinks. |

Use the smaller suites to diagnose a failure, then run the full measurement above.
Keep fixture amounts and identities synthetic. Private evidence is useful for a
separate local reconciliation, but must not become a committed fixture. When a
new test exposes a bug, check that it fails before the production fix and asserts
the expected outcome afterward. Adding branches to the implementation can increase
the measurement denominator; report the actual before/after result instead of
assuming a fixed percentage gain from the number of new cases.

## Historical measurement

The [2026-08-02 baseline](coverage/baseline-2026-08-02.txt) records an earlier
coverage campaign. It is not the current result. Current measurements belong to
an exact revision, command and platform; use that revision's CI reports rather
than a static number in this guide. The old generated HTML baseline is not stored
in the repository.
