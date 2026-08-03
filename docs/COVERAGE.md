# Test Coverage

This document describes the test-coverage measurement protocol, the strict `# pragma: no cover` policy, the CI gate, and how to extend coverage.

## Current baseline

- **Date measured**: 2026-08-02
- **Total**: 71% (4180 of 5868 statements covered)
- **Branch coverage**: branch=true
- **Floor gate**: 90% (`coverage report --fail-under=90` in CI)
- **Aspiration target**: 100% of testable code

The point-in-time measurement is preserved at [`coverage/baseline-2026-08-02.txt`](coverage/baseline-2026-08-02.txt) (text format) and [`coverage/html-2026-08-02/`](coverage/html-2026-08-02/index.html) (HTML format with per-line highlight).

## Measurement protocol

### Running coverage locally

```bash
# Install dev extras (one-time):
pip install -e ".[dev]"

# Run the full test suite under coverage:
coverage run --branch -m pytest

# Text report to stdout:
coverage report

# HTML report at htmlcov/index.html:
coverage html

# Fail if total < 90% (CI behavior):
coverage report --fail-under=90
```

### Configuration

Coverage is configured in `pyproject.toml`:

```toml
[tool.coverage.run]
source = ["edf_bill_fetcher"]   # only the package — tests/ is omitted
branch = true                    # branch coverage in addition to statement coverage

[tool.coverage.report]
fail_under = 90                  # CI gate — fail the build below 90%
show_missing = true              # show missed line numbers per file
skip_covered = true              # hide 100%-covered files from the report
```

### What IS measured

- All `.py` files under `edf_bill_fetcher/` (the package)
- Both statement and branch coverage (branch=true)

### What is NOT measured

- `tests/` — coverage of the tests themselves isn't a goal
- `edf_collector.py` / `edf_report.py` / `edf_report_docx.py` — compat shims that only re-export; covered by twin-identity tests on the canonical modules
- `main.py` — three-line console-script launcher; covered by the `edf-collector` console-script smoke unit test
- `.git/`, `__pycache__`, `build/`, `dist/`, `*.spec` — excluded by `[tool.coverage.run] omit`

## The 90% floor + 100% aspiration

- **90% is the CI floor**: a PR that drops coverage below 90% fails CI. `coverage report --fail-under=90` enforces this.
- **100% is the aspiration target**: pursue with diminishing-returns threshold. Where a line or branch genuinely cannot be exercised by a test, mark it `# pragma: no cover` with a one-line justification.

## Strict `# pragma: no cover` policy

The pragma is load-bearing — it tells the reader "I, the author, attest that this line cannot be tested in this environment, and here is why." Game the metric and you lose the reader's trust.

### When `# pragma: no cover` IS allowed

Each `# pragma: no cover` MUST include a one-line comment explaining why the line is unreachable. Examples:

```python
# 1. Platform guard for an OS we don't test on Linux CI:
if sys.platform == "darwin":  # pragma: no cover, macOS-only path not tested on Linux CI
    ...

# 2. Type-checking-only block — never executed at runtime:
if TYPE_CHECKING:  # pragma: no cover, type-only branch
    ...

# 3. ImportError fallback for a hard-required dependency in dev/CI extras:
try:
    import scipy  # type: ignore
except ImportError:  # pragma: no cover, scipy is in dev extras — ImportError impossible in CI
    HAS_SCIPY = False


# 4. Defensive fallback that the public API contract makes unreachable:
def f(x: int) -> int:
    if x < 0:
        raise ValueError("negative")
    # unreachable from public API (= >0 by type contract)
    if x is None:  # pragma: no cover, ruled out by type contract
        return 0
    return x * 2
```

### When `# pragma: no cover` is FORBIDDEN

The pragma is not a "this is annoying to test" escape hatch:

- ❌ Pragma'ing a branch just because mocking the dependency is hard
- ❌ Pragma'ing an `else:` clause for an error path that COULD be tested by injecting bad input
- ❌ Pragma'ing a Tk dialog call just because it needs a mock (`unittest.mock.patch("tkinter.filedialog.askdirectory")` works fine)
- ❌ Pragma'ing a defensive `assert` you could test by passing the violating input

If you find yourself wanting to pragma a line for any of these reasons, write the test instead. The pragma is for "no test could exercise this branch in this environment", not "the test would take effort".

### Review discipline

Every `# pragma: no cover` gets scrutinized in code review. The reviewer asks:
1. Is the branch genuinely unreachable in our test environment?
2. Does the comment explain WHY (not just WHAT)?
3. Could a test be written with reasonable effort?

If the answer to #3 is yes, the test gets written and the pragma gets removed.

## How to extend coverage

### 1. Find the gap

```bash
coverage report --skip-covered   # sort by lowest coverage
coverage html                    # then open htmlcov/index.html and click into a file
```

The HTML view highlights uncovered lines in red and uncovered branches in dim red.

### 2. Read the gap

Uncovered lines fall into three buckets:

- **Easy**: a public function with no test → write a unit test calling it with synthetic input.
- **Branchy**: a conditional branch with no test for one side → add a test that takes the uncovered branch.
- **Boundary**: a Tk dialog / file I/O / external-API call → mock at the boundary (`unittest.mock.patch("tkinter.filedialog.askdirectory", return_value="/tmp/x")`), assert on the resulting state change (never on call counts).

### 3. Write the test

Prefer real synthetic fixtures over mocked dependencies. The test suite already has synthetic frozen DataFrames in `tests/fixtures/` — extend those rather than mocking pandas.

For Tk dialogs in `ui/app.py`, mock at the `tkinter.filedialog.*`/`tkinter.simpledialog.*`/`tkinter.messagebox.*` boundary — invoke the handler, assert state changes (variable values, button states). Never assert mock-call counts (brittle to refactor).

For CLI paths in `io/cli.py`, use `monkeypatch.setattr("sys.argv", [...])` + `capsys` for stdout. Never invoke real file paths.

### 4. Run the focused test under coverage

```bash
coverage run --branch --source=edf_bill_fetcher -m pytest tests/test_your_new_test.py
coverage report -m
```

Confirm the targeted lines moved from missed to covered.

### 5. Run the full suite + gate

```bash
coverage run --branch -m pytest
coverage report --fail-under=90
```

If the gate passes, ship. If it fails, the new code dropped overall coverage — write more tests.

## Coverage regression over time

The 90% floor is the durable defense. A PR that adds new code without tests will drop coverage below 90% and fail CI. Adding new code without tests is therefore self-defeating — the gate forces you to write the test in the same PR.

## When to update this doc

- Bump the floor (e.g. 90% → 95%) — note the date and the new floor in this doc + update `pyproject.toml [tool.coverage.report] fail_under`.
- Change the source set (e.g. include `tests/` in coverage) — update the "What IS measured" section.
- Add a new measurement baseline — save it under `docs/coverage/baseline-YYYY-MM-DD.txt` and link it from the "Current baseline" section above.
