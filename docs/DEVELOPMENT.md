# Development

How to run tests, add features, and ship changes to the EDF Bill Fetcher codebase.

## Set up

```bash
git clone https://github.com/volster/edf-bill-fetcher.git
cd edf-bill-fetcher

# Default install: runtime only:
pip install -e .

# Dev install: runtime + test + lint + type-check toolchain:
pip install -e ".[dev]"

# Build install: + PyInstaller for one-file executables:
pip install -e ".[dev,build]"
```

The `edf` conda env on the maintainer's machine already has everything installed. If you're contributing from a fresh checkout, `pip install -e ".[dev]"` is sufficient for the test/lint/type-check workflow.

## Running tests

```bash
pytest -v
```

That's the entire command. The full suite runs green on:

- **Linux**: Python 3.10 / 3.11 / 3.12
- **macOS**: same
- **Windows**: same (tkinter ships with the standard Python installer on Windows)
- **Headless Linux** (CI / containers): `pytest-xvfb` (in the `dev` extras) auto-activates a virtual X server so the tkinter-dependent tests run green without a manual xvfb setup

`pytest-xvfb` is **opt-out**, not opt-in. If you want to debug a tkinter test on a real display (i.e. your laptop), pass `--no-xvfb` to disable the plugin for that run:

```bash
pytest --no-xvfb tests/test_app.py    # use your real X server
```

### Test layers

The suite is organized into three lake levels (largest first):

- **Unit tests** (`tests/test_*.py`) — most files. Pin the behavior of public functions and the structural invariants of the registry/dispatcher.
- **Audit regression tests** — pin the contracts the report depends on. Exist *because* real-data review exposed one or more real defects. Do not edit these without re-running the audit-pass analysis:
  1. `tests/test_audit_pass_1.py` — reading-pattern ordering, `detect_pdf_format`, `process_text` heuristic-fallback, `_detect_payment_patterns`, `_analyze_tariff_impact`, `_data_quality_report`, `process_pst_file` / `process_ost_file`, `compute_dispute_flags`.
  2. `tests/test_report_version.py` — cover page reflects the `pyproject.toml` version; falls back to a stable default when `pyproject.toml` is unreadable.
  3. `tests/test_dispatch_parity.py` — REGISTRY ↔ PDF dispatcher ↔ DOCX dispatcher key-set parity is locked in.
- **Integration smoke** (`tests/test_integration_pipeline.py`) — drives the bundled synthetic bill PDF (`tests/fixtures/sample_bill.pdf`) through the full PDF → engine → reportlab PDF + openpyxl XLSX pipeline and asserts the extracted fields. The fixture is regenerated via `tests/fixtures/generate_bill_fixture.py` if missing — a fully-synthetic, deterministic dataset (FAFA policy, no real EDF data).

### Running a focused subset

```bash
pytest tests/test_writers.py              # one file
pytest tests/test_writers.py::TestWriters # one class
pytest tests/test_writers.py::TestWriters::test_evidence_sheet_is_written    # one test
pytest -k "reconciliation"               # by name pattern
pytest -m "not slow"                      # by marker
```

### Common pytest flags

| Flag                       | Effect                                                 |
| -------------------------- | ----------------------------------------------------- |
| `--no-header`              | strip the leading pytest version banner                |
| `-q`                       | quiet — one dot per test                              |
| `-v`                       | verbose — full test name per line                       |
| `--lf`                     | rerun only the tests that failed last time             |
| `--no-xvfb`                | disable `pytest-xvfb` — use the real display (laptop)   |
| `-x`                      | stop on first failure                                  |
| `-k <expr>`                | run tests matching the expression                      |

## Linting / formatting / type-checking

```bash
ruff check .           # PEP 8 + PEP 257 + import-ordering — fast
ruff format --check .   # formatting check (no write)
ruff format .           # apply formatting
mypy edf_bill_fetcher   # strict type-check
```

All three are enforced in CI. A PR that fails any gate cannot merge. See `.github/workflows/ci.yml` for the exact invocation.

### Ruff rule selection

Configured in `pyproject.toml [tool.ruff.lint]`:

- `E`, `W` — PEP 8 (style + warning)
- `F` — pyflakes (unused imports, undefined names)
- `I` — isort (import ordering)
- `D` — pydocstyle (PEP 257 docstring conventions)

### PEP 257 D-rule relaxations

The `D` rule family is enforced in near-full strict mode: the nine rules relaxed during the refactor window (`D100`, `D104`, `D105`, `D107`, `D400`, `D401`, `D406`–`D409`) were un-ignored in the docstring compliance pass, and the retroactive class/method/parameter backlog (`D101`/`D102`/`D103`/`D415`/`D417`) was cleared during modularization. The only remaining relaxations are deliberate style conflicts, each with a comment in `pyproject.toml`:

| Rule    | Why relaxed                                                            |
| ------- | -------------------------------------------------------------------- |
| `D203`  | blank line before class docstring conflicts with `D211` — we use no-blank-line |
| `D213`  | multi-line docstring summary on second line — we use first-line summary (`D212`) |

Tests (`tests/*`) remain exempt from all `D` rules via per-file-ignores.

## Type checking

```bash
mypy edf_bill_fetcher
```

Configured in `pyproject.toml [tool.mypy]`:

- `check_untyped_defs = true` — even untyped functions get checked
- `disallow_incomplete_defs = true` — partially-typed function signatures fail

New code MUST have complete type annotations. Test fixtures count: any new test function with parameters must have types on all parameters including pytest fixtures (`def test_x(capfox, tmp_path)` fails; `def test_x(capfox: pytest.CaptureFixture[str], tmp_path: pathlib.Path) -> None:` passes).

## Test coverage

See [`docs/COVERAGE.md`](COVERAGE.md) for the full coverage protocol. The short version:

- CI gate: `coverage report --fail-under=90` — any PR that drops below 90% fails
- Aspiration: 100% of testable code, with `# pragma: no cover` allowed only for genuinely unreachable lines (each with a one-line justification comment)
- Run locally: `coverage run --branch -m pytest && coverage report`

## How to add a new feature

### A new Excel sheet writer

1. Create the writer function in `edf_bill_fetcher/io/writers/<sheet_name>.py`. Follow the existing pattern (file signature: takes `ws: openpyxl.Workbook.active`, `df: pandas.DataFrame`, optional config args).
2. Add the writer's name to `__all__` in `edf_bill_fetcher/io/writers/__init__.py` (eager re-exports — the PEP 562 shim layers were removed, so no `__getattr__` is needed).
3. Wire the writer into `export_to_excel` in `edf_bill_fetcher/io/writers/export.py` — call it in the correct sheet-order position with the correct conditional-emission gating.
4. Add the writer to the `tests/test_writers.py` import test (if it's a new function symbol that should be importable from `edf_bill_fetcher.io.writers`).
5. Add unit tests in `tests/test_<sheet_name>_writer.py` calling the writer with synthetic DataFrames covering the empty / single-row / multi-row branches.
6. Run gates: `ruff check .`, `mypy edf_bill_fetcher`, `pytest -v`. If the writer adds new statements, also run `coverage run --branch -m pytest && coverage report --fail-under=90`.

### A new PDF/DOCX report section

1. Add an entry to `REPORT_SECTIONS` in `edf_bill_fetcher/io/reporters/pdf_report.py` with `key`, `title`, and optionally `is_appendix`.
2. Add the matching key to `ReportOptionsDialog.SECTIONS` in `edf_bill_fetcher/ui/app.py`.
3. Add a `def create_<name>(...)` builder function in `edf_bill_fetcher/io/reporters/pdf_report.py`.
4. Wire the builder into the `section_builders` dispatch dict in **both** `generate_ombudsman_pdf` and `generate_ombudsman_docx`. Forgetting this raises a clear `RuntimeError` at report-render time.
5. The structural parity test `tests/test_dispatch_parity.py` will catch missing dispatcher wiring in CI.

### A new detector (processors layer)

1. Add the detector function in `edf_bill_fetcher/processors/<area>.py` (e.g. `detection.py` for billing-pattern detection, `matching.py` for account/contract matching).
2. Detectors receive DataFrames as function arguments — never pull from module-scope state.
3. NO framework imports at module scope in `processors/`. Pure stdlib + pandas + sibling processors + helpers. If you need openpyxl, you're at the wrong layer — the detector returns a result; the writer formats it.
4. Add the detector to `run_analysers` in `io/writers/analysis.py` if it should be one-shot available.
5. Add unit tests in `tests/test_<area>.py` calling the detector with synthetic DataFrames covering each branch.

## Architecture reference

See [`docs/ARCHITECTURE.md`](ARCHITECTURE.md) for:
- The full package layout
- The hexagonal layering rules (which layer can import what)
- The public import API (flat `from edf_bill_fetcher.io.writers import X` and submodule-scoped `from edf_bill_fetcher.io.writers.export import export_to_excel`)
- The maintenance rule (no backward-compat shims remain: top-level `edf_collector.py` / `edf_report*.py` were removed in the modularization, as were the temporary `writers` / `io.writers` PEP 562 layers — new symbols re-export eagerly via `io/writers/__init__.py`'s `__all__`)

## Commit conventions

The repo uses [Conventional Commits](https://www.conventionalcommits.org/):

- `feat:` new feature
- `fix:` bugfix
- `refactor:` code change that neither adds a feature nor fixes a bug
- `docs:` documentation-only change
- `test:` test-only change
- `chore:` tooling, dependency, config
- `style:` formatting, PEP 8/257/20 conformance fixes

Scopes are optional. Examples: `refactor(writers): extract export_to_excel to io/writers/export.py`, `test(processors): add comprehensive detection coverage`, `chore(deps): add pytest-xvfb to dev extras`.

## CI

CI runs on every push and every PR. The matrix is Python 3.10 / 3.11 / 3.12 × ubuntu / windows / macos. All four gates must pass on every cell:

1. `ruff check .`
2. `ruff format --check .`
3. `mypy edf_bill_fetcher`
4. `pytest -v`

A release is one CI green away from shippable.
