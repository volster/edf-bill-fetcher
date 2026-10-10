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
pip install -e ".[dev,msg]"  # include optional Outlook MSG parsing

# Build install: + PyInstaller for one-file executables:
pip install -e ".[dev,build]"
```

Reuse a working environment where available. Tkinter and the Xvfb binary on headless Linux are system prerequisites; pip extras do not install them. Native PST/OST builds may require compiler/autotools headers when no compatible wheel is available.

## Running tests

```bash
pytest -v
```

CI declares the following matrix; inspect results for the actual revision before claiming it passes:

- **Linux**: Python 3.10 / 3.11 / 3.12
- **macOS**: same
- **Windows**: same (tkinter ships with the standard Python installer on Windows)
- **Headless Linux** (CI / containers): install the Xvfb system binary; `pytest-xvfb` in the dev extra can then activate a virtual display

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
mypy .                 # matches the current CI invocation
```

All three are enforced by the CI workflow. See `.github/workflows/ci.yml` for the exact invocation; do not infer branch protection or successful checks from this document.

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
mypy .
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

1. Add an entry to `REPORT_SECTIONS` in `edf_bill_fetcher/models/report_sections.py` with `key`, `title`, and optionally `is_appendix`.
2. Add the matching key to `ReportOptionsDialog.SECTIONS` in `edf_bill_fetcher/ui/app.py`.
3. Add a `def create_<name>(...)` builder function in `edf_bill_fetcher/io/reporters/pdf_report.py`.
4. Wire the builder into PDF, DOCX and HTML dispatchers. Missing sections fail the artifact rather than publishing a partial success.
5. The structural parity test `tests/test_dispatch_parity.py` will catch missing dispatcher wiring in CI.

### A new detector (processors layer)

1. Add the detector function in `edf_bill_fetcher/processors/<area>.py` (e.g. `detection.py` for billing-pattern detection, `matching.py` for account/contract matching).
2. Detectors receive DataFrames as function arguments — never pull from module-scope state.
3. New detectors should return data for writers to format. Existing processors still import some writer-facade helpers; strict dependency inversion is not currently enforced. Avoid adding renderer dependencies to analytical code.
4. Add the detector to `run_analysers` in `io/writers/analysis.py` if it should be one-shot available.
5. Add unit tests in `tests/test_<area>.py` calling the detector with synthetic DataFrames covering each branch.

## Architecture reference

See [`docs/ARCHITECTURE.md`](ARCHITECTURE.md) for:
- The full package layout
- Actual dependencies and remaining processor/writer coupling
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
3. `mypy .`
4. `pytest --cov=. --cov-report=xml --cov-report=term-missing` (Linux uses `xvfb-run -a`)

CI is a regression gate. Functional extraction, saved-state reload, all requested exports, source links, GUI lifecycle and packaged runtime capabilities also need verification. Record focused checks separately from full matrix results.

## Shared result architecture

`domain/` contains standard-library-only immutable request, evidence and analysis types. Nested data uses canonical `FrozenJSON`; thawing or model decoding creates fresh objects. Money at the result boundary uses finite Decimal values and canonical two-decimal JSON strings. Missing values remain unavailable.

The service pipeline is `capture_legacy` → `prepare_snapshot` → `analyse_prepared`. The shared analysis stores all ten model entries, availability reasons, analytical frames, the workbook timeline and SAP source indexes. Keep financial calculations here. `ReportView` exposes fresh decoded frames/models; `render_xlsx`, `render_pdf`, `render_docx` and `render_html` format these values. Legacy public generator functions capture/analyse once and delegate. A new analytical section needs a stored model and explicit renderer wiring; it must not calculate inside the renderer.

`io/persistence/result_codec.py` implements the exact v3 envelope; `result_store.py` provides strict reads and atomic publication. V3 loading does not prepare or recalculate. Legacy migration preserves unknown fields and original/canonical distinctions and records historical provenance gaps. Changing analytical configuration creates a new accepted result, rather than editing a previously accepted file.

`application/run_service.py` coordinates injected collection, source-copy, result-storage and renderer functions. It records atomic stage manifests with hashes of published result/report bytes, preserves successful artifacts after failure, checks cancellation between stages and always cleans resources. The CLI supplies explicit destinations; the GUI reserves unique folders before launching a non-daemon worker. UI widget values are captured on the main thread, and progress/messages return through the event queue.

When extending this pipeline, verify an independently calculated money example, a strict persistence round trip, failure preservation and a rerender with calculation functions disabled. Exercise real Tk for worker/cancel/close changes. Keep the CI matrix, coverage threshold and lint/type checks intact; successful local examples do not establish cross-platform support.

## Validation and release provenance

CI has one required aggregate validation job which succeeds only if every Python/OS matrix leg succeeds. Tag releases call that same reusable CI workflow for their own revision. Release build and publication depend on validation; checkout and build provenance use its validated SHA. All `v*` tags use this gate. Branch CI still triggers automatic executable builds after aggregate success. Manual executable builds record `manual-unvalidated` provenance and never publish a release.

Changing workflow YAML requires reviewing the dependency graph and proving that a failing sibling matrix leg blocks publication. Remote workflow results remain the evidence for actual cross-platform builds; syntax validation alone is insufficient. Do not tag or publish as a side effect of implementation work.
