# Development

[Documentation index](README.md) · [Architecture](ARCHITECTURE.md) · [CI troubleshooting](CI.md)

## Set up once

Follow [source setup](BUILDING.md) and use `bootstrap.py develop --install-system`
to reuse or create the project Conda environment (or `.venv` if Conda is absent),
install constrained `[dev,build,msg]` dependencies and check native prerequisites.
Use `--conda-env NAME` for a custom project environment or `--venv .venv` to force venv.

Commands below use `python` to mean that environment's interpreter. For Conda,
run `conda activate edf-bill-fetcher` first or prefix commands with
`conda run --no-capture-output -n edf-bill-fetcher`. For venv without activation,
use `.venv/bin/python` on macOS/Linux or `.\.venv\Scripts\python.exe` on Windows.

## Daily checks

```console
python -m ruff check .
python -m ruff format --check .
python -m mypy .
python -m pytest tests/test_excel_review.py
```

Choose tests for the behavior you change. Use `python -m ruff format <paths>` to
format edited Python files. For the full suite with the same coverage/reporting
command as CI:

```console
python ci_pytest.py
```

That runner executes the suite once and preserves its exit code, JUnit and coverage
reports. See [coverage guidance](COVERAGE.md) for measurement details. CI runs on
Python 3.10, 3.11 and 3.12 across Ubuntu, Windows and macOS; a local pass does not
establish results for the other platforms.

Useful pytest options are `-k expression` to select names, `-x` to stop on the first
failure, `--lf` to rerun previous failures and `-q` for shorter output. For example:

```console
python -m pytest tests/test_claim_reconciliation.py -q
python -m pytest -k "reconciliation" -x
```

### GUI and native checks

Tests automatically select Qt's offscreen platform on headless Linux. CI sets
`QT_QPA_PLATFORM=offscreen` on every OS. For an interactive Linux desktop check,
select the desktop's platform, such as `QT_QPA_PLATFORM=xcb`, instead. Normal
app startup does not force the test platform. Tk and Xvfb are not prerequisites.

Use `python ci_runtime.py` for the workspace startup preflight and
`python main.py --doctor` for installed capabilities. Worker, cancellation,
shutdown and layout changes also need a real desktop check. Frozen application
changes need the [native build checks](BUILDING.md#build-for-this-computer);
source tests alone do not exercise bundled dependencies.

## Where to make a change

| Change | Main areas and useful regression examples |
| --- | --- |
| PDF/email extraction | `collectors/`, `io/adapters/`, `processors/`; `test_integration_pipeline.py`, `test_extraction_recovery.py`. |
| Financial interpretation or deduplication | Preparation/analysis services and processors; `test_financial_contracts.py`, `test_review_financial_2026.py`. |
| Saved results and rerendering | `domain/`, `io/persistence/`, `models/report_view.py`; `test_result_snapshot.py`, `test_saved_result_validation.py`, `test_saved_result_rendering.py`. |
| Corrections or review drafts | `application/record_review.py`, `review_session.py`, `io/review_workbook.py`, Qt review widgets; `test_excel_review.py`, `test_native_review.py`, `test_native_review_ui.py`. |
| Claims and complaint packs | `processors/claim_reconciliation.py`, `io/claim_workbook.py`, `application/complaint_pack.py`; `test_claim_reconciliation.py`, `test_claim_outputs.py`, `test_claim_ui.py`. |
| Desktop behavior | `ui/qt_app.py` and related Qt modules; `test_qt_desktop.py`, `test_review_gui_2026.py`. |
| Output presentation | `io/writers/`, `io/reporters/`, shared table models; `test_output_usability.py`, `test_dispatch_parity.py`. |

Test filenames in the table are under `tests/`. Read the
[architecture contracts](ARCHITECTURE.md) before changing shared calculations or
persistence. Use synthetic evidence: private bills, account exports and complaint
positions must not enter repository fixtures.

### Add an analysis or output section

1. Define the analytical result and compute it through the preparation/analysis
   services. Keep financial calculations out of renderers.
2. Store it in the accepted snapshot and expose it through `ReportView`. Define
   missing/insufficient-data behavior and compatibility with older saved results.
3. For an Excel view, follow a nearby writer's actual signature and wire it into
   `io/writers/export.py`. Re-export a new public writer through
   `io/writers/__init__.py` when needed.
4. For a document section, add its key/title to
   `models/report_sections.py` and wire all three PDF, DOCX and HTML dispatchers.
   The Qt report selector reads this registry.
5. Verify independently calculated expected values, persistence round trips and
   consistent rendering across formats. Extend dispatch-parity checks where
   appropriate. Rerender a saved result with calculation functions disabled to
   prove that formatting does not silently recompute it.

Legacy generator entry points still capture/analyse data and delegate to the
shared pipeline. New features should work through the accepted result. The old
root `edf_collector.py` and `edf_report*.py` modules have been removed; canonical
imports live under `edf_bill_fetcher`. Some processors still use the historical
`writers` helper facade. Avoid adding renderer dependencies to analysis code.

### Add a review input

Use the shared validation/audit contract for native and Excel edits. Preserve
original observations, stable IDs and evidence bindings. Changed evidence must
invalidate dependent decisions rather than quietly retain their old approval.
Exercise both import paths, missing/invalid inputs, draft recovery and repeated
review rounds. A reviewed pack is a new accepted result; it must not overwrite
the input pack.

### Write a useful regression

Reproduce the defect at its public boundary with a small synthetic example and
an expected result derived independently of the implementation. For monetary
changes, show the arithmetic. Include a plausible neighboring case that should
behave differently, such as a repeated balance versus a new transaction.

Use temporary files for I/O and mock external boundaries where needed. For GUI
tests, inspect user-visible state and behavior. Call counts are useful only when
the count itself is the contract, such as preventing duplicate publication.
Preserve failure and cancellation evidence, source hashes and filename mappings.
Read/write JSON and other UTF-8 artifacts with an explicit encoding; Windows'
default encoding can differ from Linux/macOS.

## Style and type checking

`pyproject.toml` is the source of truth for Ruff, mypy and pytest settings. Ruff
checks style, imports, docstrings and additional correctness rules; tests have
specified per-file exceptions. Do not copy a partial rule list into a new config.

Prefer complete annotations for new functions. Mypy checks untyped function
bodies and rejects partially annotated signatures (`check_untyped_defs` and
`disallow_incomplete_defs` are enabled). Fully unannotated definitions are still
allowed by the current configuration; they do not fail merely for lacking types.
If you annotate a test, include its fixture parameters and return type.

Coverage is a useful gap detector, not proof of correctness. Do not lower the
90% gate or add exclusions to make a failing change pass. See the
[coverage policy](COVERAGE.md#coverage-exclusions).

## Documentation and delivery

Update the guide that owns the changed behavior. User instructions belong in
[the user guides](README.md#using-the-app), implementation contracts in
[architecture](ARCHITECTURE.md), and setup/build commands in [building](BUILDING.md).
Keep the root README short. Record notable product changes in
[CHANGELOG.md](../CHANGELOG.md); historical checks must be labelled as historical.

Use a commit subject that describes the concrete change. Existing history uses
both plain subjects and prefixes such as `fix:`, `feat:` and `docs:`; no commit
prefix is enforced by the project tooling.

CI is configured for pushes and pull requests targeting the branches listed in
[ci.yml](../.github/workflows/ci.yml), and is reusable for releases. Its aggregate
validation succeeds only when all nine matrix legs pass. Automatic executable
builds use the validated revision; manually dispatched builds remain labelled
`manual-unvalidated`. Tagged releases validate the tagged revision before building
and publishing it. See [CI](CI.md) for diagnostics and provenance.

Workflow changes need review of the dependency graph, including whether a failing
matrix leg blocks publication. Record actual local and remote checks separately;
YAML validation alone does not establish that a packaged app runs.
