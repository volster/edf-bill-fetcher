# Test Coverage to 100% + README/Docs Refresh — Design Spec

**Date**: 2026-08-02
**Status**: Approved (via user directives: "Set up xvfb display fixture", "Fix the errors while you're at it", "If feasible let's shoot for 100% ... We might not get there, but we can try!", "I'd like to also add a readme / documentation refresh to the list")
**Branch**: `refactor` (continues from `modularization-complete` tag at `98cd0f2`)
**Predecessor**: `.omo/plans/2026-07-28-modularization-completion.md` (refactor itself, now complete)

---

## 1. Reframe — The "Pre-Existing Failures" Were a Missing Test Dependency

Throughout the entire modularization refactor, the gate was reported as `740 passed / 4 failed / 31 errors / 7 skipped`, with the 4 failures + 31 errors characterized as "pre-existing tkinter-Display baseline, zero regressions". Brainstorming-context exploration revealed this baseline was a phantom: the failures existed because `pytest-xvfb` was missing from the `edf` conda environment.

Installing `pytest-xvfb` (already packaged with `xvfbwrapper` + `pyvirtualdisplay` transitively) took the suite from `740 passed / 4 failed / 31 errors` → `786 passed / 0 failed / 0 errors / 7 skipped` instantaneously. The 4 pre-refactor failures + 31 pre-refactor errors were never "pre-existing" in any meaningful sense — they were a missing dev dependency.

**Reframe consequences:**
- Phase 1 (originally "set up xvfb + fix 4+31 errors") is effectively zero work: installing `pytest-xvfb` was the entire fix.
- The refactor's "no regressions" claim remains intact — but the baseline against which regressions were measured was wrong. The true pre-refactor baseline (with xvfb installed) would have been higher.
- Coverage starting point moves up: was 68% (with Tk-Display tests skipped) → 71% now (all 786 tests running under xvfb). The xvfb install delivered +3 percentage points of coverage for free.

**Artifact:** Add `pytest-xvfb` to the `[dev]` extras in `pyproject.toml` so future `pip install -e .[dev]` reproducibly gets the green suite. Optionally add a one-line note to the README "Run tests" section that xvfb is auto-active on Linux headless environments via the plugin.

---

## 2. Coverage Target Architecture

### 2.1 Targets
- **Floor: 90%** — non-negotiable gate. Enforced in CI via `coverage report --fail-under=90`.
- **Aspiration: 100%** — pursue with diminishing-returns threshold. Accept `# pragma: no cover` for genuinely unreachable code, with a one-line justification per pragma.

### 2.2 Strict `# pragma: no cover` Policy

Each `# pragma: no cover` MUST have a one-line comment explaining WHY the line is unreachable. Acceptable use cases:

```python
if sys.platform == "darwin":  # pragma: no cover -- platform-guard for macOS-only path on Linux CI
    ...

if TYPE_CHECKING:  # pragma: no cover -- type-only branch never runs at runtime
    ...

try:
    import scipy  # type: ignore
except ImportError:  # pragma: no cover -- scipy is in dev extras, ImportError impossible in CI
    HAS_SCIPY = False
```

**FORBIDDEN pragma uses** (would game the metric):
- Pragma'ing branches just because they're hard to test
- Pragma'ing `else:` clauses for error paths that COULD be tested with bad input
- Pragma'ing Tk dialog calls just because they need mocking (these ARE testable via `unittest.mock.patch`)
-Blanket pragma'ing at function level to skip the whole function (must be line-scoped or branch-scoped)

### 2.3 Quality Bar for New Tests — No Mock-Brittleness

To prevent the coverage push from creating tests that pass on mocks but don't catch real bugs:

- **Prefer integration-style tests against synthetic fixtures over permitted-mock tests.** For `io/adapters/pst.py`, use the existing `tests/fixtures/pst_attachment_fixture.py` synthetic `pypff` shape, extended — NOT a mock-pypff.
- **For Tk dialogs in `ui/app.py`**, mock at the `unittest.mock.patch("tkinter.filedialog.askdirectory")` boundary — single-mock-per-test, never mock Tk internals. Assert handler state-mutation, never assert mock-call-counts.
- **For `io/cli.py` CLI argv paths**, use `monkeypatch.setattr("sys.argv", [...])` + `capsys` for stdout — never invoke real file paths.
- Each new test file must satisfy PEP 257 D rules per project (project memory #255).

### 2.4 Coverage Measurement Protocol

- Coverage gate added to CI: `coverage report --fail-under=90`
- Coverage config in `pyproject.toml`:
  ```toml
  [tool.coverage.run]
  source = ["edf_bill_fetcher"]
  branch = true
  
  [tool.coverage.report]
  fail_under = 90
  show_missing = true
  skip_covered = true  # reduces noise, focuses report on gaps
  ```
- Baseline coverage report committed to `docs/coverage/2026-08-02-baseline.txt` for diff-review against future regressions

---

## 3. Phase 2 — Coverage Gap Strategy (Module-by-Module)

**Goal**: cover +934 statements (to hit 90% floor) up to +1521 statements (to hit 100% aspiration) across 15 modules. Each module has a specific test-fixture strategy — not all are equally tractable.

### 3.1 Tier A — Easy Wins (~6 modules, no exotic infra)

| Module | Current | Missed | Strategy |
|--------|---------|--------|----------|
| `io/reporters/__init__.py` + `docx_report.py` + `pdf_report.py` | 0% | 10 | PEP 562 lazy shim bodies — add tests that `getattr(module, name_in___all__)` each entry (validates re-export identity) plus a `pytest.raises(AttributeError)` for a missing name (covers the `__getattr__` fallback branch) |
| `io/writers/__init__.py` | 7% | 34 | Same PEP 562 `__getattr__` test pattern — assert every `__all__` name resolves, then a negative test for the missing-attribute branch |
| `processors/extraction.py` | 6% | 112 | Barely tested at all — likely a Task 5 extraction artifact that never had tests written. Write unit tests for each public function using synthetic DataFrames as inputs — no I/O, no mocking. **Highest ROI in the entire spec.** |
| `io/writers/statistical.py` | 79% | 25 | Branches not covered likely the "no statistical data" early-return + the conditional emission paths. Add tests that call `write_statistical_sheet` with empty / minimal DataFrames. |
| `processors/forecasting.py` | 79% | 13 | Statsmodels-import-guard branches + an empty-data early-return path. Test with `HAS_STATSMODELS=False` mocked + empty data. |
| `processors/detection.py` | 71% | 70 | Detector branches for unusual inputs (no anomalies found, all-anomaly, single-row evidence). Test each public detector with synthetic DataFrames covering each branch. |

**Tier A coverage delta**: ~264 missed → ~+264 covered. Running total: 4347+264 = 4611 (78.6%)

### 3.2 Tier B — Mid-Complexity (Test Infrastructure Required)

| Module | Current | Missed | Strategy |
|--------|---------|--------|----------|
| `io/cli.py` | 30% | 134 | CLI entry paths — use `monkeypatch.setattr("sys.argv", [...])` + `capsys` for output capture + `tmp_path` for output files. Cover `main()`, `run_cli_extract`, `run_cli_pdf_report`, `run_cli_docx_report` argument parsing + happy + error paths. `--help` paths via `SystemExit` catch. |
| `io/adapters/pst.py` | 8% | 82 | Extend `tests/fixtures/pst_attachment_fixture.py` synthetic `pypff` shape to cover all branches of `_pst_attachment_filename`: missing `PR_ATTACH_LONG_FILENAME` (falls back to short), missing both (falls back to `Attachment_N.pdf`), corrupt record set (AttributeError swallow), multiple attachments in different states. |
| `io/writers/rebilling.py` | 44% | 72 | Sheet-builder branches — call `write_rebilling_sheet` with empty rebilling list, single entry, multiple entries, each flag combination. Use synthetic DataFrames. |
| `io/writers/meter.py` | 65% | 89 | Meter writer branches — empty readings, single contract, multi-contract, missing-period edge cases. Synthetic DataFrames. |
| `io/writers/back_billing.py` | 76% | 39 | Back-billing sheet branches — no events, single event, multi-event, with/without matched-EDF context. |
| `processors/analysis.py` | 69% | 47 | Analysis branches — `compute_dispute_flags` paths for each flag, with synthetic engine records. |
| `collectors/engine.py` | 68% | 171 | EvidenceEngine branches — most-missed likely the PDF-parse error paths + multi-source ingestion branches. Mock `pdfplumber.open` to raise various exceptions and assert the error is logged but doesn't crash. Each `process_*` method needs happy + error path coverage. |

**Tier B coverage delta**: ~634 missed → ~+634 covered. Running total: ~5245 (89.4%) — just under 90% floor.

### 3.3 Tier C — Hardest (the GUI)

| Module | Current | Missed | Strategy |
|--------|---------|--------|----------|
| `ui/app.py` | 43% | 293 | Tkinter App + ReportOptionsDialog. Covered now via xvfb fixture. Remaining missed: modal dialog handlers (`_open_output_folder_picker`, `_open_report_options`, `_open_pdf_save_dialog`) — each invokes `filedialog.askdirectory` / `simpledialog` / `asksaveasfilename`. Mock these at the boundary via `unittest.mock.patch("tkinter.filedialog.askdirectory", return_value="/tmp/x")` — invoke the handler, assert the resulting state change. The EXTRACT workflow with cancel/transition states needs careful fixture sequencing. Most testable. |
| `writers/_helpers.py` | 68% | 112 | Shared writer helpers — formatting/hyperlink/date parsing. Likely some branches for empty inputs / edge cases. Synthetic DataFrame-based tests. |
| **Remaining unreachable defensive code** | — | est. ~152 | `# pragma: no cover` with one-line justifications for: `if TYPE_CHECKING:` blocks, `except ImportError` for hard-required-import paths (where ImportError impossible in CI), `if sys.platform` branches for OSes we don't test on Linux CI, dead-code fallback branches. Each pragma has a comment explaining WHY. |

**Tier C coverage delta**: ~405 of 457 missed → ~+405 covered. Running total: ~5650 (96.3%)

### 3.4 Final State Projection
- **Tier A complete**: 78.6%
- **+Tier B complete**: 89.4% (just under 90% floor — Tier C needed)
- **+Tier C complete**: 96.3% realistic ceiling
- **+Pragmas**: ~3-4% gap closed by legitimate `# pragma: no cover` → ~100% reportable, with each untested line honestly documented

**Failure budget**: 96.3% is the realistic testable ceiling. The last ~3.7% must be honest `# pragma: no cover` for genuinely-unreachable code. If during implementation we hit 96%+ and the remaining 4% is honest-defensive-only, declaring 100% with pragmas is the legitimate matching of metric to reality. If we hit diminishing returns before 95% (e.g., a Tk App path resists mocking cleanly), we stop at the level achieved and document the residual gap — better an honest 96% than a gamed 100%.

### 3.5 Execution Ordering (Parallelizable Subagent Dispatch)

Pre-extract per rule #213: each subagent brief contains the missed-line inventory for its target module(s) + concrete test patterns to follow. Subagents should NOT explore — they should write tests against the pre-extracted inventory.

**Wave 1 (parallel, 6 subagents)** — Tier A, all independent:
1. PEP 562 shim tests for `io/reporters/__init__.py` + `docx_report.py` + `pdf_report.py`
2. PEP 562 shim tests for `io/writers/__init__.py`
3. `processors/extraction.py` test suite (NEW — 6% → ~100%)
4. `io/writers/statistical.py` branch tests
5. `processors/forecasting.py` branch tests
6. `processors/detection.py` branch tests

**Wave 2 (parallel, 7 subagents)** — Tier B, all independent:
7. `io/cli.py` CLI argv + capsys tests
8. `io/adapters/pst.py` synthetic-pypff fixture extension tests
9. `io/writers/rebilling.py` branch tests
10. `io/writers/meter.py` branch tests
11. `io/writers/back_billing.py` branch tests
12. `processors/analysis.py` branch tests
13. `collectors/engine.py` error-path mock tests

**Wave 3 (single, critical-path, 1 subagent)** — Tier C `ui/app.py`:
14. `ui/app.py` GUI handler tests with filedialog/messagebox boundary mocks

**Wave 4 (parallel, 2 subagents)** — Tier C cleanup:
15. `writers/_helpers.py` branch tests
16. Final pragma audit pass — add `# pragma: no cover` with comments to genuinely unreachable code identified in waves 1-3

**Total**: 16 subagent dispatches across 4 waves. At ~30min each = ~8 hours wall-clock serially; waves 1+2+4 are parallel, so realistically 2-3 hours wall-clock if no stalls.

---

## 4. Phase 3 — README + Docs Refresh

### 4.1 Documentation Gaps Identified

1. **`README.md` (407 lines)** — has 6+ stale `edf_collector.py` references (lines 19, 70, 95, 98, 114, 191, 251). All CLI command examples + programmatic usage section reference the deleted monolith. Predates the modularization refactor entirely.
2. **No `docs/` directory** — no architecture documentation anywhere.
3. **`edf_collector.py` config-path reference** — README line 19 says `~/.edf_collector/config.json` but the package is now `edf_bill_fetcher` — config file location may have moved (verify in code).

### 4.2 Refresh Scope

**README.md update** — replace 6 stale `edf_collector` references with canonical `edf_bill_fetcher.<module>` paths. Specific updates:
- CLI examples (lines 70, 95, 98): `python -m edf_bill_fetcher.io.cli` or whatever the new entry point is per `pyproject.toml [project.scripts]`
- Programmatic Usage (lines 114, 251): `from edf_bill_fetcher.collectors import EvidenceEngine`, `from edf_bill_fetcher.io.writers.export import export_to_excel`, etc. — verify each import against the post-refactor canonical homes in `/tmp/edc_canonical_mapping.json`
- Section layout (line 175+): verify `REPORT_SECTIONS` lives in `ui/app.py:ReportOptionsDialog` per the post-refactor structure
- "Adding a new section" example (lines 188-191): rewrite to point at `ui/app.py`
- Add a **Contributing** section pointing at the new package layout (`edf_bill_fetcher/{collectors,helpers,io,models,processors,ui,writers}/`) so new contributors don't get lost

**New `docs/ARCHITECTURE.md`** — high-level overview:
- Package map (7 top-level submodules + io/ sub-packages)
- Hexagonal layering rules (helpers/ stdlib-only, processors/ stdlib+DataFrame, io/ framework imports)
- PEP 562 shim pattern explanation (why io/writers/__init__.py and writers/__init__.py use lazy `__getattr__`)
- Dual public API: flat `from edf_bill_fetcher import X` AND submodule-scoped `from edf_bill_fetcher.processors.matching import infer_contracts`

**New `docs/COVERAGE.md`**:
- Coverage measurement protocol
- 90% floor gate (CI-enforced via `coverage report --fail-under=90`)
- `# pragma: no cover` policy (link to spec section 2.2)
- How to extend coverage (add new test files, run `coverage run -m pytest --cov=edf_bill_fetcher`, then `coverage report --show-missing`)
- Baseline measurement file location (`docs/coverage/2026-08-02-baseline.txt`)

**New `docs/DEVELOPMENT.md`**:
- How to run tests (`pytest` with `pytest-xvfb` auto-active)
- How to add a new writer (canonical home is `io/writers/<name>.py`, add re-export shim in `io/writers/__init__.py`, document `__all__` entry)
- How to add a new processor (canonical home is `processors/<name>.py`)
- ruff/mypy commands (`ruff check .`, `mypy edf_bill_fetcher`)
- PEP 257 D-rule relaxation rationale (link to project memory #256 post-refactor audit)

### 4.3 Out of Scope for Phase 3

- Not rewriting `edf_report.py` or `edf_report_docx.py` — those were out-of-scope per rule #84
- Not documenting `scratch/` scripts — they're dev scratch, not API surface
- Not adding API-reference docs — separate scope (potentially via `sphinx-apidoc` in a future spec)
- Not auto-generating docs from docstrings — same as above

---

## 5. Risk + Non-Goals + Out-of-Scope

### 5.1 Risks

1. **Mock-brittleness in `ui/app.py`** — testing GUI behavior via `unittest.mock.patch` against `tkinter.filedialog.*` creates tests that pass when mocks are in place but don't catch real Tk integration issues.
   - **Mitigation**: each mock returns a synthetic value; the handler's state-mutation is the actual assertion target — never assert mock-call-counts, only assert state changes.

2. **Subagent stalls** — the 30-min-quiet pattern has killed prior coverage/refactor work (rules #222, #226).
   - **Mitigation**: pre-extract every target module's symbols + missed-line inventory per rule #213; embed concrete test patterns in subagent briefs; give each subagent a single-file scope.

3. **Tk fixturization introduces flakiness** — Xvfb may behave poorly under heavy parallel test runs.
   - **Mitigation**: add `xvfb_width=1280, xvfb_height=1024, xvfb_colordepth=24` to `pyproject.toml [tool.pytest.ini_options]` to give Xvfb more headroom; if still flaky, gate GUI tests behind a `@pytest.mark.gui` mark and mark them serial-only via `pytest-xdist --dist=loadscope`.

4. **Coverage regression over time** — even if we hit 100% once, future code can drop it.
   - **Mitigation**: `coverage report --fail-under=90` in CI locks the floor; the aspiration 100% is not gated post-achievement (it's a one-time push, not a perpetual bar).

5. **The 100% aspiration may demand pragmas that hide real untested branches** — discipline risk.
   - **Mitigation**: every `# pragma: no cover` must be commented with the WHY and reviewed in the spec self-review (section 6 of this doc).

### 5.2 Non-Goals (Explicit)

- Not implementing new features in the modularized package
- Not fixing business-logic bugs in the existing codebase (only the 4+31 xvfb-fixable ones, which got fixed for free with the pytest-xvfb install)
- Not migrating to a different test framework
- Not implementing property-based testing (hypothesis etc.) — keep within `pytest` ecosystem
- Not implementing mutation testing (even though `mutmut` would validate the quality of our coverage — out of scope for this spec)

### 5.3 Out-of-Scope

- Any change to `edf_report.py` or `edf_report_docx.py` business logic (per rule #84)
- Updating `scratch/` scripts (dev scratch, not API)
- The full PEP 8/257/20 retroactive compliance pass (project memory #256 — that's a separate post-refactor goal)
- Coverage of `tests/` files themselves (we measure `--source=edf_bill_fetcher` only; test-file coverage isn't a goal)
- Removing the PEP 562 `__getattr__` shims — they're load-bearing for the dual-public-API contract per project memory #220

---

## 6. Spec Self-Review Checklist

(to be run after spec is written, before user review)

- [ ] No "TBD", "TODO", or incomplete sections
- [ ] Internal consistency: do the coverage targets (90/100) match the pragma policy in section 2.2?
- [ ] Scope check: is this focused enough for a single implementation plan, or does it need decomposition into multiple plans (one per phase)?
- [ ] Ambiguity check: could any requirement be interpreted two different ways? Specifically, the `# pragma: no cover` policy language in section 2.2 must be unambiguous about what's allowed vs forbidden.
- [ ] Subagent briefs (will be defined in the implementation plan, not here) must include the codegraph MCP instruction per project memory #248.

---

## 7. Plan Target End-State

After this spec is implemented:
- `pytest-xvfb` is in `[dev]` extras in `pyproject.toml`
- `coverage` is in `[dev]` extras in `pyproject.toml`
- `pyproject.toml` has `[tool.coverage.run]` and `[tool.coverage.report]` sections
- `coverage report --fail-under=90` is enforced
- A new commit on branch `refactor` (after tag `modularization-complete`) delivers all test additions + README/docs refresh + pragma placements
- The 90% floor is achieved and proven via `coverage report`
- The realistic ceiling (96.3% ± 2%) is achieved and proven; remaining gaps are documented in `docs/coverage/<date>-residual-gap.md` if any
- An honest 100% via pragmas is the aspiration; wherever pragmas aren't justified, the line is genuinely tested
- `pyproject.toml` xvfb config in `[tool.pytest.ini_options]` block sets `xvfb_width=1280, xvfb_height=1024, xvfb_colordepth=24`
- README is updated with no `edf_collector` references remaining (grep returns zero hits outside docstring migration notes)
- `docs/ARCHITECTURE.md`, `docs/COVERAGE.md`, `docs/DEVELOPMENT.md` are committed

This spec is deliberately non-contractual about subagent dispatch counts, wave assignment, and per-module coverage deltas — those are implementation-plan concerns. The spec owns the WHAT and WHY; the plan owns the HOW and WHEN.

---

**Approval request**: this design locks the 90/100 targets, the strict pragma policy, the 3-phase sequencing (xvfb-already-installed → coverage gap modules → docs refresh), the no-mock-brittleness quality bar, and the out-of-scope boundaries. Once you approve, the spec is committed and we move to writing-plans.
