# Test Coverage to 100% + README/Docs Refresh — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Drive test coverage of `edf_bill_fetcher/` package from current 71% to 90% floor (CI-enforced) with 100% aspiration; refresh README and add 3 new docs/ files for the post-refactor package structure.

**Architecture:** Three sequential phases. Phase 1 configures test infrastructure (`pytest-xvfb` and `coverage` in dev extras + `pyproject.toml` coverage config). Phase 2 closes the coverage gap via 4 parallel subagent waves (16 module-by-module test additions). Phase 3 refreshes README and adds `docs/ARCHITECTURE.md`, `docs/COVERAGE.md`, `docs/DEVELOPMENT.md`. Each phase ends with independently-verifiable gates.

**Tech Stack:** Python 3.12 (edf conda env at `/home/matthew/anaconda3/envs/edf/bin/python`), pytest 9.1.1, pytest-xvfb 3.1.1, coverage 7.15.2, ruff (with D rules for PEP 257), mypy, xvfb system binary.

## Global Constraints

- All Python work uses `/home/matthew/anaconda3/envs/edf/bin/python` — never the host python (project memory #99).
- All work on branch `refactor` (after tag `modularization-complete` at commit `98cd0f2`). No commits to `dev` during this work (rule #112).
- Each new test file must satisfy PEP 257 D rules (project memory #255) — module docstring, test-class docstring, test-function docstring. Tests have ALL D rules relaxed per `[tool.ruff.lint.per-file-ignores]` in `pyproject.toml`, so docstrings are OPTIONAL on tests but REQUIRED on production code.
- No `# type: ignore` or `as any` as a substitute for correct types (rule #137).
- Subagent briefs MUST instruct subagents to use the `codegraph_codegraph_explore` MCP tool when reading code — the repo is indexed at `/home/matthew/ai/opencode/edf-bill-fetcher` (project memory #248).
- Subagent dispatch uses the pre-extract pattern (rule #213): brief contains missed-line inventory + concrete test patterns; subagent does NOT explore, only writes.
- `# pragma: no cover` is allowed ONLY for genuinely unreachable code with a one-line `# WHY:` comment. Forbidden uses per spec section 2.2.
- Coverage measured via `coverage run --source=edf_bill_fetcher --branch -m pytest` followed by `coverage report`.
- xvfb is auto-active via `pytest-xvfb` plugin — no flag needed, just install.
- The 4+31 "pre-existing failures" are gone — pytest-xvfb install fixed them. Starting state: 786 passed / 0 failed / 0 errors / 7 skipped; coverage 71% (4347/5868 covered).

---

## Phase 1: Test Infrastructure Configuration

### Task 1.1: Add pytest-xvfb and coverage to dev extras + xvfb config in pyproject.toml

**Files:**
- Modify: `pyproject.toml` — `[project.optional-dependencies]` dev extras section + `[tool.pytest.ini_options]` + new `[tool.coverage.run]` + `[tool.coverage.report]`

**Interfaces:**
- Consumes: nothing
- Produces: `pytest-xvfb` and `coverage` as locked dev-deps; CI can run `coverage report --fail-under=90`

- [ ] **Step 1: Read current pyproject.toml dependencies block**

Run: `read /home/matthew/ai/opencode/edf-bill-fetcher/pyproject.toml`

Locate `[project.optional-dependencies]` block, the `[tool.pytest.ini_options]` block, and check whether a `[tool.coverage.*]` section already exists.

- [ ] **Step 2: Add pytest-xvfb + coverage to dev extras**

In `pyproject.toml`, find the `[project.optional-dependencies]` `dev` list. Add these two entries (in alphabetical order if the list is sorted):

```toml
"coverage>=7.0",
"pytest-xvfb>=3.0",
```

If `dev` does not exist, create it:
```toml
[project.optional-dependencies]
dev = [
    "coverage>=7.0",
    "pytest-xvfb>=3.0",
    # ... existing entries ...
]
```

- [ ] **Step 3: Add xvfb config to [tool.pytest.ini_options]**

Add inside `[tool.pytest.ini_options]`:
```toml
xvfb_width = 1280
xvfb_height = 1024
xvfb_colordepth = 24
```

These give Xvfb more headroom under parallel test runs (spec risk #3 mitigation).

- [ ] **Step 4: Add [tool.coverage.run] and [tool.coverage.report] sections**

Append at the end of `pyproject.toml` (or after the existing `[tool.ruff]` section if that's at the end):

```toml
[tool.coverage.run]
source = ["edf_bill_fetcher"]
branch = true

[tool.coverage.report]
fail_under = 90
show_missing = true
skip_covered = true
```

- [ ] **Step 5: Verify pytest still runs and xvfb config is loaded**

Run: `/home/matthew/anaconda3/envs/edf/bin/python -m pytest --co -q tests/test_output_folder_var.py 2>&1 | tail -5`

Expected: collection succeeds, no errors. (Xvfb auto-activates at runtime, not collection time.)

- [ ] **Step 6: Verify coverage config loads**

Run: `/home/matthew/anaconda3/envs/edf/bin/python -m coverage run -m pytest --no-header -q tests/test_back_billing_sheet.py 2>&1 | tail -3 && /home/matthew/anaconda3/envs/edf/bin/python -m coverage report 2>&1 | tail -5`

Expected: coverage runs with `edf_bill_fetcher` as source, branch=true; report shows at least one module.

- [ ] **Step 7: Commit**

```bash
git add pyproject.toml
git commit -m "chore(deps): add pytest-xvfb + coverage to dev extras; configure xvfb + coverage

- pytest-xvfb auto-activates on headless Linux environments
- xvfb_width=1280, xvfb_height=1024, xvfb_colordepth=24 for
  parallel-test headroom (spec risk #3 mitigation)
- coverage --fail-under=90 enforces 90% floor per spec section 2.1"
```

---

## Phase 2: Coverage Gap Closure (15 modules, 4 waves)

### Task 2.1: Wave 1 — Tier A parallel subagent dispatch (6 subagents)

**Files:**
- Create: `tests/test_io_reporters_shim.py`
- Create: `tests/test_io_writers_init_shim.py`
- Create: `tests/test_processors_extraction.py`
- Create: `tests/test_writers_statistical_branches.py` (extend existing if present)
- Create: `tests/test_processors_forecasting_branches.py`
- Create: `tests/test_processors_detection_branches.py`

**Interfaces:**
- Consumes: existing `edf_bill_fetcher/` package (read-only)
- Produces: 6 new test files, ~+264 covered statements, coverage to ~78.6%

**Pre-extract work (during prep, NOT in subagent brief):** For each target module, inventory the missed-line numbers via `coverage report --show-missing <module_path>` and prepare a brief containing:
1. The missed-line inventory (concrete line numbers)
2. The module's public API surface (function/class signatures)
3. Concrete test patterns to follow (synthetic DataFrame construction, expected branches)
4. The `codegraph_codegraph_explore` instruction per project memory #248

- [ ] **Step 1: Generate missed-line inventory for Wave 1 modules**

Run:
```bash
/home/matthew/anaconda3/envs/edf/bin/python -m coverage run --source=edf_bill_fetcher --branch -m pytest --no-header -q 2>&1 | tail -3
/home/matthew/anaconda3/envs/edf/bin/python -m coverage report --show-missing --include="edf_bill_fetcher/io/reporters/*,edf_bill_fetcher/io/writers/__init__.py,edf_bill_fetcher/processors/extraction.py,edf_bill_fetcher/io/writers/statistical.py,edf_bill_fetcher/processors/forecasting.py,edf_bill_fetcher/processors/detection.py" 2>&1 > /tmp/wave1_inventory.txt
wc -l /tmp/wave1_inventory.txt
```

Expected: inventory file with concrete missed line numbers per module.

- [ ] **Step 2: Pre-extract module public-API signatures**

For each Wave 1 target module, call `codegraph_codegraph_explore` with the module name and capture the returned source signatures. For example, for `processors/extraction.py`:

```
codegraph_codegraph_explore(query="processors/extraction.py public functions")
```

Save the extracted signatures to `/tmp/wave1_apis.txt` for inclusion in subagent briefs.

- [ ] **Step 3: Dispatch 6 parallel subagents (one per Tier A module)**

Each subagent: `task(category="quick", load_skills=[], run_in_background=true)`. **WAIT for all 6 to complete before Step 4 — per rule #91, do not poll; the system sends `<system-reminder>` on completion.**

Brief template (per subagent, scoped to one module):

```
TASK: Add test coverage to <module_path> for missed lines listed in the inventory.

EXPECTED OUTCOME: New test file at <test_file_path> with passing tests that cover the missed lines enumerated below. Coverage for the module rises to ≥95%. All tests pass under `pytest`. New test file satisfies ruff + mypy gates.

REQUIRED TOOLS: edit, write, bash, codegraph_codegraph_explore (for verifying symbol sources).

MUST DO:
- Use codegraph_codegraph_explore to verify the exact source of any symbol you import — the repo is indexed.
- Cover EVERY line listed in the missed-line inventory below. Each missed line must be exercised by at least one test.
- Use synthetic DataFrames as inputs — no real file I/O, no mocking unless explicitly required.
- For PEP 562 shim modules: test that `getattr(module, name)` for each name in `__all__` succeeds AND that `getattr(module, "nonexistent_attr")` raises AttributeError (covers the `__getattr__` fallback branch).
- Each new test file has a module docstring + each test function has a docstring (PEP 257).
- Commit the new test file with message: `test(<module>): add <purpose> tests for coverage`

MUST NOT DO:
- Do NOT mock at levels finer than the public API boundary (per spec section 2.3 — no mock-brittleness).
- Do NOT use `# type: ignore` or `as any` (rule #137).
- Do NOT modify the production code in <module_path> — only add tests.
- Do NOT run mock-patch assertions on call counts; assert only state changes.

CONTEXT:
- Module: <module_path>
- Test file: <test_file_path>
- Missed line inventory: <paste missed lines from /tmp/wave1_inventory.txt for this module>
- Public API signatures: <paste from /tmp/wave1_apis.txt for this module>
- Coverage currently: <X% (Y missed out of Z stmts)>, target ≥95%
- Codebase root: /home/matthew/ai/opencode/edf-bill-fetcher
- Python: /home/matthew/anaconda3/envs/edf/bin/python
- Run tests: /home/matthew/anaconda3/envs/edf/bin/python -m pytest <test_file_path> -v
- Coverage: /home/matthew/anaconda3/envs/edf/bin/python -m coverage run --branch --include=<module_path> -m pytest <test_file_path> && /home/matthew/anaconda3/envs/edf/bin/python -m coverage report --show-missing --include=<module_path>
```

Expected: 6 subagent tasks launched (`bg_*` IDs). End response and await completion notifications per rule #91.

- [ ] **Step 4: Collect subagent results and verify coverage delta**

For each completed subagent:
1. `background_output(task_id="bg_<id>")`
2. Verify commit landed: `git log --oneline -5 | grep "test(<module>)"`
3. Run the local coverage check on the module:
   ```bash
   /home/matthew/anaconda3/envs/edf/bin/python -m coverage run --source=edf_bill_fetcher --branch -m pytest <test_file_path> --no-header -q 2>&1 | tail -3
   /home/matthew/anaconda3/envs/edf/bin/python -m coverage report --show-missing --include="<module_path>" 2>&1 | tail -3
   ```
4. Verify module coverage ≥95% (or pragma'd remainder documented). If below 95%, dispatch fix-up via `task(task_id="ses_<id>", ...)` continuation per rule #162.

- [ ] **Step 5: Run global coverage + gate check**

```bash
/home/matthew/anaconda3/envs/edf/bin/python -m coverage run --source=edf_bill_fetcher --branch -m pytest --no-header -q 2>&1 | tail -3
/home/matthew/anaconda3/envs/edf/bin/python -m coverage report 2>&1 | tail -5
```

Expected: TOTAL coverage ≥78.6% (4347 pre + ~264 Wave 1 = 4611/5868). Run ruff + mypy gates:
```bash
/home/matthew/anaconda3/envs/edf/bin/python -m ruff check . 2>&1 | tail -3
/home/matthew/anaconda3/envs/edf/bin/python -m mypy edf_bill_fetcher 2>&1 | tail -3
```

Expected: All gates green. No regressions.

- [ ] **Step 6: Tag Wave 1 checkpoint**

```bash
git tag coverage-wave-1-complete
```

### Task 2.2: Wave 2 — Tier B parallel subagent dispatch (7 subagents)

**Files:**
- Create: `tests/test_io_cli_argv.py`
- Modify: `tests/fixtures/pst_attachment_fixture.py` + Create: `tests/test_io_adapters_pst_branches.py`
- Create: `tests/test_writers_rebilling_branches.py`
- Create: `tests/test_writers_meter_branches.py`
- Create: `tests/test_writers_back_billing_branches.py`
- Create: `tests/test_processors_analysis_branches.py`
- Create: `tests/test_collectors_engine_error_paths.py`

**Interfaces:**
- Consumes: existing `edf_bill_fetcher/` package
- Produces: 7 new/modified test files, ~+634 covered statements, coverage to ~89.4%

The brief for each subagent is the same template as Wave 1, but with these module-specific overrides:

| Module | Specific instructions |
|--------|----------------------|
| `io/cli.py` | Use `monkeypatch.setattr("sys.argv", [...])` + `capsys` for stdout. Cover `main()`, `run_cli_extract`, `run_cli_pdf_report`, `run_cli_docx_report`. Stub `EvidenceEngine` with a synthetic that returns pre-computed records (no real I/O). |
| `io/adapters/pst.py` | EXTEND `tests/fixtures/pst_attachment_fixture.py` synthetic `pypff` shape to cover: missing `PR_ATTACH_LONG_FILENAME` (falls back to short), missing both (falls back to `Attachment_N.pdf`), corrupt record set (AttributeError swallow), multiple attachments. NO real `.pst` file — synthetic only. |
| `io/writers/rebilling.py` | Test `write_rebilling_sheet` with empty rebilling list, single entry, multiple entries, each flag combination. Synthetic DataFrames via `pandas.DataFrame(...)`. |
| `io/writers/meter.py` | Test empty readings, single contract, multi-contract, missing-period edge cases. Synthetic DataFrames. |
| `io/writers/back_billing.py` | Test no events, single event, multi-event, with/without matched-EDF context. |
| `processors/analysis.py` | Test `compute_dispute_flags` paths for each flag. Synthetic engine records as input. |
| `collectors/engine.py` | Mock `pdfplumber.open` to raise various exceptions (`FileNotFoundError`, `pd.errors.ParserError`, `IOError`) and assert the error path is taken without crash. Each `process_*` method needs happy + error path coverage. |

- [ ] **Step 1: Generate Wave 2 missed-line inventory**

Same pattern as Wave 1 Step 1, but for the Wave 2 module list. Save to `/tmp/wave2_inventory.txt`.

- [ ] **Step 2: Pre-extract Wave 2 module APIs**

Same pattern as Wave 1 Step 2. Save to `/tmp/wave2_apis.txt`.

- [ ] **Step 3: Dispatch 7 parallel subagents**

Same template, one subagent per Tier B module. All 7 dispatched simultaneously with `run_in_background=true`. End response and await completion notifications per rule #91.

- [ ] **Step 4: Collect + verify per-subagent**

Same as Wave 1 Step 4. Each module must rise to ≥95% (or have documented pragma remainder). Coverage for `io/cli.py` and `io/adapters/pst.py` may require fix-up iterations — use `task(task_id="ses_<id>", ...)` continuation per rule #162 if stalls occur.

- [ ] **Step 5: Global gate check — should now be ~89.4%**

```bash
/home/matthew/anaconda3/envs/edf/bin/python -m coverage run --source=edf_bill_fetcher --branch -m pytest --no-header -q 2>&1 | tail -3
/home/matthew/anaconda3/envs/edf/bin/python -m coverage report 2>&1 | tail -5
```

Expected: TOTAL coverage ~89.4%. Below 90% floor — Tier C is required to hit the gate.

- [ ] **Step 6: Tag Wave 2 checkpoint**

```bash
git tag coverage-wave-2-complete
```

### Task 2.3: Wave 3 — Tier C ui/app.py (1 critical subagent)

**Files:**
- Create: `tests/test_ui_app_dialog_handlers.py`

**Interfaces:**
- Consumes: existing `edf_bill_fetcher/ui/app.py` App + ReportOptionsDialog classes
- Produces: ~+293 covered statements in `ui/app.py` (43% → ~95%)

- [ ] **Step 1: Inventory missed lines in ui/app.py**

```bash
/home/matthew/anaconda3/envs/edf/bin/python -m coverage run --source=edf_bill_fetcher --branch -m pytest --no-header -q 2>&1 | tail -3
/home/matthew/anaconda3/envs/edf/bin/python -m coverage report --show-missing --include="edf_bill_fetcher/ui/app.py" 2>&1 > /tmp/wave3_app_inventory.txt
wc -l /tmp/wave3_app_inventory.txt
```

- [ ] **Step 2: Pre-extract App + ReportOptionsDialog behavior map**

Call `codegraph_codegraph_explore(query="ui/app.py App class methods dialog handlers ReportOptionsDialog")` and capture:
1. All `filedialog.*`, `simpledialog.*`, `messagebox.*` call sites (these are the boundary-mock targets)
2. All `_open_*` handler methods
3. EXTRACT button workflow states (`EXTRACT TO EXCEL → Cancel → Cancelling... → EXTRACT TO EXCEL`)
4. State-mutation assertions available (which App instance attrs change when handler runs)

Save to `/tmp/wave3_app_apis.txt`.

- [ ] **Step 3: Dispatch single critical subagent (no parallel — Tk is sequential by nature)**

```
task(category="deep", load_skills=[], run_in_background=true)
```

Brief contains:
- Missed-line inventory from `/tmp/wave3_app_inventory.txt`
- The behavioral map from `/tmp/wave3_app_apis.txt`
- The boundary-mock strategy per spec section 2.3:
  ```
  For each modal handler (e.g., _open_output_folder_picker):
  - Patch ONLY the filedialog.* call: `with unittest.mock.patch("tkinter.filedialog.askdirectory", return_value="/tmp/test_pick") as mock_pick:`
  - Invoke the handler: `app._open_output_folder_picker()`
  - Assert STATE change: `assert app.output_folder_var.get() == "/tmp/test_pick"`
  - NEVER assert mock_pick.call_count or call_args — assert only state mutations.
  ```
- The EXTRACT workflow tests:
  ```
  - Test idle state: app.extract_button["text"] == "EXTRACT TO EXCEL"
  - Test mid-extract: simulate StartExtract event → text becomes "Cancel"
  - Test cancel-requested: simulate Cancel event → text becomes "Cancelling..."
  - Test post-cancel: simulate ExtractComplete with was_cancelled=True → text returns to "EXTRACT TO EXCEL"
  ```

Expected: subagent runs in background; end response and await completion notification.

- [ ] **Step 4: Collect + verify**

```bash
/home/matthew/anaconda3/envs/edf/bin/python -m coverage run --source=edf_bill_fetcher --branch -m pytest tests/test_ui_app_dialog_handlers.py --no-header -q 2>&1 | tail -3
/home/matthew/anaconda3/envs/edf/bin/python -m coverage report --show-missing --include="edf_bill_fetcher/ui/app.py" 2>&1 | tail -3
```

Expected: `ui/app.py` coverage ≥95% (was 43%). Tk-specific branches that resist mocking must be `# pragma: no cover` with WHY comments per spec section 2.2.

- [ ] **Step 5: Global gate check**

```bash
/home/matthew/anaconda3/envs/edf/bin/python -m coverage run --source=edf_bill_fetcher --branch -m pytest --no-header -q 2>&1 | tail -3
/home/matthew/anaconda3/envs/edf/bin/python -m coverage report 2>&1 | tail -5
```

Expected: ~95%+ overall. If below 95%, Wave 4 (pragmas + helpers) is required.

- [ ] **Step 6: Tag Wave 3 checkpoint**

```bash
git tag coverage-wave-3-complete
```

### Task 2.4: Wave 4 — writers/_helpers.py + final pragma audit (2 parallel subagents)

**Files:**
- Create: `tests/test_writers_helpers_branches.py`
- Modify: multiple `edf_bill_fetcher/*.py` files — pragma additions only

**Interfaces:**
- Consumes: residual uncovered code identified by `coverage report --show-missing`
- Produces: ~+100 covered `writers/_helpers.py` statements + `# pragma: no cover` on genuinely unreachable lines across all modules; final coverage reportable as ~100%

- [ ] **Step 1: Generate residual gap inventory**

```bash
/home/matthew/anaconda3/envs/edf/bin/python -m coverage run --source=edf_bill_fetcher --branch -m pytest --no-header -q 2>&1 | tail -3
/home/matthew/anaconda3/envs/edf/bin/python -m coverage report --show-missing 2>&1 > /tmp/wave4_residual.txt
wc -l /tmp/wave4_residual.txt
```

- [ ] **Step 2: Triage residual into testable vs pragma-able**

For each missed line in `/tmp/wave4_residual.txt`:
- If the line is a real branch that COULD be tested (e.g., `else:` for an error path) → goes into the `_helpers.py` test subagent brief
- If the line is genuinely unreachable (`if TYPE_CHECKING:`, `except ImportError` for hard-required imports, `if sys.platform == "darwin":` on Linux CI) → goes into the pragma audit subagent brief

Document each decision in `/tmp/wave4_triage.md`.

- [ ] **Step 3: Dispatch 2 parallel subagents**

Subagent A: writes `tests/test_writers_helpers_branches.py` covering the testable residuals in `writers/_helpers.py`.

Subagent B: applies `# pragma: no cover` with one-line `# WHY:` comments to genuinely unreachable lines across all modules per the triage list. Returns a summary of every pragma added with file:line and justification.

End response and await completion notifications.

- [ ] **Step 4: Verify + run final gate**

```bash
/home/matthew/anaconda3/envs/edf/bin/python -m coverage run --source=edf_bill_fetcher --branch -m pytest --no-header -q 2>&1 | tail -3
/home/matthew/anaconda3/envs/edf/bin/python -m coverage report 2>&1 | tail -5
```

Expected: TOTAL coverage ≥96%. If at 96% with the remaining 4% as documented pragmas, declare aspiration complete (honest 100% via pragmas + tested remainder).

- [ ] **Step 5: Audit pragma list — each must have WHY comment**

For every `# pragma: no cover` added by Subagent B, grep the file:
```bash
grep -rn "pragma: no cover" edf_bill_fetcher/ | head -50
```

For each hit, verify the next column (`# WHY:` or similar inline comment) explains the unreachability. Reject any pragma without a justification.

- [ ] **Step 6: Tag coverage-complete**

```bash
git tag coverage-complete
```

---

## Phase 3: README + Docs Refresh

### Task 3.1: Refresh README.md — replace stale edf_collector references

**Files:**
- Modify: `README.md` — lines 19, 70, 95, 98, 114, 191, 251 + scan for any others

**Interfaces:**
- Consumes: post-refactor canonical mapping at `/tmp/edc_canonical_mapping.json` (built during Task 8 Phase 7a of the refactor)
- Produces: README with zero `edf_collector` references outside docstring-migration notes; programmatic usage + CLI examples updated to canonical `edf_bill_fetcher.<module>` paths

- [ ] **Step 1: Generate current canonical mapping (refresh)**

```bash
# Re-extract the canonical mapping from the live package
/home/matthew/anaconda3/envs/edf/bin/python -c "
import json, importlib
# Re-use mapping file from refactor if exists, else regenerate
import os
if os.path.exists('/tmp/edc_canonical_mapping.json'):
    mapping = json.load(open('/tmp/edc_canonical_mapping.json'))
    print(f'Mapping loaded: {len(mapping)} names')
else:
    print('Mapping missing — regenerate via AST scan of edf_bill_fetcher/')
" 2>&1
```

- [ ] **Step 2: Grep README for every stale `edf_collector` reference**

```bash
grep -n "edf_collector\b\|edf_collector\.py\|from edf_collector\|import edf_collector\|~/.edf_collector" README.md
```

Document every line that needs updating.

- [ ] **Step 3: Update CLI examples (lines 70, 95, 98)**

Replace `python edf_collector.py` with the new entry-point form. Verify the entry point by checking `pyproject.toml [project.scripts]`:

```bash
grep -A 5 "\[project.scripts\]" pyproject.toml
```

If `[project.scripts]` says `edf-collector = "edf_bill_fetcher.io.cli:main"`, then CLI examples become `edf-collector --pdf-report ...`. If there's no script entry, fall back to `python -m edf_bill_fetcher.io.cli --pdf-report ...`.

- [ ] **Step 4: Update Programmatic Usage (lines 114, 251)**

Replace `from edf_collector import ...` with canonical imports:
- `from edf_collector import EvidenceEngine` → `from edf_bill_fetcher.collectors import EvidenceEngine`
- `from edf_collector import export_to_excel` → `from edf_bill_fetcher.io.writers.export import export_to_excel`
- `from edf_collector import (run_analysers, ...)` → split the tuple into individual canonical imports

Verify EACH import against `/tmp/edc_canonical_mapping.json` before writing.

- [ ] **Step 5: Update "Adding a new section" example (lines 188-191)**

Replace `ReportOptionsDialog.SECTIONS in edf_collector.py` with the new canonical home: `ReportOptionsDialog.SECTIONS in edf_bill_fetcher/ui/app.py`.

- [ ] **Step 6: Update config-path reference (line 19)**

Verify whether the config file location moved during refactor. Run:
```bash
grep -rn "\.edf_collector\|\.edf_bill_fetcher\|CONFIG_PATH\|config_dir\|config_path" edf_bill_fetcher/ 2>&1 | head -10
```

If config still lives at `~/.edf_collector/config.json`, leave line 19 as-is (the filename didn't move). If it moved to `~/.edf_bill_fetcher/config.json`, update line 19 accordingly.

- [ ] **Step 7: Add Contributing section pointing at new package layout**

Append a "## Contributing" section to README pointing at `edf_bill_fetcher/{collectors,helpers,io,models,processors,ui,writers}/` so new contributors don't get lost. Refer readers to `docs/ARCHITECTURE.md` (created in Task 3.2) for the full package map.

- [ ] **Step 8: Verify zero stale references remain**

```bash
grep -n "edf_collector\b\|edf_collector\.py\|from edf_collector\|import edf_collector\|~/.edf_collector" README.md | grep -v "docstring migration notes\|memorial\|historical"
```

Expected: zero hits (or only hits inside an explicit "Historical" / "Migration notes" section you intentionally preserved).

- [ ] **Step 9: Commit**

```bash
git add README.md
git commit -m "docs(readme): refresh stale edf_collector references for post-refactor package

- CLI examples now use the new entry point (per pyproject.toml [project.scripts])
- Programmatic Usage section imports canonical edf_bill_fetcher.<module> paths
- 'Adding a new section' example updated for ui/app.py ReportOptionsDialog canonical home
- New Contributing section points at edf_bill_fetcher/{collectors,helpers,io,models,processors,ui,writers}/ package layout"
```

### Task 3.2: Create docs/ARCHITECTURE.md

**Files:**
- Create: `docs/ARCHITECTURE.md`

- [ ] **Step 1: Document the package map**

Write `docs/ARCHITECTURE.md` with these sections:
- **Package map** — tree of `edf_bill_fetcher/` with one-line responsibility per top-level submodule (collectors/, helpers/, io/, models/, processors/, ui/, writers/) plus io/ sub-packages (adapters/, reporters/, writers/)
- **Hexagonal layering rules** — helpers/ stdlib-only (no framework imports); processors/ stdlib + pandas DataFrame as args (no framework imports at module scope); io/ allowed framework imports (openpyxl, reportlab, tkinter, pickle); collectors/ orchestration layer (imports processors + io as needed); ui/ Tkinter; models/ plain data classes
- **PEP 562 shim pattern** — why `io/writers/__init__.py` and `writers/__init__.py` use module-level `__getattr__` (lazy resolution avoids circular imports). Link to project memory #220 + #219.
- **Dual public API** — flat `from edf_bill_fetcher import X` AND submodule-scoped `from edf_bill_fetcher.processors.matching import infer_contracts`. Both forms are supported and tested.

- [ ] **Step 2: Commit**

```bash
git add docs/ARCHITECTURE.md
git commit -m "docs(architecture): document post-refactor package map + hexagonal layering + PEP 562 shims"
```

### Task 3.3: Create docs/COVERAGE.md

**Files:**
- Create: `docs/COVERAGE.md`
- Create: `docs/coverage/2026-08-02-baseline.txt` (committed baseline measurement)

- [ ] **Step 1: Document the coverage measurement protocol**

Write `docs/COVERAGE.md` with these sections:
- **Coverage gate** — `coverage report --fail-under=90` enforces 90% floor in CI. Spec section 2.1.
- **`# pragma: no cover` policy** — link to spec section 2.2 verbatim. Each pragma needs one-line `# WHY:` justification. Forbidden uses listed.
- **How to run coverage locally**:
  ```bash
  /home/matthew/anaconda3/envs/edf/bin/python -m coverage run --source=edf_bill_fetcher --branch -m pytest --no-header -q
  /home/matthew/anaconda3/envs/edf/bin/python -m coverage report --show-missing
  ```
- **How to extend coverage** — add new test files under `tests/test_<module>_<purpose>.py`, run coverage, verify the new module's coverage rises. Each new test file must satisfy PEP 257 D rules per project memory #255 (relaxed for test files — docstrings optional but encouraged).
- **Baseline reference** — `docs/coverage/2026-08-02-baseline.txt` (the measurement taken at commit `98cd0f2` immediately before this plan's execution).

- [ ] **Step 2: Save the baseline measurement**

```bash
mkdir -p docs/coverage
/home/matthew/anaconda3/envs/edf/bin/python -m coverage run --source=edf_bill_fetcher --branch -m pytest --no-header -q 2>&1 | tail -3
/home/matthew/anaconda3/envs/edf/bin/python -m coverage report > docs/coverage/2026-08-02-baseline.txt
```

- [ ] **Step 3: Commit**

```bash
git add docs/COVERAGE.md docs/coverage/2026-08-02-baseline.txt
git commit -m "docs(coverage): document measurement protocol + commit 2026-08-02 baseline"
```

### Task 3.4: Create docs/DEVELOPMENT.md

**Files:**
- Create: `docs/DEVELOPMENT.md`

- [ ] **Step 1: Document the developer workflow**

Write `docs/DEVELOPMENT.md` with these sections:
- **Run tests** — `pytest` (pytest-xvfb auto-active on headless Linux). For specific gates: `ruff check .`, `mypy edf_bill_fetcher`, `coverage report --fail-under=90`.
- **Add a new writer** — canonical home is `edf_bill_fetcher/io/writers/<name>.py`. Add the function, add an `__all__` entry, add a PEP 562 re-export in `io/writers/__init__.py` if the writer should be importable via the flat API.
- **Add a new processor** — canonical home is `edf_bill_fetcher/processors/<name>.py`. Stdlib + pandas only — no framework imports at module scope.
- **PEP 257 D-rule relaxation rationale** — link to project memory #256 (the post-refactor PEP compliance audit goal) and explain the current relaxations in `pyproject.toml [tool.ruff.lint] ignore` list.
- **Coverage discipline** — refer to `docs/COVERAGE.md`.

- [ ] **Step 2: Commit**

```bash
git add docs/DEVELOPMENT.md
git commit -m "docs(development): document test/lint/coverage workflow + contribute patterns"
```

---

## Final Verification

### Task 4.1: Final audit + tag

- [ ] **Step 1: Run all gates fresh**

```bash
/home/matthew/anaconda3/envs/edf/bin/python -m ruff check . 2>&1 | tail -3
/home/matthew/anaconda3/envs/edf/bin/python -m mypy edf_bill_fetcher 2>&1 | tail -3
/home/matthew/anaconda3/envs/edf/bin/python -m coverage run --source=edf_bill_fetcher --branch -m pytest --no-header -q 2>&1 | tail -3
/home/matthew/anaconda3/envs/edf/bin/python -m coverage report --fail-under=90 2>&1 | tail -5
```

Expected: all green. Coverage ≥90% (gate passes). Optional: coverage ≥96% with pragmas (aspiration).

- [ ] **Step 2: Verify zero stale edf_collector references**

```bash
grep -rn "edf_collector" README.md docs/ | head -5
```

Expected: zero hits (or only intentional historical/migration notes).

- [ ] **Step 3: Tag the milestone**

```bash
git tag coverage-and-docs-complete
```

- [ ] **Step 4: Final summary report to user**

Synthesize:
- Coverage achieved: X% (target 90 floor / 100 aspiration / 96.3 realistic ceiling)
- All gates green
- README refresh landed + docs files created
- 16-module test-addition subagent dispatches executed across 4 waves
- Remaining residual gap (if any) documented with `# pragma: no cover` justification list

---

## Self-Review Checklist (run after writing this plan)

- [ ] Spec coverage: every section of `docs/superpowers/specs/2026-08-02-test-coverage-and-docs-refresh-design.md` maps to at least one task in this plan
- [ ] Placeholder scan: no "TBD", "TODO", "implement later", or vague hand-waves
- [ ] Type consistency: function names + module paths used in later tasks match those defined in earlier tasks
- [ ] Each task ends with its own test cycle + commit (per writing-plans skill)
- [ ] Subagent briefs include `codegraph_codegraph_explore` instruction per project memory #248

---

## Execution Handoff

Plan complete and saved to `docs/superpowers/plans/2026-08-02-test-coverage-and-docs-refresh.md`. Two execution options:

**1. Subagent-Driven (recommended)** — I dispatch a fresh subagent per task, review between tasks, fast iteration.

**2. Inline Execution** — Execute tasks in this session using executing-plans, batch execution with checkpoints.

Which approach?
