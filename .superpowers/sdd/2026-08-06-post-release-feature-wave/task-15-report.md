# Task 15 Report — CLI `--diff` subcommand (Wave 6g)

**Status:** DONE

## What was done

- `edf_bill_fetcher/io/cli.py`: new `run_cli_diff(args)` — `--diff OLD NEW`
  reads two records.json files (bare list or `--records-json` wrapper shape),
  calls `processors.run_diff.diff_records`, prints counts + per-row summary
  lines (`+ ADDED` / `- REMOVED` / `~ CHANGED ... [field: old → new]`);
  `--diff-output PATH` writes the workbook. `main()` gained a `--diff`/`-d`
  dispatch branch and the no-tkinter help text mentions it.
- `edf_bill_fetcher/io/writers/diff.py` (new): `write_diff_workbook(diff, path)`
  — three sheets **Added Records** / **Removed Records** / **Changed Records**.
  Changed sheet uses paired `<field> (old)` / `<field> (new)` columns plus a
  trailing `Changed Fields` column summarising deltas as `field: old → new`
  (amounts rendered `£x.xx`). Exported via `io/writers/__init__.py`.
- `tests/test_io_cli_diff.py` (new, TDD — RED verified before implementation):
  5 tests covering the summary counts + per-row lines, the 3-sheet workbook
  structure, clean exit-1 on a missing file, wrapper-JSON unwrapping, and
  `main()` dispatch.
- `README.md`: one CLI usage snippet for `--diff`.

## Verification (conda env `edf-bill-fetcher`, Python 3.11)

- Focused suite (diff + argv + run_diff tests): **52 passed**
- Full suite: **1547 passed, 9 skipped, 0 failed** (skips are pre-existing
  environment/branch skips: pypff/scipy/statsmodels slots, missing Ombudsman
  scratch PDFs)
- `ruff check .` → All checks passed
- `ruff format --check .` → 211 files already formatted
- `mypy . --exclude scratch/` → Success: no issues found in 206 source files
- End-to-end smoke via `python main.py --diff ...` → correct summary lines,
  workbook sheets `['Added Records', 'Removed Records', 'Changed Records']`,
  exit 0; missing file → `ERROR: ...` exit 1

## Notes / concerns

- None blocking. `--pdf-report` / `--docx-report` / `--html-report` behaviour
  untouched.
- Evidence file written to
  `.omo/evidence/post-release-feature-wave/task-6g2-2026-08-06-post-release-feature-wave.txt`
  (git-ignored, not staged).
