# Task 13 Report: HTML dispatcher parity + CLI/GUI wiring

**Status: DONE**

## Summary

Extended `tests/test_dispatch_parity.py` with HTML-dispatcher parity tests
(registry coverage, three-format lockstep, all-callable), added the
`--html-report` CLI entry point, and added "HTML" to the GUI Output Format
frame — threaded through both the auto-generate report path and the
LOAD & REPORT path.

## Changes

### Tests (RED first, then GREEN)
- `tests/test_dispatch_parity.py`
  - `_html_dispatcher_keys()` — same AST literal-dict walk used for PDF/DOCX.
  - `test_html_dispatcher_covers_registry`, `test_pdf_docx_and_html_dispatchers_agree`,
    `test_html_dispatchers_all_callable` — three-format lockstep.
  - `TestCliHtmlReportSmoke` — real end-to-end `run_cli_html_report` producing
    an actual `.html` file with the report title inside.
  - `TestGuiHtmlFormatPresence` — builds the real dialog, walks the widget
    tree, asserts an `html`-valued radiobutton exists.
- `tests/test_io_cli_argv.py` — `TestRunCliHtmlReport` (bare list, wrapper
  unwrap, failure exit 1, config, engine-data pickle, exception handler)
  + `main()` `--html-report` dispatch test, mirroring the DOCX class.
- `tests/test_ui_app_dialog_handlers.py` — html-only + both-includes-html
  auto-report tests, html-format LOAD & REPORT test; existing pdf/docx-only
  scenarios now pin `HAS_HTML_REPORT=False` (their semantics are unchanged).

### Implementation
- `edf_bill_fetcher/io/cli.py` — `run_cli_html_report(args)` mirroring the
  PDF/DOCX handlers (records wrapper unwrap, config, restricted engine-state
  load, `generate_html_from_gui`), `main()` dispatch for `--html-report`,
  updated the no-tkinter hint.
- `edf_bill_fetcher/ui/app.py` — `HAS_HTML_REPORT` spec probe; "HTML Only"
  radio in `ReportOptionsDialog` Output Format frame; HTML branch in
  `_run_auto_report` and `load_spreadsheet_and_report` (both `fmt in
  ("html", "both")`, same as PDF/DOCX); Report Options button enablement
  now includes HTML availability.

## Verification

- `python -m pytest tests/test_dispatch_parity.py -q` → **25 passed** (new
  HTML parity + smoke + GUI tests green)
- Affected trio of test files → **139 passed**
- Full suite → **1535 passed, 9 skipped** (77s)
- `ruff check .` → All checks passed
- `ruff format --check .` → clean (ran `ruff format` on the 5 changed files)
- `mypy . --exclude scratch/` → Success (202 source files)
- Real CLI smoke: `python main.py --html-report -i fixture.json -o out.html`
  → exit 0, 16,845-byte HTML, 15 section headings
- GUI: real `ReportOptionsDialog` under the test display exposes the `html`
  format value

## Must-NOT compliance

- No registry entries removed; the registry stays complete. HTML placeholder
  sections keep their `section_builders` entries; parity asserts key-set
  equality + callable — not full coverage — exactly as the brief required.
- Not pushed.

## Notes

- The 4 registry-parity tests pass on the merged renderer immediately (the
  HTML `section_builders` literal was deliberately AST-compatible); they
  act as the regression gate for future registry edits. RED was genuinely
  observed on the CLI/GUI tests (ImportError / missing radio / missing
  `HAS_HTML_REPORT`).
- `"both"` now emits PDF + DOCX + HTML in the auto-report paths (mirrors the
  three-format wave); three existing tests pinning pdf/docx-only semantics
  were updated by pinning `HAS_HTML_REPORT=False`.
