# Changelog

All notable changes to the EDF Energy Billing Evidence Collector
project. Dates are YYYY-MM-DD.

The format is loosely [Keep a Changelog](https://keepachangelog.com),
semver-friendly.

## [Unreleased]

## [0.5.0] - 2026-08-21

This release adds a Superseded Reconciliation view for the back-billing
pipeline and a round of evidence-integrity fixes: malformed SAP rows are
now surfaced instead of silently padded, the pickled engine no longer
drops its SAP/source-path state, and the workbook renders how each
unlawful charge was derived.

### Added

- Superseded Reconciliation worksheet recording each superseded
  back-billing invoice with its `KILLER` group header, jump links to the
  survivor's Back-billing Analysis row and the original Evidence Report
  row, and `file://` links to both saved PDFs.
- Back-billing Analysis now shows only live rows with a single union
  (no-double-count) unlawful-charge total; superseded invoices move to
  the reconciliation view.
- `Sub-Period Basis` provenance column on Back-billing Analysis, showing
  whether each unlawful charge came from per-sub-period rates or the
  day-ratio fallback.
- Evidence Report `Attachment Name` cells link to the saved PDFs; evidence
  files are named by invoice number (sanitised, deduped).
- Evidence-bundle `saved`/`missing`/`ambiguous` counts surfaced in the GUI
  summary when files are missing or an attachment is referenced by
  multiple invoices.

### Changed

- Extracted the SLC 7A 12-month limit into a named `backbilling_cutoff`
  helper (a deliberate fixed 365-day interval), with leap-day and
  month-end boundary tests.

### Fixed

- Malformed SAP CSV rows (fewer/more fields than the header) are counted
  and reported into the parse-error log instead of silently accepted.
- `EvidenceEngine` pickle round-trip now persists `sap_*_rows` and
  `source_paths`, with backwards-compatible defaults for older pickles.
- `View superseded` links point at the survivor's own `KILLER:` header row
  on the reconciliation sheet.

### Verification

- CI matrix green across Python 3.10, 3.11, and 3.12 on Linux, macOS, and
  Windows.
- 1,470 tests passed, 9 skipped, with 92.39% coverage at release
  preparation.

## [0.4.0] - 2026-08-13

This release makes the dispute-analysis pipeline substantially more reliable,
particularly for late billing, cancelled/reissued invoices, and SAP-ledger
cross-checking. It also improves report consistency, packaging, and CI
reliability.

### Added

- Legal back-billing analysis based on bill date versus consumption period,
  including unlawful-slice proration and a documented `Period Charge (£)` to
  `Amount (£)` fallback.
- Recursive rebilling supersession analysis with preserved audit rows,
  `Status`, `Superseded By`, and `Partial Overlap` fields.
- SAP back-billing matching using posting and clearing dates, period-charge
  amounts, and internal-mechanism cluster labels.
- Packaged OFGEM quarterly price-cap data loaded from JSON with validation.
- Typed `BillingRecord` and shared `PaymentAnalysis` models.
- Shared sheet-layout helpers for back-billing, evidence, payment, and SAP
  worksheets.

### Changed

- Shared parsing, matching, reconciliation, formatting, and anomaly helpers
  now have canonical implementations instead of duplicated copies.
- PyInstaller builds include packaged data files for one-file bundles.

### Fixed

- Corrected the back-billing eligibility gate to use bill date versus
  `Period From`, rather than period length or `Period To` alone.
- Corrected period-charge selection so running account balances are not used
  as the primary deduction amount.
- Corrected unlawful-charge proration, recursive supersession, and SAP
  amount/date matching edge cases.
- Fixed report payment-analysis divergence between Excel, PDF, and DOCX.
- Fixed evidence links, sheet ordering, report formatting, carry-forward,
  non-numeric amounts, date fallbacks, and dispute deltas.
- Added strict empty-extraction handling through the CLI `--strict` option.
- Hardened macOS application signing/DMG packaging and Windows Tcl/Tk CI
  reliability.

### Verification

- CI matrix green across Python 3.10, 3.11, and 3.12 on Linux, macOS, and
  Windows.
- 1,411 tests passed, 9 skipped, with 92.55% coverage at release preparation.

Adds four Excel analysis tabs (Back-billing Analysis, Rebilling &
Corrections, Meter Readings, Contract History) backed by four
pure-pandas detectors, plus multi-invoice PDF support and admit-
phrase extraction. No existing sheet writer is modified; the new
tabs append after the existing set. Test count rises from 395 to
498 (+103 new tests across 9 new test files). Gate stays green
across all 9 CI legs (Python 3.10/3.11/3.12 × linux/macos/windows).

### Added — features

- **Multi-invoice PDF slicing** (`slice_pdf_pages`): a PDF's per-page
  text is sliced at `Invoice number:` or `Page 1 of N` boundaries
  (variants `1 of 4`, `one of four`, `1/4`), so each invoice in a
  merged PDF becomes its own row. Single-invoice PDFs are unchanged
  (one slice = whole document). `process_pdf_file` dispatches each
  slice through the existing format-detect path with the `#i` slice
  index suffixed to `detail_label` and `attachment_name`. Per-slice
  try/except isolates a bad slice from the rest of the file.
- **Admit-phrase extraction** (`extract_admit_phrase`): recognises
  EDF's cover-page wording (the "we've recently cancelled some
  charges for you" family) under the new `_ADMIT_RE` regex. Returns
  the matched substring or `None`; rejects `cancel your direct
  debit`-style false positives that lack the operative charge-
  cancellation verb.
- **Back-billing detector** (`detect_back_billing`): surfaces any
  single invoice whose `Period From` → `Period To` window exceeds
  the SLC 7A 12-month limit (>365 days). `Reason Assessment` is a
  deterministic narrative (no LLM) calling out the excess days and
  whether the cover page admits a cancellation. `Cancel/Rebill
  Admitted` column comes from the admit-phrase extractor.
- **Rebilling detector** (`detect_rebilling`): pairs of invoices
  where a "killer" later-issued bill cancels and reposts a "killed"
  earlier bill. Fires on period overlap > 30 days OR jump-back > 30
  days OR long-period (≥60 days) killer whose Period From ≤ the
  killed's Period From. `Trigger Reason` lists every trigger that
  fired, joined by `; `.
- **Meter rollover detector** (`detect_meter_rollover`): walks
  Actual/Smart readings, computes delta on the `Units (kWh)` column,
  and emits a row when the delta is negative with `abs(delta)` above
  `rollover_threshold` (default 94,999 = 99,999 − 5,000). Caller-
  supplied threshold lets the algorithm tune for shorter or longer
  meter caps.
- **Contract inferrer** (`infer_contracts`): groups consecutive rows
  with the same `Tariff` value into one contract, merging adjacent
  same-tariff runs whose inter-group gap is < `merge_gap_days`
  (default 30). `N/A` tariffs are skipped.
- **Back-billing Analysis tab** (`write_back_billing_sheet`): title
  banner with SAP account, legal-context header citing Electricity
  Act 1989 s.84B and Ofgem's back-billing rule, 10-col table, total
  retrospective charges footer. `Cancel/Rebill Disclosed` carries
  one of `Admitted phrase`, `Period overlap`, `Admitted + overlap`,
  or blank, taking the row's admit-phrase flag plus an optional set
  of overlapping invoice numbers supplied by the rebilling
  detector.
- **Rebilling & Corrections tab** (`write_rebilling_sheet`): title
  banner, subheader paragraph, 7-col table.
- **Meter Readings tab** (`write_meter_readings_sheet`): A/E/M
  timeline (`A`=Actual, `E`=Estimated, `M`=Meter rollover
  candidate). Estimated Source mirrors the row's `Details` column
  for Estimated rows. Rollover rows flagged `M` show a Notes blurb
  pointing at the rollover table.
- **Contract History tab** (`write_contract_history_sheet`):
  title, 5-col table.
- **Orchestrator** (`run_analysers`): thin wrapper returning a dict
  of the four detector outputs so `export_to_excel` calls them with
  one line.
- **Wiring**: `export_to_excel` now calls `run_analysers(dfc)` after
  the existing sheet writers and before `wb.save`, appending the
  four new tabs. Account label pulled from `config['acc_num']`. No
  existing sheet writer touched.

## [Unreleased] — UI refresh (2026-07-11)

A focused QOL pass on the tkinter desktop GUI. No changes to the
extraction pipeline, the dedup walker, or the report renderers
beyond plumbing the already-existing `amalgamate_duplicates` config
key through to the UI. Test count rose from 347 to 395 (+48 new
tests, 8 new test files). Gate stays green across all 9 CI legs
(Python 3.10/3.11/3.12 × linux/macos/windows).

### Added — features

- **Output Folder picker** in Section 1. Empty value falls back to
  the first source-file's directory at run time, preserving the
  pre-refresh default behaviour.
- **Sequential non-overwriting file naming**: `<stem>_<YYYY-MM-DD>_<N>.xlsx`
  and `_Report.pdf` / `_Report.docx` variants. Counter is per-day
  per-folder and shared across all outputs in one EXTRACT batch.
- **Auto-generate report after extraction** checkbox in Section 2.
  When ON, EXTRACT also runs the configured PDF / DOCX report
  generators with the same batch counter as the xlsx. When OFF, only
  the xlsx is written (today's behaviour).
- **Report Options** button in the action bar (replaces the former
  EXPORT REPORT button). Opens the same `ReportOptionsDialog`; the
  selected format + sections are persisted to the config file on OK.
- **`~/.edf_collector/config.json`** persistence for GUI state and
  report options. Atomic write (temp + fsync + `os.replace`),
  `0o600` permissions, silent fallback to defaults when the file is
  missing or malformed. Deleting the file resets state cleanly.
- **Amalgamate toggle** now surfaces as a third nested checkbox in
  Section 3, enabled only when both *Drop duplicates found across
  sources* and *Record dropped duplicates on side sheet* are ON.
  Default OFF (matches the `export_to_excel` default).
- **Three-state EXTRACT button**: `EXTRACT TO EXCEL` → `Cancel`
  (navy) → `Cancelling...` (grey) → `EXTRACT TO EXCEL`. The separate
  Cancel button is gone; clicks on the running button call
  `_cancel`, which flips the button to `Cancelling...` and disables
  further clicks until `_finish` resets it to Idle.

### Changed — labels

- "Filter duplicate records (same date & amount)" → "Drop
  duplicates found across sources".
- "Save duplicates to separate worksheet" → "Record dropped
  duplicates on side sheet (Duplicate Entries)".
- "Save filtered-out records to worksheet" → "Keep filtered-out
  records on side sheet (Filtered (Below Min))" (relocated as an
  indented child of the filter-below row).
- "Output filename:" row moved from Section 2 to Section 1, beside
  the new Output Folder picker.

### Removed — dead code

- `App.export_report` and `App._export_legacy` deleted
  (~270 LOC). The three-state button + auto-generate flow removes
  any call site that previously reached them.

### Tests

- `test_config_persistence.py` (7) — round-trip, missing/malformed
  file, atomic write, permissions, report_options persistence.
- `test_output_folder_var.py` (4) — App declares new tk vars.
- `test_sequential_naming.py` (7) — empty folder, increment, shared
  batch counter, per-day reset, non-numeric suffix ignored.
- `test_output_folder_picker.py` (4) — label exists, var set/get,
  empty default.
- `test_ui_section2_section3.py` (14) — relocated save_filtered,
  enable/disable chain, auto-gen checkbox, relabelled dedup labels,
  amalgamate toggle state wiring.
- `test_extract_button_state.py` (8) — three-state flip + buttons.
- `test_report_dialog_persist.py` (2) — OK persists, cancel no-op.
- `test_dead_code_removed.py` (2) — export_report/_export_legacy
  are gone.

### CI

- `tests/conftest.py` probes for a usable Tk display and skips the
  GUI test files via `collect_ignore` when none is found
  (headless Linux / macOS runners). The Windows-only failure on
  `test_save_config_file_permissions_0600` is skipped via
  `@pytest.mark.skipif(os.name == "nt", ...)` because `os.chmod`
  doesn't enforce Unix mode bits there.

---

## [Dev-branch review pass] (2026-06 → 2026-07)

Two audit passes landed on the `dev` branch. The June pass hardened
the public API surface and the test/gate contract; the July pass
resolved a cluster of correctness and contract regressions in the
report renderers and the dedup walker. The combined work lifted the
test count from 162 → 347 with no behavioural regressions to the
existing suite contract.

CI matrix green across all 9 legs (Python 3.10/3.11/3.12 ×
linux/macos/windows). Same four-step gate as the June round
(`ruff check`, `ruff format --check`, `mypy`, `pytest`) enforced
locally before every push.

### Added — features

- **`EvidenceEngine.process_pst_file(path)` and
  `process_ost_file(path)`** per-file wrappers round out the public
  source-processing API. Pre-fix, only `process_pdf_file` and
  `process_htm_file` existed; PST/OST ingestion was reachable only
  via `crawl_pst(folder)` which required a pre-opened `pypff.file()`.
  `process_ost_file` is an alias since `libpff-python` accepts both
  formats.
- **`amalgamate_duplicates` config toggle** (default `False`). When
  `True`, the dedup walker produces a single *hybrid* row per
  duplicate cluster instead of the most-complete row. Each
  per-column value on the hybrid row is the first populated value
  pulled from any sibling in completeness-descending order; `Source`
  is pinned to the completeness-winner's identity. Every
  non-surviving sibling still surfaces on the Duplicate Entries
  sheet ("never drop without recording" contract). Toggle `OFF`
  preserves the most-complete-row behaviour.
- **Dedup walker kept-set selector now ranks by completeness.**
  Pre-fix the dedup sort was `["_src_pri", "_sort"]` (`keep="first"`)
  so a sparser row from a higher-precedence source would beat a
  richer row from a lower-precedence source. Post-fix the sort is
  `["_completeness" (desc), "_src_pri" (asc), "_sort" (asc)]`.
  Helper columns are stripped before the writer so saved workbook
  geometry is unchanged.
- **`report_sections` config key** lets the GUI / programmatic
  caller select a subset of sections to render. Section titles and
  numbering are derived from `REPORT_SECTIONS` so the TOC and body
  always agree regardless of which sections a user selects.

### Added — tests

- `tests/test_audit_pass_1.py` — 34 regression tests pinning
  reading-pattern ordering, `detect_pdf_format`, `process_text`
  heuristic-fallback paths, `_detect_payment_patterns`,
  `_analyze_tariff_impact`, `_data_quality_report`,
  `process_pst/ost_file`, and `compute_dispute_flags` ordering
  invariants.
- `tests/test_report_version.py` — pins the package-version helper
  (`_get_package_version`): string return, `pyproject.toml` value,
  `0.1.0` fallback on read failure.
- `tests/test_dispatch_parity.py` — 7 structural tests asserting the
  PDF and DOCX dispatchers expose the same key set as
  `REPORT_SECTIONS`, and `RenderContext()` defaults to rendering
  every registry section.
- `tests/test_integration_pipeline.py` — 2 end-to-end tests that
  drive the bundled synthetic fixture PDF
  (`tests/fixtures/sample_bill.pdf`) through the full pipeline:
  extract → reportlab PDF + openpyxl XLSX. Fixture auto-regenerated
  when missing.
- `tests/fixtures/generate_bill_fixture.py` — deterministic KI-style
  EDF bill PDF generator using `reportlab`. Renders purely
  synthetic placeholder data (no real EDF account numbers, addresses,
  or amounts).
- `tests/fixtures/sample_bill.pdf` — committed; ~3.6 KB;
  `.gitignore` carrier exception `!tests/fixtures/*.pdf` keeps it.
- `tests/test_docx_critical_fixes.py` — 10 tests pinning the DOCX
  cover-page / glossary / OFGEM-carry-forward triple-bug fix.
- `tests/test_pdf_xml_injection.py` — 9 tests pinning the reportlab
  `Paragraph(...)` XML-injection escape at every user-data callsite.
- `tests/test_pdf_tablecell_xml.py` — 3 tests pinning the reportlab
  `Table(...)` cell escape against the evidence-index and appendix
  builders, using `pdfplumber.extract_tables()` for cell-boundary-
  preserving assertions.
- `tests/test_dedup_most_complete.py` — pinned the Spec 2 "most
  complete version of the information presented" sort contract.
- `tests/test_amalgamate_duplicates.py` — 3 tests pinning the
  hybrid-row contract (column-wise merge, dup-sheet completeness,
  toggle-OFF preserves the most-complete row).
- `tests/test_amalgamate_pass2.py` — 3 tests pinning the Pass-2
  anchor unification; the dedup walker's second pass now
  hybridizes via the same `anchor_to_dup_indices` map used by
  Pass-1 instead of dropping Pass-2 duplicates silently.
- `tests/test_nat_cluster_split.py` — 2 tests pinning the NaT-
  cluster-split fix; rows with an unparseable `Period To` no longer
  fall back to the parsed `Date` (which caused unrelated same-Amount
  events to collapse via NaT-as-equal in `duplicated`).
- `tests/test_dedup_helpers.py` — 17 unit tests on `_is_populated`,
  `_completeness_score`, `_amalgamate_cluster` (parametrised over
  truthy/falsy values, N/A markers, NaN-aware scoring).
- `tests/test_save_dups_toggle.py` — 3 tests pinning the
  `save_dups=True` (drop siblings into the Duplicate Entries sheet)
  vs `save_dups=False` (skip dedup entirely; user manually prunes
  in Excel with analysis sheets recomputing via live formulas).
- `tests/test_evidence_dup_marker.py` — 2 tests pinning the
  Duplicate Entries sheet's presence/absence per the `save_dups`
  toggle and the "Duplicate Entries" sheet name.
- `tests/test_dispute_flag_warnings.py` — 2 tests pinning the
  `compute_dispute_flags` silent-`pass` → `warnings.warn(...)`
  rewrite.

### Changed

- **`pip install -e .` is now feature-complete.**
  `libpff-python` (PST/OST ingestion) and `statsmodels`
  (Holt-Winters forecasting) are pulled in by the default install
  instead of sitting behind `[pst]` and `[statsmodels]` extras.
  Optional extras reduced to `[dev]` (test/lint/typecheck
  toolchain) and `[build]` (PyInstaller).
- **`extract_new_invoice_fields` and `extract_new_credit_fields`**
  account-number regex now match both EDF renderings: compact
  `A-NNNNNNNN` and grouped `NNN NNN NNN NNN`. Pre-fix, the spaced
  form was silently dropped and the `--acc-filter` could not match.
- **`edf_report._get_package_version`** reads the version from
  `pyproject.toml`. The cover page now displays the on-disk version
  rather than the stale `v0.1.0` literal that was hardcoded there.
- **`edf_report.generate_ombudsman_pdf`** no longer offers
  `appendix_filtered` as a candidate section that has no builder.
  Removed from `all_sections` to keep request → render honest.
- **`save_dups=False` now skips dedup entirely.** Pre-fix both arms
  of the `if config.get("save_dups", True)` block set
  `dup_df = df[is_dup].copy()` verbatim — the toggle was dead.
  Per the product spec ("never drop without being recorded"), the
  contract is: `save_dups=True` runs dedup, drops duplicates into
  the Duplicate Entries sheet, and the main report sheet carries
  only the kept row; `save_dups=False` skips dedup so every row
  stays in `df`, the analysis sheets recompute against the full
  set via their IFERROR/SUM/AVERAGEIFS formulas, and the user
  manually prunes in Excel.

### Fixed

- **DOCX cover-page labels were invisible.** Pre-fix
  `python-docx`'s `Cell.__eq__` does not compare across freshly-
  fetched proxies, so the existing `cell == row.cells[0]` check
  silently evaluated `False` for every column. Column-zero labels
  were never styled NAVY/bold — they were styled DARK_GREY regular.
  Fixed by switching to `enumerate(row.cells)` + `col_idx == 0`.
- **DOCX glossary header row was blank.** Pre-fix the glossary
  builder allocated rows for header + terms but never wrote the
  header cells (it started the term loop at index 1, leaving row 0
  empty). Fixed by writing `["Term", "Definition"]` into row 0
  before the iteration.
- **DOCX OFGEM section silently marked out-of-window quarters as
  `CAP DATA UNAVAILABLE`.** Pre-fix the DOCX runner only inspected
  the hard-coded `_load_ofgem_caps()` table; quarters beyond it
  were marked unavailable even though the PDF surfaces the
  `_LATEST_KNOWN` carry-forward sentinel. Fixed by porting the
  PDF carry-forward branch verbatim: `config` parameter,
  `_load_ofgem_caps(auto_carry=False)`, sentinel consumption, and
  the `COMPLIANT (CARRIED)` summary verdict.
- **CRITICAL — `edf_report.py` Paragraph XML-injection.** Nine
  `Paragraph(...)` callsites interpolated user-derived strings (PDF /
  PST / HTM source data, exec-summary period dates, evidence-index
  source labels, methodology config bullets, dispatcher exception
  text) into the f-string passed to `Paragraph(...)` *without* XML
  escaping. Reportlab's `Paragraph` interprets inline markup
  (`<b>`, `<i>`, `<font>`, `<br/>`) plus `&`/`<`/`>` as XML, so a
  malicious payload injects new tags or parse-fails the document.
  Pre-fix behaviour confirmed by
  `ValueError: Parse error: saw </b> instead of expected </para>`.
  Fixed by `xml_escape(...)` wrapping each untrusted segment.
  Defense-in-depth: `report_date` (a trusted `datetime.now()` value)
  is now also escape-wrapped so a future producer change cannot
  regress silently.
- **PDF `Table(...)` cells with `&` / `<` / `>` characters** now
  route through the same `xml_escape` helper at six user-data
  callsites (`Source`, `Invoice #`, `Entry Type`, `Reading`,
  `Period From`/`To`, `Units`, etc.) inside `create_appendix_full_
  evidence` and `create_evidence_index`. Reportlab Tables do not
  pass cells through the miniHTML parser, but the escaped entity
  form (`<bad>`) is cleaner to read in an audit-grade report
  than a literal `<bad>` token that a reviewer might mistake for
  markup.
- **`compute_dispute_flags` silently swallowed parse errors.** Five
  `except (ValueError, TypeError, KeyError): pass` clauses dropped
  the row whenever a missing key or unparseable string bit the
  heuristic. Fixed by routing each clause through a `warnings.warn`
  helper (`_flag_or_warn`) with the row index, flag name, and
  exception. The function still completes the run when a row
  unparses so the heuristic doesn't crash; a row that fails to
  evaluate now logs a `UserWarning` instead of vanishing.
- **`process_pdf_file` page-level exception was over-broad.**
  Pre-fix `except Exception as page_err` swallowed every per-page
  pdfplumber failure including unexpected runtime errors.
  Narrowed to `(pdfplumber.utils.exceptions.PdfminerException,
  ValueError, TypeError)` so PDF-syntax and text-coercion errors
  log + skip the page but unexpected errors propagate.
- **Dedup walker NaT-cluster merging.** Pre-fix `_dedup_date` fell
  back to the parsed `Date` when `Period To` was unparseable. For
  a no-period PDF row whose `Date` was also unparseable, both
  cluster keys collapsed to NaT and `duplicated(keep="first")`
  merged unrelated same-Amount rows via NaT-as-equal. Post-fix
  `_dedup_date` stays NaT for unparseable rows; Pass-1 routes NaT
  dedup dates through the no-period bucket, which views them as
  distinct unless their own (date, amount) overlaps a 60-day
  window.

### Tightened

- Pickle surface restricted via `_RestrictedUnpickler.find_class`
  whitelist. Pickle is used for cached engine state; the restricted
  unpicker blocks arbitrary-class instantiation so a hostile
  `--engine` pickle file cannot escalate.

### Test pyramid track

| Stage | Pass count | Δ |
| --- | --- | --- |
| Pre-audit baseline (2026-06) | 162 | — |
| Audit pass A (READING fix, pypff wrappers, dead-code cleanup) | 162 → 214 | +52 |
| Audit pass B (version literals, appendix_filtered removal) | 214 → 217 | +3 |
| Audit pass C (dispatcher parity tests) | 217 → 224 | +7 |
| Windows runner CI fix (cross-platform CI) | 224 | same count |
| Dev-branch review Round 1 (DOCX triple-bug) | 224 → 301 | +77 |
| Dev-branch review Round 2 (PDF XML-injection) | 301 → 310 | +9 |
| Dev-branch review Round 3 (reconciliation warnings) | 310 → 312 | +2 |
| Dev-branch review Round 4 (save_dups toggle + dup marker) | 312 → 317 | +5 |
| Dev-branch review Round 5 (most-complete selection) | 317 → 319 | +2 |
| Dev-branch review Round 6 (amalgamate hybrid) | 319 → 322 | +3 |
| Dev-branch review Round 7 (Pass-2 anchor unification, NaT split, helper unit tests) | 322 → 347 | +25 |

Net: **+185 tests across the audit work, 0 behavioural
regressions in the previous-suite contract.** Final gate is green
at 347 passed, 2 skipped.

### CI matrix results

- Test (Python 3.10, ubuntu-latest) — green
- Test (Python 3.11, ubuntu-latest) — green
- Test (Python 3.12, ubuntu-latest) — green
- Test (Python 3.10, windows-latest) — green
- Test (Python 3.11, windows-latest) — green
- Test (Python 3.12, windows-latest) — green
- Test (Python 3.10, macos-latest) — green
- Test (Python 3.11, macos-latest) — green
- Test (Python 3.12, macos-latest) — green

All passing — `ruff check`, `ruff format --check`, `mypy`, `pytest`.

### Verification gate (latest session)

- `ruff check .` ✓ (no findings)
- `ruff format --check .` ✓ (41 files, all formatted)
- `mypy .` ✓ (47 source files, no issues)
- `pytest` ✓ (347 passed, 2 skipped, 10 warnings, 5.40s)

## Project history

The project is in alpha (Development Status :: 3 - Alpha on PyPI
classifiers). The original authored iteration was for a personal
EDF billing dispute. Long-time self-hosted workflow with PST/OST
email archives, local PDF folders, and HTM account exports as
input sources.
