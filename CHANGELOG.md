# Changelog

All notable changes to the EDF Energy Billing Evidence Collector
project. Dates are YYYY-MM-DD.

The format is loosely [Keep a Changelog](https://keepachangelog.com/),
semver-friendly.

## [Unreleased] — Audit pass (2026-06-19 → 2026-06-21)

A four-commit consultant-grade audit pass landed on `dev`. CI matrix
green across all 9 legs (Python 3.10/3.11/3.12 × linux/macos/windows).
Total test count: 224 passed (up from 162), 2 skipped.

Plus a fifth commit (this round) demotes the heavy pickers
(`libpff-python`, `statsmodels`) from optional extras to runtime
deps so paying clients get every documented feature with a single
`pip install -e .` — no tag-choosing required.

### Changed

- `pip install -e .` is now feature-complete for paying clients.
  `libpff-python` (PST/OST ingestion) and `statsmodels`
  (Holt-Winters forecasting) are pulled in by the default install
  instead of sitting behind `[pst]` and `[statsmodels]` extras.
  Optional extras are reduced to `[dev]` (test/lint/typecheck
  toolchain) and `[build]` (PyInstaller).
- `extract_new_invoice_fields` and `extract_new_credit_fields`
  account-number regex now match both EDF renderings: compact
  `A-NNNNNNNN` and grouped `NNN NNN NNN NNN`. Pre-fix, the spaced
  form was silently dropped and the `--acc-filter` could not match.
- `edf_report._get_package_version` reads the version from
  `pyproject.toml`. The cover page now displays the on-disk
  version rather than the stale `v0.1.0` literal that was hardcoded
  there.
- `edf_report.generate_ombudsman_pdf` no longer offers
  `appendix_filtered` as a candidate section that has no builder.
  Removed from `all_sections` to keep request → render honest.

### Fixed

- `READING_PATTERNS["Actual"]` was over-broad: it matched the bare
  word "actual" anywhere in a bill body, including ordinary prose
  like *"the actual amount you owe is £240"*. Fixed to require
  meter-reading language so the dispute-classification logic
  doesn't misroute records whose bill body has no reading data at
  all.
- `_data_quality_report`'s `ur_computable` had a dead code clause
  `and x != "N/A"` — unreachable after the prior `isinstance`
  guard. Dropped.
- `parse_htm_account_history` (the #15 HTM fix): three verb-aware
  regexes (charged / paid / reversed) now accept `in (?:debit|credit)?`
  so a credit-flagged statement is no longer silently dropped. Plus a
  fourth regex for the new standalone `Balance £X in credit`
  opening-line shape, with a covered-range guard so the verb-aware
  matches and the standalone match never double-count.
- `Extract_new_invoice_fields` has the same shape — the
  `Current balance` and `Total charges for this period` regexes now
  accept `credit?` and produce an `amount_side` field documenting
  which label was seen.

### Added

- `EvidenceEngine.process_pst_file(path)` and
  `EvidenceEngine.process_ost_file(path)` — the per-file wrappers
  that round out the public API. Pre-fix, only `process_pdf_file`
  and `process_htm_file` existed; PST/OST ingestion was reachable
  only through `crawl_pst(folder)` requiring a `pypff.file()`
  already opened. New wrappers open the file, drive the crawler,
  and close cleanly. `process_ost_file` is an alias since
  `libpff-python` accepts both formats.
- `tests/test_audit_pass_1.py` — 34 regression tests for the
  public-surface contracts: reading-pattern ordering,
  `detect_pdf_format`, `process_text` heuristic-fallback paths,
  `_detect_payment_patterns`, `_analyze_tariff_impact`,
  `_data_quality_report`, `process_pst/ost_file`, and
  `compute_dispute_flags` ordering invariants.
- `tests/test_report_version.py` — 3 tests pinning the package-
  version helper to (a) return a string, (b) match the `pyproject.toml`
  declared version, (c) fall back to `0.1.0` if `pyproject.toml` is
  unreadable.
- `tests/test_dispatch_parity.py` — 7 structural tests asserting
  the PDF and DOCX dispatchers expose the same key set as
  `REPORT_SECTIONS`, and that `RenderContext()` defaults to
  rendering every registry section.
- `tests/test_integration_pipeline.py` — 2 end-to-end tests that
  walk the bundled synthetic fixture PDF (`tests/fixtures/sample_bill.pdf`)
  through the full pipeline: extract → reportlab PDF + openpyxl
  XLSX. The fixture is auto-regenerated via `runpy` if missing.
- `tests/fixtures/generate_bill_fixture.py` — deterministic EDF
  KI-style bill PDF generator using `reportlab` (already a project
  runtime dep). The fixture renders only synthetic placeholder
  data (FAFA project policy): no real EDF account numbers, real
  addresses, or real amounts.
- `tests/fixtures/sample_bill.pdf` — committed; size ~3.6 KB;
  carrier exception `!tests/fixtures/*.pdf` in `.gitignore`.

### Test pyramid track during this audit

| Stage | Pass count | Comment |
| --- | --- | --- |
| Pre-audit baseline | 162 | CI green |
| Audit-pass-1 (READING fix, pypff wrappers, dead-code cleanup) | 162 → 214 | +52 tests |
| Audit-pass-2 (version literals, appendix_filtered removal) | 214 → 217 | +3 tests |
| Audit-pass-3 (dispatcher parity tests) | 217 → 224 | +7 tests |
| Windows runner CI fix | 224 | same count, cross-platform CI |

### CI matrix results (this audit)

- Test (Python 3.10, ubuntu-latest) — green
- Test (Python 3.11, ubuntu-latest) — green
- Test (Python 3.12, ubuntu-latest) — green
- Test (Python 3.10, windows-latest) — green
- Test (Python 3.11, windows-latest) — green
- Test (Python 3.12, windows-latest) — green
- Test (Python 3.10, macos-latest) — green
- Test (Python 3.11, macos-latest) — green
- Test (Python 3.12, macos-latest) — green

All passing — ruff check, ruff format --check, mypy, pytest.

## Project history

The project is in alpha (Development Status :: 3 - Alpha on PyPI
classifiers). The original authored iteration was for a personal
EDF billing dispute. Long-time self-hosted workflow with PST/OST
email archives, local PDF folders, and HTM account exports as
input sources.
