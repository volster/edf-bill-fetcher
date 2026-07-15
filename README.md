# EDF Energy Billing Evidence Collector

A personal desktop application that collects and analyses EDF Energy billing data from PST/OST email archives, local PDF bills, and HTM account history exports. Produces a multi-sheet Excel evidence workbook and a PDF **or** DOCX report suitable for Energy Ombudsman submission.

## Features

- **Multi-source extraction**: PST/OST email files, local PDF folders, HTM account exports.
- **Dual format support**: Parses both old-style and new-style (KI/KCR) EDF invoice formats.
- **Smart amount detection**: Prioritized regex patterns with configurable fallback.
- **Cross-source deduplication**: Two-pass dedup — Period To + Amount (primary), Amount within a 60-day window (secondary). Within a duplicate cluster the *most complete* row wins (substantive field fill-rate), with source precedence as the tie-breaker. Set `amalgamate_duplicates=True` to merge columns across all duplicate siblings into a single hybrid kept row (each sibling still surfaces on the Duplicate Entries sheet for audit).
- **Comprehensive Excel output**: Multi-sheet evidence workbook with annual summary, dispute flags, statistical analysis, payment analysis, and forecast sheets.
- **Professional PDF + DOCX output**: 14 dynamically-numbered sections. Numbering is **derived from `REPORT_SECTIONS` so the Table of Contents and body always agree**, regardless of which sections a user selects in the report options dialog.
- **GUI interface**: tkinter-based desktop application with progress tracking.
  - **Output Folder picker** (Section 1): choose where xlsx + report outputs land; empty falls back to the source-file directory.
  - **Sequential file naming**: `<stem>_<YYYY-MM-DD>_<N>.xlsx` (and `_Report.pdf` / `_Report.docx` variants). Counter is per-day per-folder and shared across all outputs in one batch.
  - **Auto-generate report**: a checkbox in Section 2 to produce xlsx + PDF + DOCX from a single EXTRACT run, sharing the same batch counter.
  - **Three-state EXTRACT button**: `EXTRACT TO EXCEL → Cancel → Cancelling... → EXTRACT TO EXCEL` — theEXTRACT and Cancel buttons are now collapsed into one.
  - **Config persistence**: GUI state and report options stored at `~/.edf_collector/config.json` (atomic write, `0o600`); deleting the file resets state cleanly.
  - **Amalgamate toggle**: the `amalgamate_duplicates` checkbox surfaces as a nested child under Deduplication (enabled only when both *Drop duplicates* and *Record dropped duplicates* are on; default OFF).
- **CLI mode**: Headless batch/report generation for automation.

## Installation

```bash
git clone https://github.com/volster/edf-bill-fetcher.git
cd edf-bill-fetcher

# Single install covers every documented feature:
# GUI, CLI, PST/OST archive parsing, Holt-Winters forecasting,
# Excel + PDF + DOCX report generation.
pip install -e .

# Optional toolchains (only needed if you are contributing or
# packaging binaries — the default extras-as-runtime policy above
# typical installs at one command):
#
#   [dev]    test + lint + typecheck toolchain (pytest, ruff, mypy)
#   [build]  PyInstaller for one-file Windows / macOS / Linux executables
pip install -e ".[dev,build]"   # recommended for contributors
```

What `pip install -e .` actually pulls in (current version):

| Library            | Used for                                                                      |
| ------------------ | ----------------------------------------------------------------------------- |
| pandas / numpy     | DataFrame plumbing across extraction, dedup, statistical analysis, export.    |
| pdfplumber         | Reading text + tables out of EDF bill PDFs.                                     |
| beautifulsoup4     | Stripping HTML into text body for PST / HTM ingestion.                         |
| openpyxl           | Writing the multi-sheet evidence workbook.                                      |
| reportlab          | PDF report rendering (cover, TOC, sections, summary tables, appendix).         |
| python-docx        | DOCX report rendering (sister surface to PDF).                                  |
| scipy              | Rolling stats, Shapiro-Wilk normality, linregress fallback forecasting.          |
| **`libpff-python`** | Outlook PST / OST archive ingestion — used by `EvidenceEngine.process_pst_file`. |
| **`statsmodels`**   | Holt-Winters forecasting in the Forecast section — falls back to linear+EMA     |
|                    | projection if missing (with a warning in the report).                          |

Adding optional toolchains later does not require re-installing everything:

```bash
pip install -e ".[dev]"     # bring in pytest / ruff / mypy only
pip install -e ".[build]"   # add PyInstaller
```

## Usage

### GUI Mode

```bash
python edf_collector.py
```

Or run the built executable `EDF_Evidence_Collector.exe`.

1. **Select Sources** (at least one required):
   - **PST/OST File**: Outlook email archive containing EDF emails (`.pst` / `.ost` are read via `libpff-python`; the `process_pst_file` wrapper auto-logs an error if the lib is missing, so the rest of the pipeline still runs)
   - **PDF Folder**: Directory containing EDF bill PDFs
   - **HTM Export**: EDF MyAccount "Payments and Invoices" HTM export
2. **Configure Options**:
   - **Account Filter**: Filter by EDF account number (e.g., `A-12345678` or `123 456 789 012`). Both compact and grouped-digit renderings are matched against the bill — `extract_new_invoice_fields` accepts the spaceless `A-NNNNNNNN` shape and the bank-row-style `NNN NNN NNN NNN` shape.
   - **Domain Filter**: Filter PST emails by sender domain (default: edfenergy.com)
   - **Minimum Amount**: Filter out records below this threshold (default: £500)
   - **Analysis Threshold**: Minimum bill amount for analysis tabs (default: £500)
   - **Report Account Ref**: Override account reference in report header
3. **Click "EXTRACT TO EXCEL"** — produces the evidence workbook. The same button toggles to `Cancel` while running and `Cancelling...` after you click Cancel; it returns to `EXTRACT TO EXCEL` when the worker exits.
4. **Optional: pick Output Folder** (Section 1) — defaults to the source-file directory; override to send xlsx + reports elsewhere.
5. **Optional: tick "Auto-generate report after extraction"** (Section 2) — produces a PDF + DOCX report (per the saved Report Options) in the same batch as the xlsx, sharing the same sequential counter.
6. **Report Options**: the navy `Report Options` button opens the format + section-picker dialog; selections persist across sessions.
7. **Click "LOAD & REPORT"** — load an existing xlsx and regenerate the report from it (uses sequential naming into the output folder, no save-as dialog).

### CLI Mode — headless report generation

```bash
# Generate a PDF report from already-extracted records
python edf_collector.py --pdf-report -i records.json -o report.pdf

# DOCX variant
python edf_collector.py --docx-report -i records.json -o report.docx
```

Pass `-c config.json` and `-e engine.pkl` to forward config + filtered-records state.

### Programmatic Usage

> Per-source API is symmetric — all three source types expose
> `process_<source>_file(path, source_label, detail_label, fallback_date)`
> so you can plug in any combination via the same
> call signature. PST requires `libpff-python` (a runtime dep
> since `0.1.0+`); if it's missing in some hand-built
> environment, the wrapper logs the error and the rest of the
> pipeline still runs.

```python
from edf_collector import EvidenceEngine, export_to_excel
from edf_report import generate_pdf_from_gui
from edf_report_docx import generate_docx_from_gui

config = {
    "use_anchors": True,
    "use_large": True,
    "use_reading_classification": True,
    "use_pdf_fields": True,
    "use_acc_filter": False,
    "acc_num": "",
    "min_amount": 500.0,
    "analysis_min": 500.0,
    "filter_below": True,
    "save_filtered": True,
    "use_dedup": True,
    "save_dups": True,
    "use_domain_filter": True,
    "domain_filter": "edfenergy.com",
    # report_sections tells the report generator which sections to include.
    # If absent, every section in `edf_report.REPORT_SECTIONS` is selected.
    "report_sections": [
        "exec_summary", "key_findings", "evidence_index", "detailed_findings",
        "timeline", "ofgem", "statistical", "payment", "forecast",
        "data_quality", "tariff",
        "appendix_methodology", "appendix_glossary", "appendix_full_evidence",
    ],
}

engine = EvidenceEngine(config, print)
engine.process_pdf_file("path/to/a.pdf",
                        source_label="Local PDF",
                        detail_label="bill.pdf",
                        fallback_date="2026-03-01")
engine.process_htm_file("path/to/export.htm",
                        fallback_date="2026-03-01")
engine.process_pst_file("path/to/archive.pst")
# …each call extracts whatever it can and appends to engine.records.

# Excel export
export_to_excel(engine.records, "output.xlsx", engine.error_log, config, engine.filtered_records)

# PDF export
generate_pdf_from_gui(
    records=engine.records,
    output_path="report.pdf",
    config=config,
    engine=engine,
    filtered=engine.filtered_records,
)

# DOCX export — same arguments, sees the same config / same registry
generate_docx_from_gui(
    records=engine.records,
    output_path="report.docx",
    config=config,
    engine=engine,
    filtered=engine.filtered_records,
)
```

## Report Section Layout

The PDF and DOCX reports are both built from a single section-registry so the titles and numbering always line up. The registry lives in `edf_report.REPORT_SECTIONS`:

| Class          | Sections                                                                                                                                                                                                              |
| -------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Main (numeric) | Executive Summary · Key Findings Summary · Evidence Index & Source Cross-Reference · Detailed Findings · Timeline of Events · OFGEM Price Cap Comparison · Statistical Analysis · Payment & Credit Analysis · Forecast & Projection · Data Quality Assessment · Tariff Impact Analysis |
| Appendix       | Methodology & Data Sources · Glossary · Full Evidence Table                                                                                                                                                                                                              |

Main sections are numbered **1, 2, 3, …** and appendices are lettered **A, B, C, …**, computed at render time based on the user's `report_sections` selection.

A `tests/test_dispatch_parity.py` structural test pins the invariant that the PDF dispatcher's `section_builders`, the DOCX dispatcher's `section_builders`, and `REPORT_SECTIONS` all expose exactly the same set of keys — so a future contributor who adds a section to the registry without wiring both dispatchers breaks CI, not the rendered report.

### Adding a new section

1. Add an entry to `REPORT_SECTIONS` in `edf_report.py` with `key`, `title`, and optionally `is_appendix`.
2. Add the matching key to `ReportOptionsDialog.SECTIONS` in `edf_collector.py` so it shows up in the GUI options dialog.
3. Add a `def create_<name>(...)` builder function in `edf_report.py`.
4. Wire the builder into the `section_builders` dispatch dict in **both** `generate_ombudsman_pdf` and `generate_ombudsman_docx`. Forgetting this raises a clear `RuntimeError` at report-render time — that's the loud-failure mode that keeps the dispatch in lockstep with the registry.

Removing a section: same steps in reverse.

## Output Sheets (Excel)

| Sheet              | Description |
| ------------------ | ----------- |
| **Annual Summary** | Yearly balance range, average, peak, low |
| **EDF Evidence Report** | All extracted records with live formulas |
| **Duplicate Entries** | Deduplicated records (if enabled) |
| **Filtered (Below Min)** | Records below the minimum-threshold |
| **Parse Errors** | Any extraction errors encountered |
| **Key Statistics** | Account overview, balance figures, periodic charges, reading quality, unit rates |
| **Balance Trend** | Time-series chart with rolling average and linear trend |
| **Year-on-Year** | Yearly comparison with YoY changes |
| **Period Charges** | Per-period charges with daily rates and dispute flags |
| **Dispute Flags** | Automated detection of anomalies (large jumps, billing gaps, estimated runs, reconciliation mismatches) |
| **Dispute Timeline** | Chronological event timeline for the dispute narrative |
| **Statistical Analysis** | Descriptive stats, rolling 6-period stats, EMA, momentum, volatility, z-score/IQR anomalies, Shapiro-Wilk normality tests |
| **Payment Analysis** | Payment/credit patterns, intervals, amounts, chronological detail with chart |
| **Forecast & Projection** | Linear regression, Holt-Winters exponential smoothing, EMA projection, confidence intervals, accuracy metrics |
| **Data Quality Report** | Completeness and read-quality metrics |
| **Tariff Analysis** | Tariff impact analysis |
| **Back-billing Analysis** | Invoices whose `Period From` → `Period To` window exceeds the SLC 7A 12-month limit (365 days), with the cancel/rebill ad- mission disclosed as `Admitted phrase`, `Period overlap`, `Admitted + overlap`, or blank. Legal context block cites Electricity Act 1989 s.84B and Ofgem's back-billing rule. |
| **Rebilling & Corrections** | Killer/killed invoice pairs identified by period overlap > 30 days OR jump-back > 30 days OR long-period killer ≥ 60 days reaching back into a prior invoice's window. Trigger Reason lists every matching heuristic. |
| **Meter Readings** | Actual vs Estimated timeline with meter-rollover candidates flagged `M`. Estimated Source mirrors the row's `Details` column (e.g. `Automatic estimate`). |
| **Contract History** | Contract periods inferred from tariff transitions, with ≤ 30-day gaps merging across same-tariff runs. |

## Back-billing & rebilling detection

The workbook carries four analysis tabs after the existing sheet set:

- **Back-billing Analysis**: surfaces any single invoice whose `Period From` → `Period To` window exceeds the 12-month SLC 7A back-billing limit (>365 days). Each row shows the excess days and a deterministic Reason Assessment narrative. The `Cancel/Rebill Disclosed` column flags invoices whose cover page contains an admit phrase (e.g. *"we've recently cancelled some charges for you"*) — matched case-insensitively against EDF's wording — or that overlap a prior invoice's period (`Period overlap`), or both.
- **Rebilling & Corrections**: pairs of invoices where a "killer" later-issued bill effectively cancels and reposts a "killed" earlier bill. Heuristics — period overlap > 30 days, jump-back > 30 days, or long-period (≥60 days) killer reaching back into the killed's start. Trigger Reason lists every trigger that fired.
- **Meter Readings**: A/E/M timeline (`A`=Actual, `E`=Estimated, `M`=Meter rollover candidate). Actual/Smart readings only count toward rollover detection; a negative kWh delta within 94,999 of the 5-digit cap flags the row.
- **Contract History**: one row per inferred tariff run. Adjacent same-tariff runs merge when their gap is ≤30 days (default).

The four detectors are pure-pandas functions keyed off the deduplicated evidence DataFrame — no LLM, no external service. Detectors:

| Detector | Output columns |
| --- | --- |
| `detect_back_billing(df)` | Invoice #, Bill Date, Period From, Period To, Days Billed, Net Charge (£), 12-Month Limit (days), Excess Days, Cancel/Rebill Admitted, Reason Assessment |
| `detect_rebilling(df)` | Killer Invoice, Killed Invoice, Killer Date, Killed Date, Period Overlap (days), Jump-back (days), Trigger Reason, Cancel/Rebill Admitted (Killer) |
| `detect_meter_rollover(df, rollover_threshold=94999)` | Date, Invoice #, Prev Units (kWh), Curr Units (kWh), Delta, Reading Type, Notes |
| `infer_contracts(df, merge_gap_days=30)` | Contract From, Contract To, Tariff, Days, # Invoices |

The detectors can be called directly from Python:

```python
from edf_collector import (
    detect_back_billing,
    detect_rebilling,
    detect_meter_rollover,
    infer_contracts,
    run_analysers,
)

# One-shot — returns a dict with keys back_billing, rebilling,
# meter_rollover, contracts.
analyses = run_analysers(deduped_df)
```

Multi-invoice (merged) PDF inputs — e.g. a bundle of 30+ pages containing 8 T-series invoices or 40 pages containing 10 KI-series invoices — are sliced at invoice boundaries before parsing, so each invoice gets its own row instead of being lost in the whole-file concat. The slicer looks for either an `Invoice number:` line or a `Page 1 of N` boundary marker (variants `1 of 4`, `one of four`, `1/4`).

## Supported EDF Formats

### New-Style Invoices (KI-XXXXXXXX)
- "Current balance £X in debit" or "Current balance £X in credit"
- "Total charges for this period £X in debit" or "… in credit"
- "Your charges: DD Mon YYYY - DD Mon YYYY"
- kWh usage, standing charge, tariff name
- Account number rendered as `A-NNNNNNNN` (compact) **or** `NNN NNN NNN NNN` (spaced) — both parsed since audit-pass-1

### New-Style Credit Notes (KCR-XXXXXXXX)
- "Total credits for this bill £X"

### Old-Style Bills
- "Your new account balance £X"
- Generic amount patterns with "balance", "total charges", "amount to pay"

### HTM Account History (since the #15 fix)
- "DD Mon YYYY We charged your account £X For Y kWh ... Balance £Z in debit"
- "DD Mon YYYY We charged your account £X For Y kWh ... Balance £Z in credit"
- "DD Mon YYYY You paid us £X ... Balance £Z in debit"
- "DD Mon YYYY You paid us £X ... Balance £Z in credit"
- "DD Mon YYYY Reversed account charge £X ... Balance £Z in debit | credit"
- Plus standalone opening-balance lines:
  "DD Mon YYYY Balance £X in credit" (only the credit side — debit-only opening balances get summarised by the next transaction)

## Requirements

- Python 3.10+
- tkinter (bundled with most Python installers)
- Runtime (all installed by `pip install -e .`):
  pandas, numpy, pdfplumber, beautifulsoup4, openpyxl, reportlab,
  python-docx, scipy, **libpff-python**, **statsmodels**.
- Toolchain (only for contributors / packagers, install via
  `pip install -e ".[dev,build]"`):
  pytest, ruff, mypy, pyinstaller.

## Development

```bash
pip install -e ".[dev]"
pytest -v
```

Tests are organised into three lake levels:

- **Unit tests** (most files in `tests/test_*.py`) pin the
  behaviour of public functions and structural invariants of the
  registry/dispatcher.
- **Audit regression tests** (`tests/test_audit_pass_1.py`,
  `tests/test_report_version.py`, `tests/test_dispatch_parity.py`)
  pin the contracts the report depends on. These exist
  *because* real-data review exposed one or more real defects — see
  `CHANGELOG.md` for the audit trail. Do not edit these without
  re-running audit-pass analysis:
  1. `tests/test_audit_pass_1.py` — reading-pattern ordering,
     `detect_pdf_format`, `process_text` heuristic-fallback,
     `_detect_payment_patterns`, `_analyze_tariff_impact`,
     `_data_quality_report`, `process_pst_file` / `process_ost_file`,
     `compute_dispute_flags`.
  2. `tests/test_report_version.py` — cover page reflects the
     `pyproject.toml` version; falls back to a stable default when
     `pyproject.toml` is unreadable.
  3. `tests/test_dispatch_parity.py` — REGISTRY ↔ PDF dispatcher ↔
     DOCX dispatcher key-set parity is locked in.
- **Integration smoke** (`tests/test_integration_pipeline.py`) drives
  the bundled synthetic bill PDF (`tests/fixtures/sample_bill.pdf`)
  through the full PDF → engine → reportlab PDF + openpyxl XLSX
  pipeline and asserts the extracted fields. The fixture is
  regenerated via `tests/fixtures/generate_bill_fixture.py` if
  missing — a fully-synthetic, deterministic dataset (FAFA policy,
  no real EDF data).

### Linting / formatting / type-checking

```bash
ruff check .
ruff format .
mypy .
```

All three are enforced in CI on Python 3.10 / 3.11 / 3.12 ×
ubuntu / windows / macos — see `.github/workflows/ci.yml`. A
release is one CI green away from shippable.

## Configuration

All options are available in the GUI. For programmatic use, see the `config` dict in the usage example above.

Key options:
- `use_anchors`: enable smart-context amount patterns
- `use_large`: enable large-amount fallback
- `use_reading_classification`: classify Estimated/Actual/Smart readings
- `use_pdf_fields`: extract kWh, standing charge, invoice number
- `use_acc_filter`: filter by account number
- `min_amount`: minimum amount threshold
- `analysis_min`: threshold for analysis tabs
- `use_dedup`: enable cross-source deduplication
- `use_domain_filter`: filter PST emails by sender domain
- `report_sections`: list of registry keys to include in PDF/DOCX reports — if absent or empty, every section is included

## License

MIT License — see `LICENSE` for details.

## Disclaimer

This tool was created for personal use in an EDF billing dispute. It is provided as-is without warranty. Always verify extracted data against original documents before using in any formal dispute.

## Release & Test Status

Goes green on the full CI matrix (Python 3.10 / 3.11 / 3.12 ×
ubuntu / windows / macos) — see `.github/workflows/ci.yml`.  The
local test baseline runs `pytest -v`, `ruff check .`, `ruff format
--check .`, and `mypy .` end-to-end; the "Test Status" code path is
documented in *Development* above.  See
[CHANGELOG.md](CHANGELOG.md) for the per-pass audit trail.
