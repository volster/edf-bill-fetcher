# EDF Energy Billing Evidence Collector

A personal desktop application that collects and analyses EDF Energy billing data from PST/OST email archives, local PDF bills, and HTM account history exports. Produces a multi-sheet Excel evidence workbook and a PDF **or** DOCX report suitable for Energy Ombudsman submission.

## Features

- **Multi-source extraction**: PST/OST email files, local PDF folders, HTM account exports.
- **Dual format support**: Parses both old-style and new-style (KI/KCR) EDF invoice formats.
- **Smart amount detection**: Prioritized regex patterns with configurable fallback.
- **Cross-source deduplication**: Two-pass dedup — Period To + Amount (primary), Amount within a 60-day window (secondary). Within a duplicate cluster the *most complete* row wins (substantive field fill-rate), with source precedence as the tie-breaker. Set `amalgamate_duplicates=True` to merge columns across all duplicate siblings into a single hybrid kept row (each sibling still surfaces on the Duplicate Entries sheet for audit).
- **Comprehensive Excel output**: Multi-sheet evidence workbook with annual summary, dispute flags, statistical analysis, payment analysis, and forecast sheets — see the *Output Sheets* section for the full conditional-emission list.
- **SAP financial-ledger integration**: when `scan_sap_dumps` is on (default), the engine also digests EDF's three SAP-CSV-in-PDF dumps (`*_Contract-History.pdf`, `*_Meter-Readings.pdf`, `*_Financial-Transactions.pdf`) and renders three source sheets + two analyser sheets (`SAP Back-billing Events` and `SAP ↔ EDF Matched Events`) + a cross-source `Reconciliation` sheet. See the *SAP ledger integration* section below for the full design.
- **Professional PDF + DOCX output**: 14 dynamically-numbered sections. Numbering is **derived from `REPORT_SECTIONS` so the Table of Contents and body always agree**, regardless of which sections a user selects in the report options dialog.
- **GUI interface**: tkinter-based desktop application with progress tracking.
  - **Output Folder picker** (Section 1): choose where xlsx + report outputs land; empty falls back to the source-file directory.
  - **Sequential file naming**: `<stem>_<YYYY-MM-DD>_<N>.xlsx` (and `_Report.pdf` / `_Report.docx` variants). Counter is per-day per-folder and shared across all outputs in one batch.
  - **Auto-generate report**: a checkbox in Section 2 to produce xlsx + PDF + DOCX from a single EXTRACT run, sharing the same batch counter.
  - **Three-state EXTRACT button**: `EXTRACT TO EXCEL → Cancel → Cancelling... → EXTRACT TO EXCEL` — theEXTRACT and Cancel buttons are now collapsed into one.
  - **Config persistence**: GUI state and report options stored at `~/.edf_collector/config.json` (atomic write, `0o600`); deleting the file resets state cleanly.
  - **Amalgamate toggle**: the `amalgamate_duplicates` checkbox surfaces as a nested child under Deduplication (enabled only when both *Drop duplicates* and *Record dropped duplicates* are on; default OFF).
- **CLI mode**: Headless batch/report generation for automation.

## Package layout (post-modularization)

The codebase has been refactored from a single 10,645-line `edf_collector.py` monolith into a hexagonal `edf_bill_fetcher/` package. The full layout with layering rules is documented in [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md); the short version:

| Subpackage             | Layer            | Responsibility                                                                 |
| ---------------------- | ---------------- | ------------------------------------------------------------------------------ |
| `edf_bill_fetcher.helpers`      | pure stdlib       | Date math, formatting, theme constants, excel helpers                          |
| `edf_bill_fetcher.processors`   | stdlib + pandas   | Business logic: dedup, pattern detection, reconciliation, forecasting        |
| `edf_bill_fetcher.collectors`   | framework boundary| `EvidenceEngine` orchestrator; ingests PDF / HTM / PST                          |
| `edf_bill_fetcher.io.writers`   | openpyxl          | Excel sheet builders — one submodule per sheet (13 writers)                    |
| `edf_bill_fetcher.io.reporters` | reportlab + docx  | PDF + DOCX report renderers                                                    |
| `edf_bill_fetcher.io.adapters`  | libpff-python      | PST/OST archive adapter                                                        |
| `edf_bill_fetcher.ui`            | tkinter           | GUI: `App`, `ReportOptionsDialog`                                              |
| `edf_bill_fetcher.models`       | dataclasses       | Event / data models                                                            |
| `edf_bill_fetcher.writers`      | internal facade  | Re-exports `writers._helpers` analysis helpers; canonical writer entry points live in `edf_bill_fetcher.io.writers` |

Backward-compat note: the legacy `edf_collector.py` / `edf_report.py` / `edf_report_docx.py` modules at the repo root were deleted in the modularization (no top-level shims remain — existing scripts must import from `edf_bill_fetcher.*`). The temporary PEP 562 re-export layers were removed after consumers migrated; writer functions import from the canonical `edf_bill_fetcher.io.writers` path (or the per-sheet submodule, e.g. `edf_bill_fetcher.io.writers.export`).

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
#   [dev]    test + lint + typecheck toolchain (pytest, pytest-xvfb,
#            pytest-cov, ruff, mypy)
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
pip install -e ".[dev]"     # bring in pytest / pytest-xvfb / pytest-cov / ruff / mypy only
pip install -e ".[build]"   # add PyInstaller
```

## Usage

### GUI Mode

```bash
# Either via the installed console-script entry point:
edf-collector

# Or via the main.py launcher at the repo root:
python main.py
```

Or run the built executable `EDF_Evidence_Collector.exe`.

1. **Select Sources** (at least one required):
   - **PST/OST File**: Outlook email archive containing EDF emails (`.pst` / `.ost` are read via `libpff-python`; the `process_pst_file` wrapper auto-logs an error if the lib is missing, so the rest of the pipeline still runs)
   - **PDF Folder**: Directory containing EDF bill PDFs
   - **HTM Export**: EDF MyAccount "Payments and Invoices" HTM export
2. **Configure Options**:
   - **Account Filter**: Filter by EDF account number (e.g., `A-12345678` or `123 456 789 012`). Both compact and grouped-digit renderings are matched against the bill — `extract_new_invoice_fields` accepts the spaceless `A-NNNNNNNN` shape and the bank-row-style `NNN NNN NNN NNN` shape.
   - **Domain Filter**: Filter PST emails by sender domain (default: edfenergy.com)
   - **Minimum Amount**: Filter out records below this threshold (default: £50)
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
python main.py --pdf-report -i records.json -o report.pdf

# DOCX variant
python main.py --docx-report -i records.json -o report.docx
```

Pass `-c config.json` and `-e engine.pkl` to forward config + filtered-records state. The CLI dispatch logic lives in `edf_bill_fetcher.io.cli`; the `main()` entry point at the repo root is a one-line re-export of `edf_bill_fetcher.io.cli.main`.

### Programmatic Usage

> Per-source API is symmetric — all three source types expose
> `process_<source>_file(path, source_label, detail_label, fallback_date)`
> so you can plug in any combination via the same
> call signature. PST requires `libpff-python` (a runtime dep
> since `0.1.0+`); if it's missing in some hand-built
> environment, the wrapper logs the error and the rest of the
> pipeline still runs.
>
> The canonical post-refactor import paths are shown below; all
> writer functions import from `edf_bill_fetcher.io.writers`
> (the flat `edf_bill_fetcher.writers` facade only re-exports the
> `writers._helpers` analysis helpers).

```python
from edf_bill_fetcher.collectors import EvidenceEngine
from edf_bill_fetcher.io.writers.export import export_to_excel
from edf_bill_fetcher.io.reporters.pdf_report import generate_pdf_from_gui
from edf_bill_fetcher.io.reporters.docx_report import generate_docx_from_gui

config = {
    "use_anchors": True,
    "use_large": True,
    "use_reading_classification": True,
    "use_pdf_fields": True,
    "use_acc_filter": False,
    "acc_num": "",
    "min_amount": 50.0,
    "analysis_min": 500.0,
    "filter_below": False,
    "save_filtered": True,
    "use_dedup": True,
    "save_dups": True,
    "use_domain_filter": True,
    "domain_filter": "edfenergy.com",
    # report_sections tells the report generator which sections to include.
    # If absent, every section in `edf_bill_fetcher.io.reporters.pdf_report.REPORT_SECTIONS` is selected.
    "report_sections": [
        "exec_summary",
        "key_findings",
        "evidence_index",
        "detailed_findings",
        "timeline",
        "ofgem",
        "statistical",
        "payment",
        "forecast",
        "data_quality",
        "tariff",
        "appendix_methodology",
        "appendix_glossary",
        "appendix_full_evidence",
    ],
}

engine = EvidenceEngine(config, print)
engine.process_pdf_file(
    "path/to/a.pdf", source_label="Local PDF", detail_label="bill.pdf", fallback_date="2026-03-01"
)
engine.process_htm_file("path/to/export.htm", fallback_date="2026-03-01")
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

The PDF and DOCX reports are both built from a single section-registry so the titles and numbering always line up. The registry lives in `edf_bill_fetcher.io.reporters.pdf_report.REPORT_SECTIONS`:

| Class          | Sections                                                                                                                                                                                                              |
| -------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Main (numeric) | Executive Summary · Key Findings Summary · Evidence Index & Source Cross-Reference · Detailed Findings · Timeline of Events · OFGEM Price Cap Comparison · Statistical Analysis · Payment & Credit Analysis · Forecast & Projection · Data Quality Assessment · Tariff Impact Analysis |
| Appendix       | Methodology & Data Sources · Glossary · Full Evidence Table                                                                                                                                                                                                              |

Main sections are numbered **1, 2, 3, …** and appendices are lettered **A, B, C, …**, computed at render time based on the user's `report_sections` selection.

A `tests/test_dispatch_parity.py` structural test pins the invariant that the PDF dispatcher's `section_builders`, the DOCX dispatcher's `section_builders`, and `REPORT_SECTIONS` all expose exactly the same set of keys — so a future contributor who adds a section to the registry without wiring both dispatchers breaks CI, not the rendered report.

### Adding a new section

1. Add an entry to `REPORT_SECTIONS` in `edf_bill_fetcher/io/reporters/pdf_report.py` with `key`, `title`, and optionally `is_appendix`.
2. Add the matching key to `ReportOptionsDialog.SECTIONS` in `edf_bill_fetcher/ui/app.py` so it shows up in the GUI options dialog.
3. Add a `def create_<name>(...)` builder function in `edf_bill_fetcher/io/reporters/pdf_report.py`.
4. Wire the builder into the `section_builders` dispatch dict in **both** `generate_ombudsman_pdf` and `generate_ombudsman_docx`. Forgetting this raises a clear `RuntimeError` at report-render time — that's the loud-failure mode that keeps the dispatch in lockstep with the registry.

Removing a section: same steps in reverse.

## Output Sheets (Excel)

Sheets are written in roughly this order; conditional sheets only appear when their gating input is non-empty and the relevant toggle is on:

| Sheet | When emitted | Description |
| --- | --- | --- |
| **Provenance** | always (first tab) | Tool version, generation timestamp, account reference, record counts, and a full snapshot of the run configuration — so a filed workbook is self-documenting about how its conclusions were produced |
| **Annual Summary** | always | Yearly balance range, average, peak, low |
| **EDF Evidence Report** | always | All extracted records with live formulas |
| **Duplicate Entries** | `use_dedup and save_dups` (default on/ off respectively) | Deduplicated records, retained for audit so the dropped siblings are visible |
| **Filtered (Below Min)** | `filter_below and save_filtered` (default off/on) when there *are* records whose **absolute** amount is below the minimum-threshold | Records whose `abs(Amount)` is below the minimum threshold (refunds like `-£1000` are KEPT in the main report — only small-magnitude amounts are shelved here) |
| **Parse Errors** | when the engine captured one or more extraction errors | Any extraction errors encountered |
| **Key Statistics** | always (when ≥2 records survive dedup) | Account overview, balance figures, periodic charges, reading quality, unit rates |
| **Balance Trend** | always (≥2 records) | Time-series chart with rolling average and linear trend |
| **Year-on-Year** | always (≥2 records) | Yearly comparison with YoY changes |
| **Period Charges** | always (≥2 records) | Per-period charges with daily rates and dispute flags |
| **Dispute Flags** | always (≥2 records) | Automated detection of anomalies — large jumps, billing gaps, estimated runs, reconciliation mismatches |
| **Dispute Timeline** | always (≥2 records) | Chronological event timeline for the dispute narrative |
| **Statistical Analysis** | always (≥2 records) | Descriptive stats, rolling 6-period stats, EMA, momentum, volatility, z-score/IQR anomalies, Shapiro-Wilk normality tests |
| **Payment Analysis** | always (≥2 records) | Payment/credit patterns, intervals, amounts, chronological detail with chart |
| **Forecast & Projection** | always (≥2 records) | Linear regression, Holt-Winters exponential smoothing, EMA projection, confidence intervals, accuracy metrics |
| **Data Quality Report** | always | Completeness and read-quality metrics |
| **Tariff Analysis** | always (≥2 records) | Tariff impact analysis |
| **Back-billing Analysis** | always (≥2 records) | Invoices whose `Period From` → `Period To` window exceeds the SLC 7A 12-month limit (365 days), with the cancel/rebill admission disclosed as `Admitted phrase`, `Period overlap`, `Admitted + overlap`, or blank. Legal context block cites Electricity Act 1989 s.84B and Ofgem's back-billing rule. |
| **Rebilling & Corrections** | always (≥2 records) | Killer/killed invoice pairs identified by period overlap > 30 days OR jump-back > 30 days OR long-period killer ≥ 60 days reaching back into a prior invoice's window. Trigger Reason lists every matching heuristic. |
| **Meter Readings** | always (≥2 records) | Actual vs Estimated timeline with meter-rollover candidates flagged `M`. Estimated Source mirrors the row's `Details` column (e.g. `Automatic estimate`). |
| **Contract History** | always (≥2 records) | Contract periods inferred from tariff transitions, with ≤ 30-day gaps merging across same-tariff runs. |
| **SAP Contract History** | `scan_sap_dumps` (default on) and a SAP Contract History PDF is supplied | Contract periods extracted from the SAP Contract History PDF dump |
| **SAP Meter Readings** | `scan_sap_dumps` and a SAP Meter Readings PDF is supplied | Meter readings from the SAP Meter Readings PDF dump, parsed by a multi-regex fallback chain |
| **SAP Financial Transactions** | `scan_sap_dumps` and a SAP Financial Transactions PDF is supplied | Per-row extract of every financial transaction from the SAP Financial Transactions PDF. The widened parser surfaces **26 columns** per row (the 16 historically-surfaced columns plus 10 analyser-relevant extensions: Contract, Sub Item, Clearing Posting Date, Clearing Amount, Statistical Key Flag, Tax Code / Tax Code Description, G/L Account / G/L Description, Deferral Date). |
| **SAP Back-billing Events** | `scan_sap_dumps` and a SAP Financial Transactions PDF is supplied | One summary row per **Clearing Document cluster** of ≥4 underlying SAP rows, sorted by Clearing Date ascending. Each summary row carries # rows, net amount, has-credit-for-consum-billing flag, largest single posting, posting-date range, an evidence trail narrative, and a hyperlink to its first underlying row on `SAP Financial Transactions`. Underlying rows are written as collapsible outline-group sub-rows beneath each summary (default collapsed); click `+` in the left margin to expand. Debt-management rows (`Statistical Key Flag == 'Installment Plan Item'`) are filtered out of clustering. See the design spec at `scratch/Docs/Superpowers/Specs/2026-07-21-sap-back-billing-analysis-design.md`. |
| **SAP ↔ EDF Matched Events** | `scan_sap_dumps` and a SAP Financial Transactions PDF is supplied | One row per SAP back-billing event × EDF invoice candidate pair whose fuzzy match confidence is Low or better. Confidence is computed by the §3.3 algorithm in the spec: date score (in-span = 50, within 3 days = 25, within 14 days = 5) + amount score (within 5% = 40, within 25% = 20, within 50% = 5); net-zero clusters match any EDF invoice whose gross equals some cluster row's gross within the same bands → bands High (≥75) / Medium (≥40) / Low (≥10) / unmatched (omitted). Bidirectional hyperlinks: SAP Clearing Doc cell links to the event on the previous sheet, EDF Invoice # cell links to the row on `EDF Evidence Report`. |
| **Reconciliation** | `scan_sap_dumps` and `generate_reconciliation_sheet` (independent toggle, default on) | Cross-source reconciliation between SAP-ledger signals (clearing clusters, credit-for-consum-billing postings) and EDF-invoice signals (back-billing / rebilling / meter-rollover detections). Toggled separately so a reviewer can keep just the SAP-ledger sheets without the cross-source view. |

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

The detectors live in `edf_bill_fetcher.processors.detection` and `edf_bill_fetcher.processors.matching`. They can be called directly from Python:

```python
from edf_bill_fetcher.processors.detection import (
    detect_back_billing,
    detect_rebilling,
    detect_meter_rollover,
)
from edf_bill_fetcher.processors.matching import infer_contracts
from edf_bill_fetcher.io.writers.analysis import run_analysers

# One-shot — returns a dict with keys back_billing, rebilling,
# meter_rollover, contracts.
analyses = run_analysers(deduped_df)
```

Multi-invoice (merged) PDF inputs — e.g. a bundle of 30+ pages containing 8 T-series invoices or 40 pages containing 10 KI-series invoices — are sliced at invoice boundaries before parsing, so each invoice gets its own row instead of being lost in the whole-file concat. The slicer looks for either an `Invoice number:` line or a `Page 1 of N` boundary marker (variants `1 of 4`, `one of four`, `1/4`).

## SAP ledger integration

EDF's SAP ledger PDFs are the *behind-the-scenes* truth of every billing event; the EDF-branded invoice PDFs only show what EDF chose to send the customer. The three SAP source PDFs the engine looks for (under the `Scan SAP dumps` GUI toggle, default on, or `config["scan_sap_dumps"]` programmatic key) are:

- **SAP Contract History** (`*_Contract-History.pdf`) → `SAP Contract History` sheet
- **SAP Meter Readings** (`*_Meter-Readings.pdf`) → `SAP Meter Readings` sheet
- **SAP Financial Transactions** (`*_Financial-Transactions.pdf`) → `SAP Financial Transactions` sheet

The Financial-Transactions parser surfaces **26 columns** per row (16 historical + 10 analyser extensions).

Two analyser sheets sit adjacent to `SAP Financial Transactions`, deriving their data from it:

- **SAP Back-billing Events** — one summary row per Clearing Document cluster of ≥4 underlying rows. Each cluster's net amount, evidence trail narrative, and link back to its underlying rows on `SAP Financial Transactions`. Sub-rows are hidden under collapsible outline groups (default collapsed) so the sheet opens at a readable ~62 rows. Debt-management rows whose `Statistical Key Flag == 'Installment Plan Item'` are excluded from clustering up-front.
- **SAP ↔ EDF Matched Events** — fuzzy-join Sheet 1 events to EDF invoice records. Each matched pair lists the date delta vs EDF `Period To`, the SAP event's net amount and the EDF invoice's amount, a confidence band, and bidirectional hyperlinks to Sheet 1 and `EDF Evidence Report` rows. Per spec §3.3 the algorithm scores on two axes:
  - **date** — Clearing Date inside EDF `[Period From, Period To]` = 50, within 3 days = 25, within 14 days = 5
  - **amount** — for net-£0 clusters, the closest cluster row's gross within 5% of EDF amount = 40 / within 25% = 20 / within 50% = 5; for non-zero net clusters, the same bands applied to `event.net / edf.amount`

  Bands: **High** ≥ 75, **Medium** ≥ 40, **Low** ≥ 10. SAP events that produce no candidate at Low-or-above are omitted from Sheet 2 (but remain on Sheet 1).

Both new sheets honour the same `scan_sap_dumps` toggle as the existing three SAP sheets — there is no separate "Show SAP back-billing" checkbox (YAGNI).

The full design spec is preserved at
`scratch/Docs/Superpowers/Specs/2026-07-21-sap-back-billing-analysis-design.md`.

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
  pytest, pytest-xvfb, pytest-cov, coverage, ruff, mypy, pyinstaller.

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

The tkinter-dependent tests require a display server. On headless Linux CI / dev environments, the `pytest-xvfb` plugin (in the `dev` extras) auto-activates a virtual X server — install with `pip install -e ".[dev]"` and the GUI tests run green without any manual xvfb setup. The plugin is **opt-out** (use `--no-xvfb` to disable).

### Linting / formatting / type-checking

```bash
ruff check .
ruff format .
mypy .
```

All three are enforced in CI on Python 3.10 / 3.11 / 3.12 ×
ubuntu / windows / macos — see `.github/workflows/ci.yml`. A
release is one CI green away from shippable.

### Test coverage

Coverage is measured with `coverage` (configured in `pyproject.toml [tool.coverage.run]`) over the `edf_bill_fetcher/` source tree. The CI gate is **`fail_under = 90`** in `pyproject.toml [tool.coverage.report]` — CI runs `pytest --cov=. --cov-report=xml` (Linux wraps it in `xvfb-run -a`) — see [`docs/COVERAGE.md`](docs/COVERAGE.md) for the measurement protocol, the strict `# pragma: no cover` policy, and how to extend coverage.

## Configuration

All options are available in the GUI. For programmatic use, see the `config` dict in the usage example above. The configuration contract is typed as `edf_bill_fetcher.models.config.ConfigDict` (a `TypedDict` — every key optional, documented with its default in one place); passing a dict literal with a misspelled key or wrong value type to any consumer is a `mypy` error, and the GUI/CLI boundary normalises the legacy short key names (`use_acc_filt`, `use_reading_class`) to the canonical long ones.

Key options:
- `use_anchors`: enable smart-context amount patterns
- `use_large`: enable large-amount fallback
- `use_reading_classification`: classify Estimated/Actual/Smart readings
- `use_pdf_fields`: extract kWh, standing charge, invoice number
- `use_acc_filter`: filter by account number
- `min_amount`: minimum amount threshold (compared against `abs(amount)` — high-magnitude refunds are kept in the main report)
- `analysis_min`: threshold for analysis tabs
- `use_dedup`: enable cross-source deduplication
- `use_domain_filter`: filter PST emails by sender domain
- `scan_sap_dumps`: (default `True`) emit the three SAP source sheets + the two new SAP analyser sheets when a SAP-dump PDF is supplied. See *SAP ledger integration*.
- `generate_reconciliation_sheet`: (default `True`) emit the cross-source Reconciliation sheet. Independent of `scan_sap_dumps` for cases where the SAP sheets alone are wanted.
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
