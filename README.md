# EDF Energy Billing Evidence Collector

A desktop application for collecting and analyzing EDF Energy billing data from multiple sources (PST/OST email archives, PDF bills, HTM account exports) to produce a comprehensive Excel evidence report and professional PDF report for Energy Ombudsman dispute resolution.

## Features

- **Multi-source extraction**: PST/OST email files, local PDF folders, HTM account history exports
- **Dual format support**: Parses both old-style and new-style (KI/KCR) EDF invoice formats
- **Smart amount detection**: Multiple regex patterns with fallback strategies for finding billing amounts
- **Deduplication**: Cross-source deduplication using billing period dates and amounts
- **Comprehensive Excel output**: 14-sheet report with evidence, summaries, charts, and dispute analysis
- **Professional PDF output**: 16-section Ombudsman-ready report with executive summary, evidence index, timeline, OFGEM comparison, statistical analysis, and methodology appendix
- **GUI interface**: Simple tkinter-based desktop application with progress tracking
- **CLI mode**: Headless PDF report generation for automation

## Output Sheets

| Sheet | Description |
|-------|-------------|
| **Annual Summary** | Yearly balance range, average, peak, low |
| **EDF Evidence Report** | All extracted records with live formulas |
| **Duplicate Entries** | Deduplicated records (if enabled) |
| **Filtered (Below Min)** | Records below minimum threshold |
| **Parse Errors** | Any extraction errors encountered |
| **Key Statistics** | Account overview, balance figures, periodic charges, reading quality, unit rates |
| **Balance Trend** | Time-series chart with rolling average and linear trend |
| **Year-on-Year** | Yearly comparison with YoY changes |
| **Period Charges** | Per-period charges with daily rates and flags |
| **Dispute Flags** | Automated detection of anomalies (large jumps, billing gaps, estimated runs, reconciliation mismatches) |
| **Dispute Timeline** | Chronological event timeline for dispute narrative |
| **Statistical Analysis** | Descriptive stats, rolling 6-period stats, EMA, momentum, volatility, z-score/IQR anomalies, Shapiro-Wilk normality tests |
| **Payment Analysis** | Payment/credit patterns, intervals, amounts, chronological detail with chart |
| **Forecast & Projection** | Linear regression, Holt-Winters exponential smoothing, EMA projection, confidence intervals, accuracy metrics |

## PDF Report Sections

The PDF report (generated via **EXPORT PDF REPORT** button or CLI) contains:

1. **Cover Page** — Account reference, period, confidentiality notice
2. **Table of Contents** — Full section listing
3. **Executive Summary** — Financial totals, key findings narrative, conclusion for Ombudsman
4. **Key Findings** — Severity-grouped table (HIGH/MEDIUM/INFO) with detail
5. **Evidence Index** — Source cross-reference (PST/HTM/PDF) with record counts
6. **Detailed Anomalies** — Large Jumps, Billing Gaps, Estimated Runs, High Daily Rates, Reconciliation Mismatches
7. **Timeline** — Chronological merged view of all bills + automated flags
8. **OFGEM Price Cap Comparison** — Template for comparing bill unit rates against OFGEM caps
9. **Statistical Analysis** — Descriptive stats, rolling stats, normality tests
10. **Payment Analysis** — Payment patterns, intervals, chronology
11. **Forecast & Projection** — Linear, Holt-Winters, EMA projections with confidence intervals
12. **Data Quality Assessment** — Completeness, duplication, source distribution, pass/warn/fail status
13. **Tariff Impact Analysis** — Unit rate stats by tariff, change detection, charge impact
14. **Appendix A: Methodology** — Data sources, extraction logic, deduplication, config used
15. **Appendix B: Glossary** — 10 key terms (MAPE, Holt-Winters, IQR Anomaly, etc.)
16. **Professional Styling** — Navy/orange theme, headers/footers, page numbers, confidentiality headers

The report is designed for direct submission to the **Energy Ombudsman**.

## Installation

### From Source (Development)

```bash
git clone https://github.com/volster/edf-bill-fetcher.git
cd edf-bill-fetcher
pip install -e ".[dev]"
```

### Build Executable

```bash
pip install -e ".[build]"
pyinstaller --onefile --windowed --name EDF_Evidence_Collector edf_collector.py
```

The executable will be in `dist/`.

## Usage

### GUI Mode

```bash
python edf_collector.py
```

Or run the built executable `EDF_Evidence_Collector.exe`.

1. **Select Sources** (at least one required):
   - **PST/OST File**: Outlook email archive containing EDF emails
   - **PDF Folder**: Directory containing EDF bill PDFs
   - **HTM Export**: EDF MyAccount "Payments and Invoices" HTM export

2. **Configure Options**:
   - **Account Filter**: Filter by EDF account number (A-XXXXXXXX)
   - **Domain Filter**: Filter PST emails by sender domain (default: edfenergy.com)
   - **Minimum Amount**: Filter out records below this threshold (default: £500)
   - **Analysis Threshold**: Minimum bill amount for analysis tabs (default: £500)
   - **Report Account Ref**: Override account reference in report header

3. **Click "EXTRACT TO EXCEL"** — produces the 14-sheet Excel workbook

4. **Click "EXPORT PDF REPORT"** — produces the professional Ombudsman-ready PDF (enabled after extraction completes)

The report will be saved next to your source files.

### CLI Mode — PDF Report Generation

For automated/pipeline generation of the PDF report:

```bash
# Generate PDF from previously extracted records
python edf_collector.py --pdf-report -i records.json -o report.pdf

# With optional config and engine data
python edf_collector.py --pdf-report -i records.json -o report.pdf -c config.json -e engine.pkl
```

Options:
- `-i, --records`: Path to extracted records JSON (required)
- `-o, --output`: Output PDF file path (required)
- `-c, --config`: Config JSON file (optional)
- `-e, --engine-data`: Pickled EvidenceEngine for filtered records (optional)

### Programmatic Usage

```python
from edf_collector import EvidenceEngine, export_to_excel
from edf_report import generate_pdf_from_gui

config = {
    'use_anchors': True,
    'use_large': True,
    'use_reading_classification': True,
    'use_pdf_fields': True,
    'use_acc_filter': False,
    'acc_num': '',
    'min_amount': 500.0,
    'analysis_min': 500.0,
    'filter_below': True,
    'save_filtered': True,
    'use_dedup': True,
    'save_dups': True,
    'use_domain_filter': True,
    'domain_filter': 'edfenergy.com',
}

engine = EvidenceEngine(config, print)
engine.crawl_local_pdfs('/path/to/pdfs')
# ... process other sources ...

# Excel export
export_to_excel(engine.records, 'output.xlsx', engine.error_log, config, engine.filtered_records)

# PDF export
generate_pdf_from_gui(
    records=engine.records,
    output_path='report.pdf',
    config=config,
    engine=engine,
    filtered=engine.filtered_records,
)
```

## Supported EDF Formats

### New-Style Invoices (KI-XXXXXXXX)
- "Current balance £X debit"
- "Total charges for this period £X debit"
- "Your charges: DD Mon YYYY - DD Mon YYYY"
- kWh usage, standing charge, tariff name

### New-Style Credit Notes (KCR-XXXXXXXX)
- "Total credits for this bill £X"

### Old-Style Bills
- "Your new account balance £X"
- Generic amount patterns with "balance", "total charges", "amount to pay"

### HTM Account History
- "DD Mon YYYY We charged your account £X For Y kWh ... Balance £Z in debit"
- "DD Mon YYYY You paid us £X ... Balance £Z in debit"
- "DD Mon YYYY Reversed account charge £X ... Balance £Z in debit"

## Requirements

- Python 3.10+
- tkinter (usually included with Python)
- Dependencies (see `pyproject.toml`):
  - pandas
  - pdfplumber
  - beautifulsoup4
  - openpyxl
  - numpy
  - libpff-python (optional, for PST/OST support)
  - reportlab (optional, for PDF report generation)

### Optional Extras

```bash
# For PST/OST support
pip install -e ".[pst]"

# For PDF report generation
pip install -e ".[pdf]"

# For development (test, lint, build)
pip install -e ".[dev]"

# All extras
pip install -e ".[pst,pdf,dev]"
```

## Development

### Run Tests

```bash
pip install -e ".[dev]"
pytest -v
```

### Linting & Formatting

```bash
ruff check .
ruff format .
```

### Type Checking

```bash
mypy .
```

## Configuration

All options are available in the GUI. For programmatic use, see the `config` dict in the usage example above.

Key config options:
- `use_anchors`: Enable smart context amount patterns
- `use_large`: Enable large amount fallback
- `use_reading_classification`: Classify readings (Estimated/Actual/Smart)
- `use_pdf_fields`: Extract kWh, standing charge, invoice number
- `use_acc_filter`: Filter by account number
- `min_amount`: Minimum amount threshold
- `analysis_min`: Threshold for analysis tabs
- `use_dedup`: Enable cross-source deduplication
- `use_domain_filter`: Filter PST emails by sender domain

## License

MIT License - see LICENSE file for details.

## Disclaimer

This tool was created for personal use in an EDF billing dispute. It is provided as-is without warranty. Always verify extracted data against original documents before using in any formal dispute.