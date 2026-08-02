# Architecture

This document describes the post-modularization structure of the EDF Bill Fetcher codebase. The codebase was refactored from a single 10,645-line `edf_collector.py` monolith into the `edf_bill_fetcher/` hexagonal package (tag `modularization-complete`).

## Top-level layout

```
edf-bill-fetcher/
├── edf_bill_fetcher/         # The package (canonical home for all production code)
│   ├── collectors/           # framework boundary — engine orchestrator
│   ├── processors/           # stdlib + pandas — business logic
│   ├── io/
│   │   ├── writers/          # openpyxl — Excel sheet builders (one per sheet)
│   │   ├── reporters/        # reportlab + docx — PDF + DOCX report renderers
│   │   ├── adapters/         # libpff-python — PST/OST archive adapter
│   │   ├── cli.py            # argv parsing + CLI entrypoints
│   ├── helpers/              # pure stdlib — date math, formatting, excel helpers, theme
│   ├── models/               # dataclasses — event/data models
│   ├── ui/                   # tkinter — GUI (App, ReportOptionsDialog)
│   └── writers/__init__.py   # PEP 562 lazy re-export shim for backward compat
├── edf_collector.py          # Compat shim re-exporting from edf_bill_fetcher.* — DO NOT extend
├── edf_report.py             # Compat shim re-exporting from io.reporters.pdf_report — DO NOT extend
├── edf_report_docx.py        # Compat shim re-exporting from io.reporters.docx_report — DO NOT extend
├── main.py                   # Console-script entrypoint: one-line re-export of io.cli.main
├── tests/                    # Test suite (790+ tests)
└── docs/                     # This directory
```

## Hexagonal layering

The package enforces a strict hexagonal dependency direction. Outer layers may import from inner layers; inner layers MUST NOT import from outer layers.

```
        ┌─────────────────────┐
        │      ui/            │ tkinter (outermost)
        ├─────────────────────┤
        │   io/               │ framework adapters (openpyxl, reportlab, docx, libpff-python, pickle)
        │   ├── writers/      │
        │   ├── reporters/    │
        │   ├── adapters/     │
        │   └── cli.py        │
        ├─────────────────────┤
        │   collectors/       │ framework boundary (pdfplumber, bs4)
        ├─────────────────────┤
        │   processors/       │ business logic (stdlib + pandas only — NO framework imports at module scope)
        ├─────────────────────┤
        │   helpers/          │ pure stdlib (NO third-party imports)
        ├─────────────────────┤
        │   models/           │ dataclasses + stdlib
        └─────────────────────┘
```

**Enforcement rules:**

| Layer              | Allowed imports                                  | Forbidden imports                              |
| ------------------ | ------------------------------------------------ | ---------------------------------------------- |
| `helpers/`         | stdlib only                                      | any third-party, any other layer               |
| `models/`          | stdlib + dataclasses                             | any third-party, any other layer               |
| `processors/`      | stdlib + pandas + sibling processors + helpers + models | third-party frameworks (openpyxl, reportlab, pdfplumber, libpff-python) |
| `collectors/`      | stdlib + pandas + pdfplumber + bs4 + processors + helpers | tkinter, openpyxl, reportlab            |
| `io/writers/`      | openpyxl + pandas + writers._helpers + helpers + processors | tkinter                                  |
| `io/reporters/`    | reportlab + docx + openpyxl + pandas + writers._helpers + helpers + processors | tkinter                       |
| `io/adapters/`     | third-party (libpff-python) + helpers             | anything else                                    |
| `io/cli.py`        | stdlib (argparse, pickle, importlib) + collectors + io.writers.export + io.reporters | tkinter                  |
| `ui/`              | tkinter + collectors + io.writers.export + io.reporters + processors + helpers |                  |

The `processors/` no-framework-import rule is the load-bearing constraint. It keeps business logic (dedup, pattern detection, reconciliation, forecasting) unit-testable with synthetic DataFrames — no Excel/PDF/tkinter required.

## Dual public API

The package exposes two import paths for the same symbols, kept in lockstep by PEP 562 lazy shims:

| Style                    | Example                                                       | Use when                                     |
| ------------------------ | ------------------------------------------------------------- | -------------------------------------------- |
| **Flat** (compat)        | `from edf_collector import EvidenceEngine, export_to_excel`   | Legacy scripts, backward compat               |
| **Flat-canonical**       | `from edf_bill_fetcher.writers import export_to_excel`         | New scripts — short path, package-internal   |
| **Submodule-scoped**     | `from edf_bill_fetcher.io.writers.export import export_to_excel` | New code that wants to be dependency-explicit |

The `edf_bill_fetcher.writers/__init__.py` shim exposes all 33 writer function names via PEP 562 `__getattr__` — the actual implementations live in `edf_bill_fetcher.io.writers.*`. Twin-identity tests in `tests/test_writers.py` assert that the flat-namespace re-export resolves to the same object as the canonical submodule import.

## PEP 562 lazy shims

Several `__init__.py` files use PEP 562 module-level `__getattr__` to defer re-exports to first attribute access. This pattern breaks circular imports that would otherwise occur when submodules back-ref the parent package.

Shim locations:

| File                                   | Why lazy                                                                                                                    |
| -------------------------------------- | -------------------------------------------------------------------------------------------------------------------------- |
| `edf_bill_fetcher/writers/__init__.py` | Resolves writer functions from `io.writers.*` lazily — eager import would cycle through `io.writers.export`'s 16 sibling imports. |
| `edf_bill_fetcher/io/writers/__init__.py` | Aggregator shim — deferred resolution avoids triggering writer submodules' `writers._helpers` import while `writers/__init__.py` is mid-init. |
| `edf_bill_fetcher/io/writers/back_billing.py`, `meter.py`, `rebilling.py`, `sap.py`, `export.py`, `reporters/__init__.py` | Leaf shim files with PEP 562 `__getattr__` — legacy paths still expose the function-bearing submodule names that they used to own before extraction. |

### When to use PEP 562 vs eager import

- **Use PEP 562** when two modules import each other transitively (cycle would form).
- **Use eager import** when there's no cycle risk — it gives better static-analysis support (mypy sees the actual type, not `object` via `__getattr__`).

The `# type: ignore  # noqa: F821` annotation on PEP 562-resolved names at call sites is required because static analyzers can't follow the lazy resolution. Two call sites in the codebase use this annotation today (`ui/app.py` and `io/cli.py` for `export_to_excel`-family names).

## The `EvidenceEngine` — ingestion surface

`edf_bill_fetcher.collectors.engine.EvidenceEngine` is the central orchestrator. It exposes symmetric per-source processing methods:

| Method                       | Source            | Required dep           |
| ---------------------------- | ----------------- | ---------------------- |
| `process_pdf_file(path, …)`  | local PDF bills   | pdfplumber             |
| `process_htm_file(path, …)`  | HTM account export | beautifulsoup4         |
| `process_pst_file(path)`     | PST/OST archive   | libpff-python (optional — logs + skips if missing) |

Each method appends to `engine.records` (a `list[dict]`). After ingestion, `dedup_records()` de-duplicates cross-source. `filter_records(min_amount)` applies the threshold filter. The engine carries its own `error_log: list[str]` for soft failures.

## Writers — Excel sheet builders

Each Excel sheet has its own function in `io/writers/<sheet_name>.py`:

| File                               | Builds                                                              |
| ---------------------------------- | ------------------------------------------------------------------ |
| `evidence.py`                       | `EDF Evidence Report` + `Annual Summary` + `Duplicate Entries`     |
| `statistical.py`                   | `Statistical Analysis` + `Balance Trend` + `Year-on-Year` sheets   |
| `payment.py`                       | `Payment Analysis` sheet                                            |
| `forecast.py`                      | `Forecast & Projection` sheet                                       |
| `data_quality.py`                  | `Data Quality Report` sheet                                          |
| `tariff.py`                        | `Tariff Analysis` sheet                                            |
| `back_billing.py`                 | `Back-billing Analysis` sheet                                       |
| `rebilling.py`                    | `Rebilling & Corrections` sheet                                    |
| `meter.py`                        | `Meter Readings` + `Contract History` sheets                       |
| `reconciliation.py`              | `Reconciliation` sheet + `_recon_parse_iso_date` / `_recon_amount_to_float` helpers |
| `sap.py`                          | All five SAP sheets (Contract History, Meter Readings, Financial Transactions, Back-billing Events, ↔ EDF Matched Events) |
| `export.py`                       | `export_to_excel` — the giant orchestrator (~1,600 lines) that calls each writer to render a sheet into the workbook |
| `analysis.py`                     | `run_analysers` — one-shot dispatcher that returns `{back_billing, rebilling, meter_rollover, contracts}` |

`export_to_excel` ~1,628 lines is the largest single function in the codebase. It's the orchestrator: it wires every writer into the output workbook in the correct sheet order with the correct conditional-emission gating. It intentionally sits in `io/writers/` (the openpyxl layer), not in `processors/`, because its job is purely I/O orchestration — never business logic.

## Reporters — PDF + DOCX renderers

Two parallel renderers share a single section registry:

- `io/reporters/pdf_report.py` — `generate_pdf_from_gui` + `generate_ombudsman_pdf` (reportlab-based)
- `io/reporters/docx_report.py` — `generate_docx_from_gui` + `generate_ombudsman_docx` (python-docx-based)

The registry `REPORT_SECTIONS` lives in `pdf_report.py`; the DOCX reporter imports it. A structural parity test (`tests/test_dispatch_parity.py`) locks the invariant that the registry, the PDF dispatcher's `section_builders`, and the DOCX dispatcher's `section_builders` all expose exactly the same set of keys.

## Processors — pure business logic

Located in `edf_bill_fetcher.processors`:

| Module             | Public surface                                                                     |
| ------------------ | ---------------------------------------------------------------------------------- |
| `detection.py`     | `detect_back_billing`, `detect_rebilling`, `detect_meter_rollover`, `extract_*_statement_rows` |
| `matching.py`      | `infer_contracts`, `account_number_matches`, source-precedence helpers              |
| `reconciliation.py` | `detect_reconciliation_statement`, cross-source reconciliation matcher             |
| `analysis.py`      | `compute_dispute_flags`, period-anomaly labeling, reversal matching                 |
| `patterns.py`      | `AMOUNT_PATTERNS`, `READING_PATTERNS`, regex pattern banks                           |
| `extraction.py`    | `extract_*` field-extraction functions used by `collectors/engine.py`               |
| `forecasting.py`   | Holt-Winters + linear+EMA fallback forecasting (statsmodels-guarded)               |
| `sap_parsers.py`   | Multi-regex SAP PDF parser (Contract History, Meter Readings, Financial Transactions) |

Processors receive DataFrames as function arguments — they do not pull data from any module-scope state. This is what makes them unit-testable with synthetic input.

## Compat shims — the legacy layer

The repo-root modules `edf_collector.py`, `edf_report.py`, and `edf_report_docx.py` are thin re-export shims. They exist so existing user scripts that do `from edf_collector import EvidenceEngine, export_to_excel, detect_back_billing, …` continue to work without modification.

**Maintenance rule**: do NOT extend these shim files. New code goes into `edf_bill_fetcher/`. The shims are documented in `tests/test_writers.py`, `tests/test_engine.py`, and `tests/test_collectors.py` (twin-identity tests that assert the shim re-export resolves to the same object as the canonical submodule import).

## CI gate

CI (`.github/workflows/ci.yml`) runs on Python 3.10 / 3.11 / 3.12 × ubuntu / windows / macos. The four gates:

1. `ruff check .` — PEP 8 + PEP 257 + import ordering (see `pyproject.toml [tool.ruff]` for the rule selection, including the D-rule relaxations documented in ruff config)
2. `ruff format --check .` — formatting
3. `mypy edf_bill_fetcher` — strict type check (`check_untyped_defs = true`, `disallow_incomplete_defs = true`)
4. `pytest -v` — full test suite (auto-activates `pytest-xvfb` on headless Linux for tkinter tests)

A release is one CI green away from shippable.
