# Configuration reference

[Documentation index](README.md) · [User guide](USER_GUIDE.md) · [CLI](CLI.md)

Use **Preferences** for the next desktop extraction, **Report settings…** for
formats and sections, and `--config file.json` for command-line settings. A saved
analysis retains the settings used for that result; changing today's preferences
does not change yesterday's accepted figures.

## Where settings live

| Settings | Location and behavior |
| --- | --- |
| Desktop preferences | `~/.edf_collector/config.json` on each OS (`~` is your home folder). Saves `gui_state`, output location/name and `report_options`. Selected source paths are not persisted. |
| CLI configuration | A plain UTF-8 JSON object passed with `--config`; it is not the nested desktop preferences file. Explicit extraction flags override matching file values. |
| Accepted run settings | Stored in `result.json`; the workbook's Provenance sheet also records configuration. |
| Review decisions | The review draft, main review workbook and subsequently accepted result. These are not preference switches. |

Preferences save automatically. Close the app and rename its settings file to
reset them; keep the renamed copy if you may want to recover values. This does
not remove evidence packs or review drafts. Files are written atomically and
permissions are restricted where the OS supports them; this is not encryption.

## Extraction and source retention

These are fresh desktop/CLI defaults. Direct callers constructing
`EvidenceEngine({})` must supply their intended configuration: individual legacy
consumers have fallback defaults which can differ from the entry points.

| Key | Default | Effect |
| --- | --- | --- |
| `use_anchors` | `true` | Recognise common bill amount patterns. |
| `use_large` | `true` | Permit an amount fallback when structured bill fields are missing. Review the result against the source. |
| `use_reading_classification` | `true` | Identify actual, estimated and other supported reading labels. |
| `use_pdf_fields` | `true` | Extract invoice references, usage, rates and other supported PDF fields. |
| `save_evidence_files` | `true` | Copy sources into the pack and write the mapping/index. Turning this off reduces portability. |
| `scan_sap_dumps` | `true` | Read supported EDF SAP contract, meter and financial PDF dumps. |
| `generate_reconciliation_sheet` | `true` | Include cross-source SAP/bill reconciliation where data exists. It does not supply missing SAP inputs. |

There is no cloud OCR setting. The reader chooses existing PDF text or local
English OCR according to the page content. Native OCR discovery and missing-engine
recovery are described in [troubleshooting](TROUBLESHOOTING.md#scanned-bills-produce-little-or-no-text).

## Filters and duplicates

| Key | Default | Effect |
| --- | --- | --- |
| `use_acc_filter` / `acc_num` | `false` / empty | Require the selected account when enabled. Unidentified accounts can be rejected, so inspect account-filter notes. |
| `use_domain_filter` / `domain_filter` | `true` / `edfenergy.com` | Restrict supported email ingestion by sender. Multiple entries use commas. |
| `filter_below` / `min_amount` | `false` / `50.0` | When enabled, separate observations whose absolute transaction amount (or balance when no transaction exists) is below the threshold. |
| `save_filtered` | `true` | Show available below-threshold records on a side sheet. It does not make them active evidence. |
| `analysis_min` | `0.0` | Select bills for analysis by magnitude, using period charge with balance fallback. Keeps payments/credits and nominal old-consumption cases. Does not remove rows from the main evidence table. |
| `use_dedup` | `true` | Resolve supported duplicate events across sources. |
| `save_dups` | `true` | Display removed siblings on an audit sheet. This does not switch duplicate detection on or off. |
| `amalgamate_duplicates` | `false` | Fill missing fields from supported duplicate siblings after duplicate membership is established. |

Changing `min_amount` alone does **not** enable below-threshold filtering in the
GUI/CLI. For example, to filter magnitudes below £25 in the CLI, supply both
`"filter_below": true` and `"min_amount": 25.0` in the config file. There is a
`--no-filter-below` flag to disable it, but no positive enabling flag.
`min_amount` can also influence fallback amount selection during extraction, so
record its value even when below-threshold filtering is disabled. Large negative
credits are compared by magnitude, not discarded merely because
they are negative. Rows without a usable amount are not assumed to be zero.

A sender entry `edfenergy.com` matches that domain and its subdomains;
`person@example.com` matches that address exactly. `*.edfenergy.com` is also
accepted for the domain form. Matching is case-insensitive; arbitrary glob
patterns and regular expressions are not supported.

## Analysis and report options

Some analytical settings below are supplied through CLI JSON rather than a
desktop control. A setting stored in a result is part of that result's history,
not a promise of a matching Preferences checkbox.

| Key | Meaning |
| --- | --- |
| `as_of` | Valuation date, preferably `YYYY-MM-DD`. New desktop extractions capture today's date; reviewed packs retain their baseline date. CLI extraction supplies today's date when omitted. |
| `report_account_ref` | Display reference for report headings. Does not filter or change source account identities. |
| `report_sections` | Array of section keys below. Omitted means all registered sections. The GUI requires a usable selection; do not use an empty list to mean “all”. |
| `credit_interest_rate` | Optional annual simple-interest fraction; `0.08` means 8%, **not** `8`. Default `0.0` produces no interest scenario. |
| `credit_hold_days` | Days before the configured credit-interest scenario starts; default `90`. Must be finite and non-negative. |
| `fuel`, `payment_method`, `tariff_type`, `cap_region` | Optional reference-comparison context retained with the analysis. These do not rewrite evidence identity or turn a GB-average reference into a regional entitlement calculation. |

Thresholds and interest rates must be finite and non-negative. JSON does not use
`NaN`, infinity, comments, Python `True` or trailing commas. The typed developer
contract is [`ConfigDict`](../edf_bill_fetcher/models/config.py); it is not a
promise that every arbitrary JSON key is rejected at runtime. Check spelling
against this reference and the effective settings in the output.

Report format is chosen by the CLI operation or desktop report dialog, rather
than a `format` key in this flat config. The optional `cover` and `toc` section
keys control report framing; the numbered registry contains:

| Key | Section |
| --- | --- |
| `exec_summary` | Executive Summary |
| `key_findings` | Key Findings Summary |
| `evidence_index` | Evidence Index & Source Cross-Reference |
| `detailed_findings` | Detailed Findings |
| `timeline` | Timeline of Events |
| `ofgem` | OFGEM Price Cap Comparison |
| `statistical` | Statistical Analysis |
| `payment` | Payment & Credit Analysis |
| `forecast` | Forecast & Projection |
| `data_quality` | Data Quality Assessment |
| `tariff` | Tariff Impact Analysis |
| `compensation` | Potential Billing Adjustments |
| `appendix_methodology` | Methodology & Data Sources |
| `appendix_glossary` | Glossary |
| `appendix_full_evidence` | Full Evidence Table |

## Example CLI configuration

Save as `review-config.json` beside the checkout, then run from the selected
project environment:

```json
{
  "filter_below": false,
  "analysis_min": 0.0,
  "use_dedup": true,
  "save_dups": true,
  "save_evidence_files": true,
  "scan_sap_dumps": true,
  "generate_reconciliation_sheet": true,
  "credit_interest_rate": 0.0,
  "report_sections": ["exec_summary", "key_findings", "detailed_findings", "evidence_index"]
}
```

```console
python main.py --extract --pdf-dir "bills" -o "evidence.xlsx" --config review-config.json
```

The section selection takes effect when you subsequently generate a document
report with that configuration. The extraction command itself produces Excel.
For a v3 rerender, changing only section selection or the display account reference
reuses accepted models. Changing analytical options through the report CLI creates
an analysed result revision; it does not re-read PDFs. Re-extract originals for
parser changes, and use a reviewed pack for record/claim corrections.

Ledger grouping is derived from the supplied source identities, not an account-name
preference or report-heading override. Changing `report_account_ref` cannot merge
accounts or validate an alias. When source identity is conflicting or unassigned,
review the evidence and use the supported correction path; no hidden setting
forces all records into one account.
