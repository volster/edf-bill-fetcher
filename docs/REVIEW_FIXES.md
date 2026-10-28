# Application review fixes — historical record

[Documentation index](README.md)

This records the earlier R01–R20 application review and the checks performed at
that time. It is not a current feature inventory or certification of the latest
revision. The Tk implementation described in R19 has since been removed; the
supported desktop uses Qt. See [architecture](ARCHITECTURE.md) for current
contracts, [development](DEVELOPMENT.md) for checks, and [CI](CI.md) for the
platform workflow.

The review identified errors in monetary interpretation, evidence identity, calculations, ingestion and run handling. These changes prioritize preserving observations and making unavailable information explicit.

| Finding | Change |
| --- | --- |
| R01 — balances treated as transactions | Added explicit balance/transaction fields and signed movements; summaries and payment figures use actual transactions, including statement charges and refund offsets. |
| R02 — daily rate divided by 30 again | Billed daily rates use charges and consumption days directly. Period Charges uses bill charges; balance movements are labelled separately. |
| R03 — output surfaces use different preparation | Excel and all document formats share normalization, duplicate resolution, analysis selection and the recorded valuation date. JSON preserves the same canonical records and configuration. |
| R04 — unsafe deduplication and winner selection | Completeness sorts first. Conflicting invoices, ledgers and amounts remain distinct. Repeated transactions need a transaction identity or source-row identity. Hiding the duplicates tab does not change deduplication. |
| R05 — admission/diagnostic fields dropped | Preparation preserves all original fields, admission flags and source occurrences. Workbook evidence exposes admission and identity fields; filtered JSON retains full records. |
| R06 — reconciliation counts, values and links wrong | Matching requires comparable values, dates and compatible identity, consumes each match once, and uses source sheet coordinates. Unavailable contract/register values are unassessed. |
| R07 — report and evidence bundles overwritten | Each GUI run reserves a unique directory shared by its workbook, reports, source copies, index and state. Imported-workbook reports also get unique directories. |
| R08 — overlapping interest and compensation | Interest is opt-in and accrues once over changing credit principal. Refunds require transaction evidence. Back-billing allocations use the union of live, unsuperseded consumption slices. |
| R09 — whole-bill rates and issue-date cap comparison | Only extracted energy rates are compared, split by consumption quarter and weighted by consumption/time. Context and missing historical data remain unassessed; results describe a reference benchmark. |
| R10 — account filtering bypassed | Exact normalized account selection applies to billing and SAP paths. Account/meter/fuel identity is retained and analytical ledgers are separated. |
| R11 — EML/MSG attachments ignored | Adapters return attachment bytes separately from bodies. Supported PDFs use the common ingestion and evidence-copy pipeline; unsupported attachments are counted and logged. |
| R12 — basename collisions and invented links | Sources have origin/content identities and hashes. Colliding names remain distinct. Workbook links refer only to copied files; page ranges and original names are retained. |
| R13 — inconsistent date interpretation | Scalar and vector parsing share explicit ISO and UK rules. Missing dates remain missing rather than becoming a synthetic year/date. |
| R14 — SAP-only/single-record outputs suppressed | SAP data can produce a workbook without billing records. Small datasets retain useful evidence and analysis tabs with clear insufficient-data results. |
| R15 — partial reports reported as success/state metadata lost | Failed sections fail publication. Atomic writes preserve prior complete files. JSON reload restores configuration, counters and filtered evidence; GUI run status records failure/cancellation/review requirements. |
| R16 — consumption treated as meter registers | Register readings have a distinct field. Consumption cannot trigger a rollover. Unknown reading classification remains unknown. |
| R17 — missing/nonfinite data marked complete | Completeness checks use valid dates, finite numbers and recognized reading labels; duplicate quality counts use actual duplicate decisions. |
| R18 — irregular observations forecast as monthly | Forecasts use consecutive month-end balances, retain calendar month dates and reject gaps/multiple ledgers. MAPE remains a fraction for percentage formatting. Bands are labelled advisory. |
| R19 — mutable Tk state and shutdown hazards | Configuration and paths are captured before work. Workers use queued UI callbacks and non-daemon threads. Window close cancels and waits for cleanup. |
| R20 — spreadsheet formula injection | External evidence, SAP, reconciliation, filtered/error and diff values are written as literal cells. App-created workbook formulas remain separate. |

The section registry now lives independently of the PDF renderer. Dependency ranges agree between the project and convenience requirements, package discovery excludes tests, and executable builds include MSG support and resolved dependency metadata. The architecture document describes actual dependencies and report flow.

Missing account identity, unavailable cumulative readings, unknown tariff applicability and incomplete historical cap data remain visible limitations of the source evidence. The app preserves them without inventing financial or legal conclusions.

## Validation at the time of the review

Local verification passed linting, formatting and type checking across 222 source files, 534 regression checks in the changed test files, and a wheel build whose contents exclude tests. Functional checks exercised fixture PDF extraction, evidence copying/indexing, version 2 JSON and report reloads, XLSX/PDF/DOCX/HTML generation, and SAP-only workbooks.

A real Tk worker run verified that changing controls during extraction does not change its captured inputs or output location. A second run preserved the complete first bundle, and window shutdown allowed worker cleanup to finish. Those local checks did not establish a full operating-system/Python matrix result.
