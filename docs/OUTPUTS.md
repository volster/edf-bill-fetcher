# Understanding your outputs

[Documentation index](README.md) · [User guide](USER_GUIDE.md)

Start with the workbook's **Start here** sheet or the report's executive summary.
Use the original documents to check findings, and keep the entire run folder so
records remain linked to their evidence.

## Files in a pack

Names of the main workbook and reports follow your chosen pack name. Reports,
source copies and complaint outputs depend on the options selected and the
available evidence.

| File or folder | Purpose |
| --- | --- |
| Main `.xlsx` workbook | Financial overview, evidence tables, analysis and editable review inputs. |
| Main `.html`, `.pdf`, `.docx` reports | Readable versions of the accepted analysis, with selected sections and source references. HTML works offline. |
| `result.json` | The complete accepted analysis. Open this to review a saved run or regenerate reports with its original calculations and valuation date. |
| `run.json` | Run status, completed stages, artifact hashes and failure/cancellation information. |
| `records.json` | Compatibility export where produced; use `result.json` for a complete saved analysis. |
| `evidence_files/` | Saved source copies when enabled. Keep them with the pack. |
| `evidence_manifest.json` | Source identities, original-to-saved filename mappings, hashes and available page references. |
| `evidence_index.docx` | A readable source index when generated, with links and filename mappings. |
| `result.review-draft.json` | Unpublished in-app review decisions, created beside the accepted result when you edit it. |
| `complaint_pack/` and `Complaint_pack.zip` | Optional focused complaint outputs created through reviewed-pack generation. |

The focused pack contains `Complaint_Draft.html`, `.docx` and `.pdf`, a
reference-only `Claim_Schedule.xlsx`, its manifest and verified relevant originals.
Make corrections in the **main** workbook or app, then regenerate; the focused
schedule cannot be imported as a review workbook.

## Run status

The desktop completion message and `run.json` describe the operation, rather than
certify the underlying bill or legal position. Inspect the notes as well as the
status:

| Status | Meaning and next action |
| --- | --- |
| `complete` | Requested artifacts were produced without the coordinator recording a review condition. Still verify evidence before relying on it. |
| `review_required` | Outputs were produced, but logged extraction errors, missing information or warnings need attention. Review the named records/sources/claims. |
| `empty` | No billing observations or SAP rows were available to render. Saved analysis/status can exist without a workbook or report. Check source selection, filters and parse notes. |
| `failed` | A stage or cleanup failed. Keep the accepted result and earlier successful artifacts; inspect `error` and the last stage before retrying. |
| `cancelled` | Cancellation was observed between processing steps. Some completed artifacts can remain. |
| `interrupted` | When loaded, a manifest left in a nonterminal state is interpreted as interrupted, for example after a forced process exit. Preserve the folder and rerun into a new destination. |

The coordinator can have saved the analysis before a later report fails. Conversely,
a failure during collection may leave no accepted result. The manifest records
actual artifact paths, sizes and SHA-256 hashes. Absolute paths describe the
location at generation time; a moved pack's manifest is not a relocation tool.

## Choosing what to reopen

| Goal | Input to use |
| --- | --- |
| Different report format/sections, same accepted figures | `result.json` through Generate reports or a report CLI command. |
| Correct values or reinstate/exclude records | `result.json` in Review evidence, or the main amended review workbook. |
| Change a complaint proposal or credit allocation | The app's Dispute ledger/Reconcile claims tools, or the main workbook's claim input sheets. |
| Apply new extraction/parser behavior | Original PDFs/messages/archives in a new extraction run. |
| Recover the most complete state from an old run | Its `result.json` and sources, if available; compatibility records/workbooks can contain less provenance. |

A draft is not an accepted result, a focused claim schedule is not a review
workbook, and a hand-edited report is not a replacement for either. Keep these
roles distinct when naming or archiving copies.

## Finding your way around Excel

The **Workbook guide** explains the review workflow within the file. The
**Start here** directory links to the sheets present in that particular run.
Some sheets require suitable data, enabled options or a newer accepted analysis;
an absent sheet does not by itself establish that no issue exists.

| Task | Sheets to use |
| --- | --- |
| Get an overview | Start here, Ledger overview, Annual Summary, Data Quality Report. |
| Inspect extracted records | EDF Evidence Report; Duplicate Entries, Filtered (Below Min) and Parse Errors when present. |
| Correct a value or exclude a record | Review & Corrections; Review history shows accepted changes. |
| Follow bills and credits | Bill revisions, Rebilling ledger, Rebilling components, Rebilling comparisons. |
| Record and consolidate a challenge | Dispute schedule, Claim groups, Credit allocations, Claim overlap review; Reconciled claims shows the accepted result. |
| Examine balances and payments | Key Statistics, Balance Trend, Year-on-Year, Period Charges, Payment Analysis. |
| Investigate possible problems | Dispute Flags, Dispute Timeline, Back-billing Analysis, Superseded Reconciliation, Rebilling & Corrections. |
| Check meter/tariff evidence | Meter Readings, Contract History, Tariff Analysis. |
| Examine statistical scenarios | Statistical Analysis, Forecast & Projection, Compensation when available. |
| Compare SAP exports with bills | SAP source/event sheets, Reconciliation and Reconciliation Drill-down. |
| Check how the run was produced | Provenance and Workbook guide. |

The balance/trend overview sheets need at least two selected observations from
one ledger. Several other analysis sheets remain present for small datasets and
explain insufficient evidence. SAP source sheets require their corresponding
parsed data and `scan_sap_dumps`; cross-source reconciliation also depends on
`generate_reconciliation_sheet`. Compensation appears when there are scenario
rows. Older saved results can lack newer revision/claim views until you create a
reviewed result.

Pale green cells in the designated review tables are inputs. Other sheets are
reference outputs. Excel formulas and charts do not replace regeneration through
the app. See [Excel corrections](USER_GUIDE.md#correct-records-in-excel).

## Understanding the figures

- **Balance:** what a source says the account owed or held in credit at a date.
  Do not sum successive balances as transactions.
- **Transaction:** a movement in the account. Charges increase debt; payments
  and credits decrease it. Refunds of account credit increase it toward zero.
- **Period charge:** the charge for a consumption period. A cancellation/rebill
  can have a net credit while containing positive replacement charges.
- **Gross electricity charges and components:** explain what a bill contains;
  they are not additional transactions to add again. Printed totals may include
  VAT while energy components exclude it.
- **Proposed reduction:** a human-reviewed claim, separate from recorded financial
  totals. Successive bills may repeat the same demand. Consolidate overlaps and
  credit allocations before relying on a combined proposal.
- **Missing value:** unknown, not zero. Account, fuel, meter and register identity
  limit which records can be combined.

A statement's document date or a MyAccount posting date does not prove the date
an underlying charge was first demanded. Register readings are distinct from
consumption in kWh. A mixed reading basis for an entire bill is also distinct
from the classification of its closing reading.

Back-billing analysis flags consumption older than a nominal 365-day window and
shows its calculation basis. It does not establish eligibility or a confirmed
write-off. The dispute schedule's calendar-year reference and reviewed claims
are described in the [rebilling methodology](REBILLING_LEDGER.md#how-comparisons-work).

Price-cap comparisons use a bundled GB-average electricity/direct-debit reference,
not a finding about the cap applicable to a particular account. Credit interest
requires a configured rate. Forecasts need at least three consecutive month-end
balances in one ledger; displayed bands are advisory scenarios.

## Reports

Use **Report settings…** to choose formats and sections. HTML, PDF and Word share
the same section registry and accepted analysis. Section numbers follow the
selection, so a section's number can change when you omit earlier sections.

The available sections cover the executive summary, key findings, evidence index,
detailed findings, timeline, Ofgem comparison, statistics, payments, forecast,
data quality, tariffs and potential billing adjustments. Appendices cover methodology, glossary
and the full evidence table. Detailed findings includes stored revision and
claim information when available. The workbook provides the fuller tabular
reference where report views limit automatic suggestions.

Editing a Word report changes that document only. To correct the underlying
figures consistently across outputs, use the app or main workbook and regenerate.

## Source files and OCR

The manifest maps original filenames to the actual saved copies. A single-invoice
PDF may be saved under its invoice number. A merged PDF is copied whole once;
several records can reference different original pages in it. Filename collisions
are kept distinct. Missing, ambiguous or changed sources produce review notes
instead of guessed links.

PST/OST message bodies may be saved as labelled generated text exports. These
are not original EML messages. Their archive occurrence remains part of provenance.

PDF pages without useful text can be recognised locally with English Tesseract
OCR. Normal text pages use their existing text. The manifest retains recognised
page text and failures, including pages yielding no billing rows. OCR records
carry page/method/confidence information and require review against the original.
OCR preserves the original PDF and cannot recover an attachment that an archive
reader failed to extract. Substantial but incorrect existing text layers may
still require manual correction.

## Saved results and older packs

Version 3 `result.json` preserves original observations, prepared records,
configuration, valuation date and calculated models. Loading it for another
report reuses those calculations. Re-extract sources to apply parser improvements;
create reviewed outputs to apply corrections or calculate newer review models.

Legacy JSON and older workbooks retain only the information they originally
stored. Import records provenance limits and may recalculate with a newly captured
context when historical details are absent. Current review workbooks preserve a
full baseline and support audited corrections; they are not the same as those
older, lossy imports.

Generating another report from a saved result can retain links to originals in
the earlier pack rather than copying them into the new report folder. Keep that
source pack available too. Exclusion is not redaction, and checksums identify
file contents rather than certify their truth.

## Before sharing a pack

1. Check the **run status**, evidence gaps and selected source pages. Resolve
   extraction mistakes through a reviewed pack so reports and workbook agree.
2. For a complaint total, inspect **Reconciled claims**, actual relief allocations
   and remaining unresolved positions. A reviewed subtotal may cover only part
   of the dispute.
3. Extract a copy of `Complaint_pack.zip` into a new location and open its draft,
   index and a few source links. Check the invoice and original page referenced,
   rather than only whether a file with a similar name opens.
4. Inspect whole shared PDFs and generated text/JSON for material you do not intend
   to disclose. Exclusion does not remove pages or original observation history.

Keep an intact private master pack. A separately redacted sharing copy must be
labelled as such; editing copied originals can invalidate recorded hashes. Do not
rewrite the master manifest to make changed bytes appear to be the preserved source.

## Linked accounts, readings and automatic adjustments

Newly analysed packs can include **Ledger overview**, a protected table of dated
balances and movements by analytical ledger. Explicit paired legacy/Kraken account
references in SAP evidence can link those account renderings. Missing fuel, meter
or register context is grouped only when unique within the supplied evidence;
conflicting and unassigned identities remain separate. Original scalar identities
remain in the evidence, and **Ledger Identity Basis** explains derived grouping.
The overview does not calculate a current balance by adding incomplete transactions.

A payment reminder repeats a demand; it is not another charge. General advice
about estimated bills or smart meters does not establish the reading used for a
specific bill. Repeated invoice payment-history rows use their parent invoice and
line identity. Statement/history links also require a corroborated sequence over
multiple dates and unique occurrences within each source. Ambiguities remain
separate and appear in review priorities/run warnings rather than silently merging.

The **Meter Readings** legend includes `D` for a deemed/estimated source. A bare
SAP `A` code is not described as supplier-confirmed when its read-source description
contradicts that interpretation. Identical SAP observations can be displayed once
with occurrence counts; the SAP source sheet retains the original rows and codes.

The report section **Potential Billing Adjustments** uses the existing
`compensation` selection key. Its workbook tab remains **Compensation** for
compatibility. These scenarios are separate from **Reconciled claims**. A positive
net rebill can contain much larger replacement charges and cancellations; its
calculation basis explicitly identifies the gross amount, net adjustment and
available VAT/sub-period basis. No automatic entitlement or concession is implied.

The legacy **SAP ↔ EDF Matched Events** sheet now labels its rows as candidates.
Only a unique High-confidence candidate is assigned in an event summary, and even
that remains subject to checking originals. Low/Medium and ambiguous candidates
are not promoted into a definitive invoice relationship. Date distance measures
proximity of the supporting posting/clearing date to the billed period: an in-span
date has distance zero.
