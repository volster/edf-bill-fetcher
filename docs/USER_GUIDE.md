# Using and reviewing evidence

[Documentation index](README.md) · [Output reference](OUTPUTS.md) · [Troubleshooting](TROUBLESHOOTING.md)

The usual workflow is **collect → check → correct → regenerate**. The app keeps
original observations and review decisions so you can explain how a later pack
differs from the first one.

## Build your first pack

Open **Build evidence pack** and select one or more sources:

| Source control | Choose |
| --- | --- |
| PDF bills | A folder of EDF bills, credit notes, statements or supported SAP dump PDFs. Merged PDFs can contain several invoices. |
| Account history | An EDF MyAccount HTM/HTML export. Printed PDF histories go in the PDF folder. |
| Outlook archive | A PST or OST file. |
| Email files | A folder containing EML or MSG messages and their attachments. |

Choose **Save to** and a **Pack name**. Leave **Include original source files**
selected for a portable evidence pack. Select **Include PDF, Word or HTML
reports** and use **Report settings…** to choose formats and sections. Excel
is included in every extraction pack.

Check **Preferences** if you need an account or sender filter, amount thresholds,
duplicate handling or SAP settings. Filters can explain apparently missing
records; start with the intended account and avoid unnecessary exclusions.

Click **Create evidence pack**. Readiness checks identify missing dependencies
before extraction. **Cancel run** requests cancellation; the current processing
step may need to finish first. On completion, inspect **Details** for review
notes and use **Open folder** to find the output. Every desktop run uses a new
folder. A failed run can still contain successfully saved analysis or artifacts;
its status records the failure.

## Check the result

Open the workbook's **Start here** sheet for the overview, review priorities
and links to other sheets. In the app, open the run's `result.json` in
**Review evidence**. Search records and compare their fields with the original
PDF pages. **Source files** includes sources such as SAP dumps that may not have
ordinary invoice rows.

Check invoice references, dates, account/meter identity, charges and reading
classifications before relying on a finding. OCR-derived values need particular
attention to digits, signs and decimal points. An account balance is a dated
observation, not another charge to add to the bill total. See
[financial meanings](OUTPUTS.md#understanding-the-figures).

If you need only a different report format, choose `result.json` in **Generate
reports** and select the desired formats/sections. This reuses the saved
calculations. To benefit from an updated parser, re-extract the original sources;
opening an earlier result does not re-read its PDFs.

## Correct records in the app

1. In **Review evidence**, select a canonical record and choose **Edit selected
   record**. Compare it with the adjacent source.
2. Correct fields or clear **Include** to exclude the record. Enter a reason,
   then choose **Apply to draft**. You can mark it **Verified against source**
   or **Needs information**.
3. Use **Preview changes** to inspect recalculated figures without creating files.
4. Choose **Create reviewed pack** to publish a fresh workbook, saved analysis
   and reports using your report settings. The original pack remains available.

Source text, file hashes and occurrence metadata cannot be edited. Draft changes
are automatically saved beside the accepted analysis as
`result.review-draft.json`; **Save draft** saves immediately. Keep it beside its
original `result.json`. Reopening resumes the draft, but the session's undo/redo
stack is not restored after restart. A draft bound to a different or changed
analysis is rejected.

Unapplied forms prompt before navigation or closing. Generation failure or
cancellation leaves the original analysis and draft available. If you correct
sources underlying an existing dispute position or claim, revisit any decisions
marked **NEEDS REVIEW** before relying on their amounts.

## Correct records in Excel

Use the **main evidence workbook**, not the focused complaint pack's
`Claim_Schedule.xlsx`.

1. Save a copy of the workbook beside the pack's `evidence_files/` folder.
2. Open **Review & Corrections**. Edit the pale green inputs, give a reason for
   each change and keep all rows and reference cells. Set **Include = No** to
   exclude a record; do not delete the row.
3. Save and close the workbook. In the app's **Generate reports** page, select
   the amended `.xlsx`, choose the output location and report settings, then
   generate the new pack.
4. Open its **Start here** and **Review history** sheets to check the recalculated
   result and the recorded before/after changes.

Use real Excel dates, `DD/MM/YYYY` or `YYYY-MM-DD`. Paste values rather than
formulas and keep identifiers formatted as Text to preserve leading zeroes.
Blank amounts mean unknown. Payments and credits are negative transactions;
refunds of account credit are positive. Correcting a period charge also updates
a transaction that previously matched it, unless you edit the transaction
explicitly.

Grey reference fields and hidden row IDs bind edits to the original evidence.
To sort, use **Review → Unprotect Sheet** (no password) and sort the **whole
table**. Expand grouped columns using Excel's outline controls or Unhide Columns;
unprotect the sheet if Excel disables outlines. Keep hidden review metadata and
reference sheets intact.

Excel does not rerun the app's analysis. Summary and reference tabs show the last
generated result until you import the amended workbook. Editing those tabs is
not a supported correction input. Invalid dates, formulas, missing/duplicate IDs,
changed protected references and unexplained corrections can block import.
Adding manual records and editing SAP dump rows are outside this workflow.

The new pack retains the original valuation date, settings, observations and
source identities. Previously excluded records can be reinstated. Their source
files may need the earlier pack, so keep it available. Changes to previously
verified records reset their verification state; unchanged decisions carry forward.
Older workbooks without review metadata use a limited import path. If you have
their `result.json`, generate a fresh workbook from it first.

For automation, see [amended workbook commands](CLI.md#import-an-amended-workbook).

## Review revisions and prepare a complaint

**Bill revisions** lets you inspect and confirm or reject suggested replacement
and cancellation relationships with a reason. A confirmed link does not remove
a charge or credit from the ledger: an existing cancellation credit may already
offset the original bill.

Use **Dispute ledger** to record which charges you challenge and why. Then use
**Reconcile claims…** to consolidate repeated demands, account for credits and
resolve overlaps before asserting one reduction. The same claim decisions can
be edited in the main workbook. The [rebilling and claim guide](REBILLING_LEDGER.md)
explains the fields, evidence requirements and focused complaint pack.

## Keep the evidence together

Move or archive the **whole run folder**, not just the workbook. Links depend on
the source copies and manifest; source PDFs are not embedded in Excel. Preserve
earlier packs for excluded records and review history.

Excluding a record does not redact a PDF. Shared PDFs remain whole, and saved
analysis retains original observations and review history. Inspect the files you
intend to share; prepare redacted copies separately if needed. Workbook protection
prevents accidental edits but is not encryption.
