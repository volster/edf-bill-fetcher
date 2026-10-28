# Rebilling ledger and dispute schedule

The ledger answers which documents changed a bill, what changed in the observed components, and which amounts you propose to challenge. It preserves recorded account figures separately from your proposed reduction.

## Review in the app

1. Open the run's `result.json` in the evidence review workspace.
2. Use **Bill revisions** to confirm or reject a suggested replacement or cancellation credit. You can add a relationship the suggestions missed. A reason is required; confirmation does not remove a charge or credit from the accounts.
3. Open **Dispute ledger**. The history groups bills and credits by their account, fuel, meter and register. It includes excluded documents with explicit labels. Earlier estimates remain visible through repeated corrections.
4. Select a document and use **Show this original source** to inspect its PDF and original page range. Check the full earlier history, rather than only the immediately preceding estimate.
5. Set the dispute position, challenged amount, proposed reduction, basis/scope, and evidence EDF must substantiate. **Use full positive bill adjustment** copies the positive period charge into the two amount fields. You can enter a different evidenced amount manually. A net credit does not cap a challenge to positive replacement components.
6. Enter a first-demand date only when you have verified when the additional charges were demanded. A meter-reading date, the invoice date, or a later statement repeating a balance does not automatically establish that date.
7. Save the position to the draft. Undo, redo, autosave and draft recovery also cover these positions. Create a reviewed pack to publish them into a new workbook and reports.

Changing a related record, its inclusion, or a revision relationship marks an existing position **NEEDS REVIEW**. Its old amount remains visible for traceability. Check it against the changed history and save the position again to acknowledge the current evidence. Record verification status alone does not acknowledge a dispute position.

**Clear position** removes the proposal from the current schedule; undo and the regenerated review history preserve the decision. Withdrawing a position is also available as an explicit status.

## Review in Excel

The workbook contains four supporting views:

| Worksheet | Purpose |
| --- | --- |
| Rebilling ledger | Chronological documents, original monetary meanings, reading dates, relationships and source links. Each credit document appears once, even if several bills reference it. |
| Rebilling components | Observed energy, standing-charge, VAT, discount and cancellation components. These are explanations of bills, not extra transactions to add to a balance. |
| Rebilling comparisons | Every stored component comparison, including unmatched periods and partial overlaps. Earlier and later sources link to their original pages. |
| Dispute schedule | Editable human positions alongside protected source facts and stable row IDs. |

In **Dispute schedule**, edit the pale green cells:

- **Status:** `unreviewed`, `challenged`, `not_challenged`, or `withdrawn`.
- **Challenged amount (£)** and **Proposed reduction (£):** non-negative amounts with at most two decimal places. A reduction requires a challenged amount and cannot exceed it. Unknown amounts can remain blank.
- **First demand date:** a real Excel date, `DD/MM/YYYY`, or `YYYY-MM-DD`; blank means not verified.
- **Basis / scope:** identify precisely which charges are challenged and the supporting evidence. Required for a reviewed position.
- **EDF must substantiate:** identify missing earlier demands, readings, allocations or credit/replacement records.

Keep all rows and grey reference cells. To sort, unprotect the worksheet without a password and sort the whole table. Formulas and Excel error values are rejected; paste values instead. To clear a saved position, choose `unreviewed` and blank the other five input fields. The removal is audited when you regenerate.

Save a copy beside `evidence_files` and select it in **Generate reports**, or run:

```console
python main.py --review-workbook --workbook amended.xlsx --output-dir output
```

The app validates the workbook before creating a new run folder. Reference sheets remain calculated outputs; editing them does not instruct the app to change the evidence. Excel does not recalculate this analysis itself.

An unchanged proposal retains its original evidence binding. If you also correct the underlying record in Excel, the proposal becomes stale. Changing a position's inputs acknowledges the current corrected evidence; you can also verify and re-save it in the app. Excluded positions remain explicitly labelled and are not active complaint claims.

## How comparisons work

- Automatic relationships require a known matching account and compatible fuel, meter and register. Period overlaps and explicit credit references produce suggestions. They do not confirm replacement, cancellation, allocation, or recoverability. Explicit relationships reject cycles and conflicting identities.
- All canonical bills and credits remain in the history, even if relationship suggestions reach their 500-suggestion limit. Every explicit relationship is retained.
- Energy components are compared only when there is one uniquely matching consumption period in each document. Repeated periods that might represent multiple registers remain unallocated. An earlier-only component, missing values or a partial overlap never receives a fabricated allocation.
- For a uniquely matched energy component, using units `U` and price `R` in pence/kWh:
  - units effect = `(U_later − U_earlier) × R_earlier / 100`;
  - rate effect = `U_later × (R_later − R_earlier) / 100`;
  - residual = observed charge difference minus the two displayed, penny-rounded effects.
- The order of that decomposition is explicit. It explains the observed charge difference; it does not establish which units were previously unbilled, when EDF received a reading, or an entitlement to relief.
- Standing-charge and VAT observations remain separate. Their observed differences are shown only for the same whole-document consumption period and when both documents contain the component. Absence is unknown, not zero. These extracted lists may be incomplete.
- Printed total-electricity amounts retain their printed basis: newer invoice layouts can include VAT in that total. Energy component amounts exclude VAT; a printed gross total must not silently be treated as ex VAT.
- Original extracted components are not rewritten by corrections to scalar fields such as the total units or rate. Both remain available for checking discrepancies against the source.
- The calendar date one year before a verified first demand is a reference, not a calculated protected amount. There is no automatic day-based percentage concession or write-off.

## Outputs and compatibility

The accepted result stores this analysis once. Excel, HTML, PDF and DOCX consume that stored model rather than recomputing relationships during rendering. The detailed findings section includes the full document history and every saved dispute position, including stale, excluded and withdrawn labels. Reports include all explicit component comparisons and up to 30 automatic suggestions; the workbook contains every stored comparison. HTML source links use saved source identities and original page numbers, with unavailable sources labelled explicitly.

**Schedule amounts are not summed.** Successive revised bills may demand the same charges. Credits already applied must not be claimed twice. The app does not convert proposals into a revised account balance or a confirmed legal entitlement. Reconcile overlap and credit allocation before asserting a single total to EDF.

Original observations and source files stay unchanged. A record exclusion does not redact pages from a shared PDF. If a correction merges duplicate canonical records, a position on an absorbed record is archived with an audit entry rather than transferred or added to the retained record's proposal.

Previously saved accepted results remain readable without silently gaining a newly calculated model. Open one in review and create a reviewed pack to generate the ledger and schedule. Older review drafts without schedule data remain readable. No private account evidence or complaint position belongs in repository fixtures; regression tests use synthetic examples.

## Consolidate claims and account for credits

In the app, choose **Dispute ledger → Reconcile claims…**. The separate editor has tabs for claims/credits, overlap review and complaint wording. It shares the review draft, undo/redo, autosave and reviewed-pack workflow.

1. Create a short claim ID and title. Assign challenged bills representing the **same economic charges**. Supporting earlier bills remain in the history without becoming extra claims.
2. Enter one consolidated challenged amount and one requested reduction **before relief already received**, including applicable VAT. These are human proposals. The app never adds successive invoice totals or caps a challenge at a net rebill increase. Blank means unknown; zero is an explicit value.
3. Allocate recorded credit sources with a reason for each allocation:
   - `already_applied_relief` is relief received for these charges and reduces the further adjustment requested;
   - `cancellation_offset` explains cancellation of an earlier bill and is disclosed without a second deduction from the proposal.
4. Mark the group reviewed only after checking its scope and the full earlier history. A source credit cannot be allocated beyond its recorded negative transaction, including allocations across different reviewed groups. Payments are not credit-allocation sources.
5. Resolve potential overlaps. Repeated demands should normally share one group. Use **Confirm separate charges** only with evidence that the two proposals cover distinct charges. A document in two groups cannot be waived as independent. Decisions are bound to both groups and their evidence.
6. Check **Reconciled claims**. A combined proposal is withheld while active challenges are ungrouped, overlapping, stale or otherwise unresolved. A reviewed subtotal is labelled separately. Conditional balances use a dated, unambiguous observation within the same ledger; they do not establish current liability or concede remaining charges.

Known conflicting fuels, meters or registers cannot be consolidated. Linking different account references requires a shared known meter/fuel and an explicit source-backed account-link reason. Unknown account identities require review. Evidence changes make unchanged group and overlap decisions stale. Removing a group, changing its figures, credit allocations and complaint wording are audited when a reviewed pack is created. Duplicate absorption archives affected groups instead of transferring or adding their proposals.

### Excel equivalents

New workbooks include the same editable decisions:

| Worksheet | Inputs |
| --- | --- |
| Dispute schedule | **Claim groups**: assign the short group ID after creating it below. Multiple temporary memberships use semicolons and prevent double counting until resolved. |
| Claim groups | One consolidated amount, reduction before received relief, status, title, scope/evidence and any account-link evidence per ID. |
| Credit allocations | Claim ID, source reference, amount, treatment and reason. Choose `CRED…` references from the protected source catalogue on the right; they bind to original row IDs, including after corrections in the same import. |
| Claim overlap review | Two group IDs, `separate_charges` and a supporting reason, or `unreviewed`. |
| Reconciled claims | Protected output showing the accepted proposal, credit treatment, issues and scope. Regenerate to update it. |

Each input table supports up to 1,000 rows. Larger or malformed inputs are rejected before creating a run folder; rows are never silently dropped. Paste values rather than formulas, retain the catalogue/reference cells, and sort complete input tables when necessary. Unchanged inputs retain their earlier evidence binding. Correcting a source without changing a decision makes that decision stale. To acknowledge the changed evidence, verify and re-save in the app, or update the Excel decision's supporting reason after checking it.

Older accepted results, review drafts and workbooks remain supported. Their existing schedule headers are retained until a new reviewed result contains claim reconciliation. Opening an old result does not silently calculate new proposals.

## Focused complaint draft

Creating reviewed outputs with active challenges also creates `complaint_pack/` and `Complaint_pack.zip`. Use the editor's **Complaint draft** tab to enter your summary/outcome or disable this additional pack. Nothing is sent automatically.

The pack contains `Complaint_Draft.html`, `.docx`, `.pdf`, a reference-only `Claim_Schedule.xlsx`, a source index and portable `evidence_manifest.json`, and verified relevant original files. The letter leads with the requested outcome and a complete reconciled proposal when available. Otherwise it withholds one total and identifies the remaining work. It asks EDF to substantiate first demands, additional charges, readings and their receipt, and the cancellation/credit/replacement trail, without inventing eligibility, payment conduct or a percentage concession.

Curation retains linked and overlapping compatible earlier history, including higher estimates. Excluded-only sources and withdrawn claim details are omitted. Shared source PDFs remain **whole** and may contain other pages; this is not redaction. The index preserves source identities, original page ranges, saved relative filenames and SHA256 hashes. Hashes establish which bytes were included, not the truth of their contents.

Missing/changed originals, unverified page ranges or first demands, and unresolved reconciliation appear as review gaps and make the run require review. A missing saved copy is never replaced by a same-named file or an unchecked original. A recorded original can be used only when no saved copy was promised and its recorded hash verifies. Source absence does not discard the preserved extraction or invent a replacement.

The focused schedule is an outward reference, **not** a review-import workbook. Make corrections in the main workbook or app and regenerate all artifacts together. The pack is staged and published with its ZIP only after successful generation; cancellation or failure removes temporary files while preserving the main run's accepted analysis. Existing complaint packs are never overwritten. Run manifests include hashes for the focused artifacts.
