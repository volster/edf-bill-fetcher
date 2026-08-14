# Superseded Reconciliation Page + Back-billing Analysis Cleanup

**Date:** 2026-08-14
**Status:** Approved design (brainstormed 2026-08-14)

## Problem

The `Back-billing Analysis` worksheet is hard to read and produces totals that
appear to double-count:

1. **Superseded vs live rows are visually indistinguishable.** A superseded
   invoice (killed by a later cancel-and-rebill) renders with the same font,
   fill and bold unlawful amounts as a live invoice. In the reference run the
   sheet shows 4 superseded rows whose unlawful charges sum to £9,187.77 and
   look identical to the 3 live rows.
2. **Three inconsistent totals.** Summing the visible `Unlawful Charge (£)`
   column gives £41,312.92; the trailing "SURVIVING INVOICES" row sums the 3
   live rows' own unlawful to £32,125.15; the "UNION OF CONSUMPTION DAYS" row
   walks every unlawful day once at the earliest-billed rate and gives
   £30,567.32. A reader cannot tell which figure is "the answer", and the
   surviving total genuinely double-counts the consumption the superseded
   rows already claimed.
3. **No PDF hyperlinks.** The "Open PDF" columns are a misnomer — every
   hyperlink cell only jumps *within* the workbook to the Evidence Report
   sheet; there are zero external/file links anywhere in the workbook.
4. **Garbage evidence filenames.** `evidence_files/` copies keep the raw
   source filename (`671078701920_060241004086_20190416.pdf`) rather than the
   invoice number, so filenames do not identify the bill.

## Goals

- A dedicated **Superseded Reconciliation** page that moves the superseded
  rows off Back-billing Analysis and provides full before→after navigation
  (killer on the spreadsheet, original invoice on the spreadsheet, and real
  PDF links for both).
- A cleaned Back-billing Analysis sheet with only live rows and **one**
  defensible no-double-count total.
- Invoice-number evidence filenames and real PDF hyperlinks throughout (the
  new page depends on both).

## Non-goals

- No change to the union *algorithm* (`compute_unlawful_union_total`) — it is
  correct; we change which rows feed it on the live sheet.
- No change to the underlying detection/domination logic.
- No new data collection; everything derives from existing records.

## Reference numbers (2026-08-14 run)

| Figure | Value |
|---|---|
| Sum of all rows' Unlawful (visible column) | £41,312.92 |
| Sum of superseded rows' Unlawful | £9,187.77 |
| Sum of live rows' Unlawful | £32,125.15 |
| Union over all 7 events | £30,567.32 |
| Union over live rows only | £32,125.15 (= sum of live unlawful) |

Key finding: once superseded rows are removed from the set, **union-over-live
= sum-of-live-unlawful** (£32,125.15). The earlier gap (£1,557.83) existed
only because superseded rows (billed earlier) claimed the shared days at their
rates first. With the 3 live rows not overlapping each other, the union and the
simple sum coincide — but the union is the defensible figure to keep because it
stays correct if live rows ever overlap.

## Design

### 1. Back-billing Analysis sheet (cleaned)

- Renders **only `Status="Live"`** rows.
- The trailing total is a single row:

  ```
  TOTAL UNLAWFUL CHARGES — UNION (no double count)   £32,125.15
  ```

  computed as `compute_unlawful_union_total(bb[bb live])`. No second
  "SURVIVING" row (they are identical in current data; a single row avoids the
  three-numbers confusion).
- Each live row additionally shows a "View superseded" link (small cell or
  suffix) that jumps to the Superseded Reconciliation page filtered to that
  survivor's chain. Implementation: the reconciliation sheet groups rows by
  survivor (a `Killer Invoice #` header group, or a per-survivor section
  label), and the link targets the first row of that survivor's group. If no
  superseded chain exists the cell is blank.
- Superseded rows are **not** written here.

### 2. Superseded Reconciliation sheet (new)

- One row per superseded invoice (a key in `domination_map`), sorted by Bill
  Date, **grouped by killer**: a section label row per survivor (e.g.
  `KILLER: T78701920068`) followed by that survivor's superseded rows, so a
  reader navigating from a live row's "View superseded" link lands on the
  right group. The group label row is the hyperlink target.
- Columns (left→right):

  | # | Column | Notes |
  |---|---|---|
  | 1 | Invoice # | superseded invoice |
  | 2 | Bill Date | |
  | 3 | Period From | |
  | 4 | Period To | |
  | 5 | Days Billed | |
  | 6 | Period Charge (£) | |
  | 7 | Unlawful Charge (£) | superseded invoice's own unlawful |
  | 8 | Excess Days | |
  | 9 | Cancel/Rebill Disclosed | `_disclosed_label` |
  | 10 | Reason Assessment | + chain note (existing supersession text) |
  | 11 | Killer on spreadsheet | hyperlink → survivor's row on Back-billing Analysis |
  | 12 | Original invoice on spreadsheet | hyperlink → this invoice's row on EDF Evidence Report |
  | 13 | Original invoice PDF | `file://` link → `evidence_files/<inv>.pdf` |
  | 14 | Killer invoice PDF | `file://` link → `evidence_files/<killer>.pdf` |
  | 15 | Partial Overlap | `Yes`/blank from domination map |

- Trailing total row:

  ```
  TOTAL SUPERSEDED UNLAWFUL CHARGES (absorbed into survivors)   £9,187.77
  ```

  labelled explicitly as absorbed/audit — never added to the live total.
- Sheet name: `Superseded Reconciliation`.

### 3. Invoice-number filenames

`save_evidence_files` in `io/writers/evidence_bundle.py`:

- Build `invoice → attachment_name` from `evidence_df` (`Invoice #` ↔
  `Attachment Name`).
- Destination name = sanitised `<Invoice #><ext>` (e.g. `T78701920034.pdf`,
  `KI-31105244-0014.pdf`). Sanitise: strip path separators, leading/trailing
  whitespace, and characters illegal on Windows (`<>:"/\|?*`).
- Collisions append `-2`, `-3`, … (existing pattern).
- Rows whose `Invoice #` is `N/A`/empty, or where multiple distinct
  attachments map to one invoice, fall back to the raw attachment name.
- Return value stays `{attachment_name: destination_path}` so the bundle index
  keeps working.

### 4. Real PDF hyperlinks

- New helper `pdf_hyperlink_cell(ws, row, col, relative_path, tooltip)` in
  `helpers/excel_utils.py` that writes a cell whose hyperlink `target` is the
  relative path `evidence_files/<inv>.pdf` (openpyxl emits a `file://` link for
  relative targets when opened in Excel). Display: invoice number (or a short
  label like the filename) in blue underline.
- Evidence Report: the `Attachment Name` cell becomes a PDF hyperlink to
  `evidence_files/<Invoice #>.pdf` (display = invoice number), so the sheet
  itself links to the saved file.
- Back-billing Analysis and Superseded Reconciliation use the same helper for
  their PDF cells.
- The existing `open_pdf_hyperlink_cell` (in-workbook jump) is **renamed** to
  reflect its real purpose (jump to Evidence Report) and is reused for the
  "…on spreadsheet" link cells; it is not used where a PDF link is intended.

### 5. Data flow

- `domination_map: dict[superseded → (survivor, partial_overlap)]` (already
  computed by `compute_transitive_domination` in the export pipeline) drives
  the split and the reconciliation rows.
- `evidence_df` provides `Invoice # → Attachment Name`.
- `source_paths` (engine) provides `Attachment Name → absolute path`.
- Combined mapping `Invoice # → saved PDF path` is built after
  `save_evidence_files` returns (it returns `{attachment_name → dest}`);
  the reconciliation writer takes this mapping to emit the PDF links.
- New writer `write_superseded_reconciliation_sheet(ws, bb, domination_map,
  evidence_index, pdf_paths, ...)`; wired into the export pipeline next to
  `write_back_billing_sheet`. Sheet order: place immediately after
  `Back-billing Analysis`.

### 6. Error handling

- Missing `source_paths` / missing PDF for an invoice → log and render the PDF
  link cell as plain text (italic grey "no file"), never raise.
- `Invoice #` with no attachment → fallback raw name (Section 3) or plain-text
  cell.
- Empty `domination_map` → reconciliation sheet still written with zero rows
  (header + total £0.00), matching the existing "no domination map → all live"
  behaviour.

### 7. Testing

- `test_back_billing_sheet.py`:
  - superseded rows no longer rendered; only live rows + single union total;
  - "View superseded" link present on live rows that have a chain;
  - total equals `compute_unlawful_union_total(live subset)`.
- New `test_superseded_reconciliation_sheet.py`:
  - row per superseded invoice with all data columns;
  - the 4 link cells resolve to the expected targets (killer sheet row, evidence
    row, and the two `file://` PDF paths);
  - trailing absorbed-total row = sum of superseded unlawful;
  - empty domination map → header + £0.00.
- `test_evidence_bundle.py`: invoice-number naming, sanitisation, `-2` dedupe,
  `N/A` fallback.
- `test_excel_utils.py`: `pdf_hyperlink_cell` emits `target` = relative path.
- Full suite + `mypy .` + `ruff` gates must stay green.
