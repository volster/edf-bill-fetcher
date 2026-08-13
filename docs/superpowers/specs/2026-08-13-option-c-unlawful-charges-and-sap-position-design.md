# Option C Unlawful Charges, Union Total, and "Backbilling According to SAP" Tab

Date: 2026-08-13
Status: Approved (design) — awaiting spec review

## 1. Context

The Back-billing Analysis tab currently computes each back-billing invoice's
"Unlawful Charge (£)" by prorating the invoice's *net* "Period Charge (£)"
("Your charges for this period (including VAT)") by its excess-days ratio.
Comparison against Claude's spreadsheet (`scratch/claude spreadsheet/EDF_Billing_Statement_Updated.xlsx`)
showed this massively understates the position:

| Invoice | Our unlawful | Claude unlawful | Ratio |
|---|---|---|---|
| T34 | £3,781.49 | £16,984.15 | 4.5× |
| T67 | £2,010.39 | £3,849.00 | 1.9× |
| T68 | £990.38 | £15,051.58 | 15.2× |
| **Total** | **£6,677.15** | **£35,884.73** | 5.4× |

Root cause: the invoices are cancel-and-rebill invoices. The actual
back-billed window is disclosed on the PDF as "We cancelled your electricity
charges (excluding VAT) £X from <pf> - <pt>", with a per-sub-period
"About your charges" table (period, readings, units kWh, rate, charge). Our
detector prorates the *net* charge instead of reconstructing the window from
its sub-periods.

The design chosen by the user:
- **Option C base**: unlawful charge per invoice = sum over per-sub-period
  slices of `kWh × rate` for the consumption supplied more than 365 days
  before the bill date.
- **Union total**: the trailing total must NOT double-count overlapping
  consumption (e.g. T67's unlawful window sits inside T68's). Each
  consumption day is counted once, at the rate EDF charged when it FIRST
  recovered that day.
- **"Mine everything"**: extract ALL per-sub-period rows from PDFs that
  disclose multiple periods.
- **New tab "Backbilling According to SAP"**: cross-referenced view —
  reconciliation statement (period-level charges/reversals) + SAP financial
  transactions (ledger postings), reconciled against our PDF-derived
  Back-billing Analysis.

## 2. Verified data facts

- All 6 T-series invoice PDFs carry a per-sub-period "About your charges"
  table. A single robust regex (period pair + variable-length reading tokens
  + `units kWh ratep £charge`) captures all **23 rows**:
  T33: 2, T34: 5, T56: 4, T65: 3, T67: 3, T68: 6.
- The sub-periods **partition the full billed window** (T34 includes a 1-day
  row `04 Sep 19 - 04 Sep 19, 113 kWh`; T68 spans 6 rows covering
  02 Oct 2020 → 09 Aug 2023).
- **KI-31105244-0014 has no invoice PDF** in the ombudsman download — only the
  reconciliation statement `A-31105244-28261421` exists for the KI era. It
  therefore has no sub-period table → day-ratio fallback with an explicit
  "sub-period data unavailable" note.
- All 7 back-billing invoices have SAP financial-transaction cluster matches,
  so the SAP cross-check can cover all of them.
- SAP data already on disk / in workbook:
  - SAP Financial Transactions (908 rows): `Dr- Consum Billing Receivable`
    (280) and `Cr- Credit for Consum Billing` (18) postings, clustered by
    Clearing Document.
  - SAP Back-billing Events (631 rows): clearing-doc clusters — "a big page
    of charges" that doesn't state the position.
  - Reconciliation statement (ref 28261421, 43 rows in evidence report):
    24 `Charge` rows with explicit periods + amounts, 14 `Credit`
    (reversed) rows with periods embedded in Details, e.g.
    `Reversed electricity charge (14 May 2024 - 30 Sept. 2024) £-1596.70`.

## 3. Design

### 3.1 Sub-period extraction ("mine everything")

- New regex `SUB_PERIOD_RE` in `edf_bill_fetcher/processors/patterns.py`:
  `(?P<pf>\d{1,2}\s+\w{3}\s+\d{2,4})\s+-\s+(?P<pt>...) <readings…> (?P<units>[\d,]+)\s*kWh\s+(?P<rate>[\d.]+)p\s+£(?P<charge>[\d,]+\.\d{2})`
  with a non-greedy variable-length middle capturing the reading tokens.
- `BillingRecord` (models/records.py) gains `sub_periods: list[dict]` where
  each dict = `{"period_from", "period_to", "units_kwh", "rate_p", "charge"}`.
  Default `[]`.
- `_process_new_invoice` (collectors/engine.py) calls a new
  `extract_sub_periods(text)` helper in `processors/patterns.py`
  (returning a list of dicts; lives beside `SUB_PERIOD_RE` which it uses)
  and stores the list on the record.
- `BillingRecord.to_dict()` emits `Sub Periods` as a compact serialized
  string (e.g. pipe-joined `pf|pt|units|rate|charge`) so it survives the
  record→dataframe→writer pipeline without a schema change to the evidence
  frame. A dedicated column `Sub Periods` is added to the evidence report.
- `detect_back_billing` (processors/detection.py) parses `Sub Periods` back
  into structured rows; helper `_parse_sub_periods(raw) -> list[dict]`.

### 3.2 Unlawful charge per invoice (Option C)

In `detect_back_billing`, replace the single whole-period proration:

- `cutoff = bill_date - 365 days` (same legal gate as today).
- For each sub-period `[pf, pt]` with `(units, rate)`:
  - `unlawful_days = days in [pf, min(pt, cutoff)]`
  - if `unlawful_days <= 0`: skip (fully lawful)
  - if `unlawful_days >= sub-period days`: fully unlawful → add
    `rate/100 * units`
  - else straddle: add `rate/100 * units * unlawful_days / total_days`
- Sum → `Unlawful Charge (£)`.
- New column `Sub-Period Basis`: `"Sub-period × rate"` when sub-periods were
  used, `"Day-ratio fallback"` (existing whole-period proration) when the
  invoice has no sub-period rows (KI-0014) — with the Reason Assessment
  noting "sub-period data unavailable".
- `Period Charge (£)`, `Excess Days`, `Value Source` columns unchanged.

Sanity check (T68, bill 09/08/2023, cutoff 09/08/2022):
fully-unlawful 02 Oct 20→12 May 22 sub-periods (£11,528.88) + straddling
13 May 22→31 Mar 23 slice (£15,951 × 88/322 = £4,359.63) ≈ **£15,888**,
vs today's £990.38 and Claude's £15,051.58. Residual delta is the day-ratio
split on the straddle row vs Claude's meter-read split — acceptable; noted in
the tab.

### 3.3 Union total (no double count)

New `compute_unlawful_union_total(bb: pd.DataFrame, ...) -> float` in
`processors/detection.py`:

- Iterate back-billing rows in **bill-date order** (already the detector's
  sort order).
- For each row, for each unlawful sub-period slice, emit
  `(day_range, rate_p, kwh_per_day = units/total_days, bill_date)`.
- Build a day→(rate, kwh_per_day) map; a day is **claimed once** by the
  earliest-bill-date invoice that recovers it (the rate EDF first charged).
- Total = Σ over claimed days of `rate/100 * kwh_per_day`.

The Back-billing Analysis writer's trailing block gains a second row:
`TOTAL UNLAWFUL CHARGES — UNION OF CONSUMPTION DAYS (no double count)`.
The existing `TOTAL RETROSPECTIVE CHARGES — SURVIVING INVOICES` row
(period-charge sum of live rows) is retained unchanged alongside it — the
union row is the new bottom-line for the unlawful position, the surviving
row remains the headline period-charge figure. The per-row `Unlawful
Charge (£)` column stays per-invoice (superseded rows remain visible,
hyperlinked, for the auditable chain — decision 1). A note explains the
union vs the naive per-row sum.

### 3.4 "Backbilling According to SAP" tab (cross-referenced)

New analyser `analyse_sap_back_billing(sap_financial, evidence_df)` (home:
`processors/matching.py` or `processors/sap_parsers.py`) + new writer
`write_sap_back_billing_position_sheet(ws, ...)` (home: `io/writers/sap.py`).

Inputs:
- `sap_financial` rows (already parsed by `parse_sap_financial_transactions`).
- Reconciliation-statement rows already in the evidence frame
  (Source = `Statement Reconciliation`).
- Our Back-billing Analysis output (the PDF-derived events + unlawful
  charges + union total).

Content:
1. **Events table** — SAP back-billing events, restricted to clearing-doc
   clusters containing a `Cr- Credit for Consum Billing` reversal
   (decision 2): Clearing Doc, Clearing Date, Reason, # rows, Net Amount,
   Period(s) affected (from reconciliation-statement rows or transaction
   text), Original / Cancelled / Re-billed amounts, matched EDF invoices.
2. **Reconciliation vs our Back-billing Analysis** — per event: our
   invoice(s) + unlawful charge + verdict
   (`Reconciled` / `Partial` / `SAP-only` / `Ours-only` / `Δ £`).
3. **Position summary** — SAP-side total vs our union total; note the
   reconciliation statement (ref 28261421) is the KI-era source while T-era
   events come from the financial clusters.

### 3.5 Error handling

- Sub-period rows with unparseable dates are skipped per-row (never crash the
  invoice); if ALL sub-period rows fail, fall back to day-ratio + note.
- Straddle day-ratio uses `max(1, total_days)` to avoid div-by-zero on
  1-day rows (T34's `04 Sep 19 - 04 Sep 19`).
- Union day map is bounded by construction (spans ≤ ~1042 days); no
  performance concern.
- SAP analyser tolerates missing reconciliation-statement rows (T-era): event
  still listed from financial clusters with period "—".

## 4. Testing

- `test_sub_period_extraction`: fixture PDF text → expected 23 rows across
  the 6 invoices, exact figures (incl. T34's 1-day row, T68's 6 rows).
- `test_option_c_unlawful`: T68 unlawful ≈ £15,888 (not the old £990.38);
  T34/T67/T56/T65/T33 recomputed; fallback path for KI-0014.
- `test_union_total`: overlapping T67∩T68 unlawful days counted once at
  first-recovery rate; superseded rows included in union but not double
  counted.
- `test_sap_bb_position_sheet`: fixture SAP financial + recon rows → events
  table (reversal-containing clusters only), reconciliation verdicts,
  position summary totals.
- Regenerate the workbook from `scratch/input`; compare new union total
  against Claude (£35,884.73) and confirm SAP tab cross-checks.
- Full suite + ruff + mypy gate.

## 5. Files

- `edf_bill_fetcher/processors/patterns.py` — `SUB_PERIOD_RE`,
  `extract_sub_periods`
- `edf_bill_fetcher/models/records.py` — `BillingRecord.sub_periods`,
  `to_dict`
- `edf_bill_fetcher/collectors/engine.py` — `_process_new_invoice` wiring
- `edf_bill_fetcher/processors/detection.py` — `detect_back_billing` Option C,
  `_parse_sub_periods`, `compute_unlawful_union_total`
- `edf_bill_fetcher/io/writers/back_billing.py` — trailing union total +
  `Sub-Period Basis` column + note
- `edf_bill_fetcher/processors/matching.py` (or `sap_parsers.py`) —
  `analyse_sap_back_billing`
- `edf_bill_fetcher/io/writers/sap.py` — `write_sap_back_billing_position_sheet`
- `edf_bill_fetcher/io/writers/export.py` — wiring new tab + union total;
  evidence-report `Sub Periods` column
- Tests: new `tests/test_sub_period_extraction.py`,
  `tests/test_option_c_unlawful.py`, `tests/test_union_total.py`,
  `tests/test_sap_bb_position.py`; updates to existing back-billing tests.

## 6. Non-goals

- No change to the 365-day legal gate or eligibility logic.
- No change to rebilling/domination logic or the existing "Surviving
  Invoices" period-charge total row — the union row is added alongside it.
- No new PDF extractor for the reconciliation statement beyond what exists;
  the SAP tab consumes existing parsed rows.
- Issues 1, 3, 4, 5 (Provenance formatting, Annual Summary links, Open PDF,
  evidence_files naming) remain out of scope for this spec.
