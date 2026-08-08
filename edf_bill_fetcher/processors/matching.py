"""Cross-source matching helpers: evidence index, contract inference, SAP<->EDF event match.

Extracted from ``edf_collector.py`` as part of the modularization refactor
(Task 5 - Phase 4).  Pure-pandas helpers keyed off the deduplicated evidence
DataFrame.

Compat re-exports live in ``edf_collector.py`` so callers using
``from edf_collector import infer_contracts`` continue to work;
stripped by Task 7.
"""

from __future__ import annotations

import pandas as pd

from edf_bill_fetcher.helpers.date_utils import _safe_to_datetime
from edf_bill_fetcher.models.events import SapBackBillingEvent, SapEdfMatch
from edf_bill_fetcher.processors.sap_parsers import _parse_amount_for_event

# SAP match band constants — local definitions to keep module self-contained.
_SAP_MATCH_AMOUNT_BANDS = ((0.05, 40), (0.25, 20), (0.50, 5))
_SAP_MATCH_DAY_BANDS = ((0, 50), (3, 25), (14, 5))


def _confidence_band(score: int) -> str | None:
    """Map a numeric SAP<->EDF match score to a confidence band name."""
    if score >= 75:
        return "High"
    if score >= 40:
        return "Medium"
    if score >= 10:
        return "Low"
    return None


def build_evidence_index(df: pd.DataFrame, header_row_offset: int = 1) -> dict[str, int]:
    """Map match-key signatures to the Excel row on the Evidence Report."""
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return {}
    index: dict[str, int] = {}
    rows_iter = df.iterrows()
    for i, r in rows_iter:
        row_no = header_row_offset + 1 + i  # Excel row (header row + i + 1)
        inv = r.get("Invoice #", "")
        if isinstance(inv, str) and inv and inv != "N/A":
            key = f"inv:{inv}"
            index.setdefault(key, row_no)
        amt = r.get("Amount (£)", "")
        pf = pd.to_datetime(r.get("Period From"), dayfirst=True, errors="coerce")
        pt = pd.to_datetime(r.get("Period To"), dayfirst=True, errors="coerce")
        if pd.isna(pf) or pd.isna(pt):
            continue
        try:
            amt_f = float(amt)
        except (TypeError, ValueError):
            continue
        days = str((pt - pf).days)
        index.setdefault(f"amt_days:{amt_f:.2f}|{days}", row_no)
    return index


def infer_contracts(df: pd.DataFrame, merge_gap_days: int = 30) -> pd.DataFrame:
    """Infer contract periods from tariff transitions (spec \u00a73.4).

    Walks the rows of *df* sorted by ``Date``, skips ``N/A`` tariffs,
    groups consecutive rows sharing the same ``Tariff`` into one
    contract, and merges adjacent same-tariff groups whose gap is
    shorter than ``merge_gap_days`` (default 30). Returns one row per
    contract with the start/end dates, total days, and invoice count.

    Output columns:
        Contract From, Contract To, Tariff, Days, # Invoices.
    """
    columns = ["Contract From", "Contract To", "Tariff", "Days", "# Invoices"]
    if df is None or df.empty:
        return pd.DataFrame(columns=columns)
    work = df.copy()
    work["_dt"] = _safe_to_datetime(work.get("Date"))
    work = work.dropna(subset=["_dt", "Tariff"])
    work = work[work["Tariff"] != "N/A"]
    if work.empty:
        return pd.DataFrame(columns=columns)
    work = work.sort_values("_dt").reset_index(drop=True)
    # Build raw runs: consecutive rows with the same tariff value.
    runs: list[dict] = []
    cur_start_idx = 0
    cur_tariff = work.iloc[0]["Tariff"]
    for i in range(1, len(work)):
        if work.iloc[i]["Tariff"] != cur_tariff:
            runs.append(
                {
                    "start_idx": cur_start_idx,
                    "end_idx": i - 1,
                    "tariff": cur_tariff,
                }
            )
            cur_start_idx = i
            cur_tariff = work.iloc[i]["Tariff"]
    runs.append(
        {
            "start_idx": cur_start_idx,
            "end_idx": len(work) - 1,
            "tariff": cur_tariff,
        }
    )
    # Merge adjacent runs of the same tariff if gap < merge_gap_days.
    merged: list[dict] = []
    for run in runs:
        # Calculate this run's dates.
        start_dt = work.iloc[run["start_idx"]]["_dt"]
        end_dt = work.iloc[run["end_idx"]]["_dt"]
        start_raw = work.iloc[run["start_idx"]]["Date"]
        end_raw = work.iloc[run["end_idx"]]["Date"]
        n = run["end_idx"] - run["start_idx"] + 1
        candidate = {
            "Contract From": start_raw,
            "Contract To": end_raw,
            "_from_dt": start_dt,
            "_to_dt": end_dt,
            "Tariff": run["tariff"],
            "# Invoices": n,
        }
        if merged and merged[-1]["Tariff"] == candidate["Tariff"]:
            prev_end = merged[-1]["_to_dt"]
            gap_days = (candidate["_from_dt"] - prev_end).days
            if 0 <= gap_days < merge_gap_days:
                # Merge: extend previous contract's end and invoice count.
                merged[-1]["Contract To"] = candidate["Contract To"]
                merged[-1]["_to_dt"] = candidate["_to_dt"]
                merged[-1]["# Invoices"] += candidate["# Invoices"]
                continue
        merged.append(candidate)
    rows = []
    for c in merged:
        days = int((c["_to_dt"] - c["_from_dt"]).days)
        rows.append(
            {
                "Contract From": c["Contract From"],
                "Contract To": c["Contract To"],
                "Tariff": c["Tariff"],
                "Days": days,
                "# Invoices": int(c["# Invoices"]),
            }
        )
    if not rows:
        return pd.DataFrame(columns=columns)
    out = pd.DataFrame(rows, columns=columns)
    sort_idx = _safe_to_datetime(out["Contract From"]).sort_values().index
    out = out.loc[sort_idx].reset_index(drop=True)
    return out[columns]


def match_sap_events_to_edf(
    events: list[SapBackBillingEvent],
    edf_records: list[dict],
) -> list[SapEdfMatch]:
    """Fuzzy-match SAP events to EDF Evidence Report rows.

    Spec §3.3.  Returns one SapEdfMatch per (event × matched EDF
    candidate).  SAP events with no candidate at Low confidence or
    above are omitted from the returned list (but remain on Sheet 1).
    """
    if not events or not edf_records:
        return []

    # Parse the EDF records into a list of (idx, period_from, period_to, amount, invoice).
    # Reuse the module-level _safe_to_datetime so EDF UK-format dates
    # (DD/MM/YYYY) are parsed day-first; the earlier inline ``_to_ts``
    # helper called ``pd.to_datetime`` with the default MM/DD, which
    # mis-split the smoking-gun cluster-vs-EDF pairing by ~30 days.
    parsed_edf: list[
        tuple[
            int,
            pd.Timestamp | pd._libs.tslibs.nattype.NaTType,
            pd.Timestamp | pd._libs.tslibs.nattype.NaTType,
            float,
            str,
        ]
    ] = []
    for i, rec in enumerate(edf_records):
        invoice = str(rec.get("Invoice #", "")).strip()
        if not invoice or invoice in ("N/A", "None"):
            continue
        pf = _safe_to_datetime(rec.get("Period From"))
        pt = _safe_to_datetime(rec.get("Period To"))
        # Canonical amount axis: Period Charge (£) is the charge for the
        # billing period; Amount (£) is the running balance. Prefer Period
        # Charge, falling back to Amount only when it is N/A/unparseable.
        period_charge_raw = rec.get("Period Charge (£)")
        amt_raw = rec.get("Amount (£)")
        try:
            if pd.isna(period_charge_raw) or str(period_charge_raw).strip().upper() in (
                "N/A",
                "NONE",
                "",
            ):
                raise ValueError("Period Charge is N/A")
            amt = float(str(period_charge_raw).replace(",", "").lstrip("£").strip())
        except (TypeError, ValueError):
            try:
                amt = float(str(amt_raw).replace(",", "").lstrip("£").strip())
            except (TypeError, ValueError):
                amt = 0.0
        parsed_edf.append((i, pf, pt, amt, invoice))

    matches: list[SapEdfMatch] = []
    for ev in events:
        if pd.isna(ev.clearing_date):
            continue
        ev_cd = pd.Timestamp(ev.clearing_date)
        posting_ts = pd.NaT
        for r in ev.rows:
            p_raw = r.get("Posting Date")
            if p_raw is None or str(p_raw).strip().upper() in ("", "N/A", "NONE"):
                continue
            p_ts = _safe_to_datetime(p_raw)
            if not pd.isna(p_ts):
                posting_ts = pd.Timestamp(p_ts)
                break
        for idx, pf, pt, edf_amt, _invoice in parsed_edf:
            # Compute the date delta in days vs Period To (or From fallback)
            if not pd.isna(pt):
                date_delta_days = int(abs((ev_cd - pd.Timestamp(pt)).days))
            elif not pd.isna(pf):
                date_delta_days = int(abs((ev_cd - pd.Timestamp(pf)).days))
            else:
                continue

            # Amount score — computed before the date-score branch
            # (spec §3.1 — Option C: the in-span date bonus is now
            # conditional on amount correspondence; previously a
            # SAP clearing date happening to fall inside a wide EDF
            # invoice period scored Medium with zero amount match,
            # flooding the matched-events sheet with 129 fake rows).
            amount_score = 0
            if abs(ev.net_amount) < 1.0 and edf_amt > 0:
                # Net-zero cluster: try matching any underlying row's
                # gross amount against the EDF invoice.  Find the best
                # (lowest) band that the closest row fits.
                best_rel_delta = float("inf")
                for r in ev.rows:
                    row_amt = _parse_amount_for_event(r.get("Amount"))
                    if abs(row_amt) < 1 or abs(edf_amt) < 1:
                        continue
                    rel_delta = abs(abs(row_amt) - edf_amt) / max(abs(edf_amt), 0.01)
                    if rel_delta < best_rel_delta:
                        best_rel_delta = rel_delta
                if best_rel_delta != float("inf"):
                    for band_amt, band_score in _SAP_MATCH_AMOUNT_BANDS:
                        if best_rel_delta <= band_amt:
                            amount_score = band_score
                            break
            elif ev.net_amount != 0 and edf_amt > 0:
                ratio = ev.net_amount / edf_amt
                if 0.95 <= ratio <= 1.05:
                    amount_score = 40
                elif 0.75 <= ratio <= 1.25:
                    amount_score = 20
                elif 0.50 <= ratio <= 1.50:
                    amount_score = 5

            # Date score — Posting Date (preferred) or Clearing Date within
            # the EDF Period span gets the 50-point bonus ONLY when amount
            # also matched (spec §3.1 — Option C).  Without amount
            # correspondence the in-span case falls through to the day-band
            # ladder measured against the nearer boundary, so a pure
            # coincidental date-in-span can no longer reach Medium.
            date_score = 0
            date_in_span = False
            date_axis = ev_cd
            if (
                not pd.isna(posting_ts)
                and not pd.isna(pf)
                and not pd.isna(pt)
                and pd.Timestamp(pf) <= posting_ts <= pd.Timestamp(pt)
            ):
                date_axis = posting_ts
            if (
                not pd.isna(pf)
                and not pd.isna(pt)
                and pd.Timestamp(pf) <= date_axis <= pd.Timestamp(pt)
            ):
                date_in_span = True
                if amount_score > 0:
                    date_score = 50
                else:
                    delta_to_pf = abs((date_axis - pd.Timestamp(pf)).days)
                    delta_to_pt = abs((date_axis - pd.Timestamp(pt)).days)
                    nearest_delta = min(delta_to_pf, delta_to_pt)
                    for band_days, band_score in _SAP_MATCH_DAY_BANDS:
                        if nearest_delta <= band_days:
                            date_score = band_score
                            break
            else:
                for band_days, band_score in _SAP_MATCH_DAY_BANDS:
                    if date_delta_days <= band_days:
                        date_score = band_score
                        break

            total_score = date_score + amount_score
            band = _confidence_band(total_score)
            # Gate Medium+ on amount correspondence (spec §3.1 —
            # Option C).  Pure-date matches (in-span or near-boundary)
            # cap at Low; below-Low total drops to None.
            if band in ("High", "Medium") and amount_score == 0:
                band = "Low" if total_score >= 10 else None
            if band is None:
                continue

            amt_delta = round(ev.net_amount - edf_amt, 2)
            if band == "High":
                notes = (
                    "Clearing date inside EDF period + amount within 5%"
                    if date_in_span
                    else f"Within {date_delta_days}d of period-end + amount within 5%"
                )
            elif band == "Medium":
                # amount_score > 0 is guaranteed by the Medium+ gate above.
                if date_in_span:
                    notes = "Clearing date inside EDF period + amount within 25%"
                else:
                    notes = f"Within {date_delta_days}d of period-end + amount within 25%"
            else:  # band == "Low"
                if date_in_span and amount_score == 0:
                    notes = (
                        "Clearing date inside EDF period but amounts do not "
                        "correspond — likely coincidental"
                    )
                else:
                    notes = f"Within {date_delta_days}d of period-end; may be coincidental"

            matches.append(
                SapEdfMatch(
                    event=ev,
                    edf_record=edf_records[idx],
                    confidence_band=band,
                    confidence_score=total_score,
                    amount_delta=amt_delta,
                    date_delta_days=date_delta_days,
                    notes=notes,
                )
            )
    return matches


# ---------------------------------------------------------------------------
# HTM account-history parser
# ---------------------------------------------------------------------------
#
# EDF MyAccount exports "Payments and Invoices" in HTML. The recurring
# row shapes we recognise:
#
#   "DD Mon YYYY We charged your account £X.XX For Y kWh between D Mon YYYY and D Mon YYYY Balance £Z.ZZ in debit|credit"
#   "DD Mon YYYY You paid us £X.XX [Bank Transfer] Balance £Z.ZZ in debit|credit"
#   "DD Mon YYYY Reversed account charge £X.XX Balance £Z.ZZ in debit|credit"
#   "DD Mon YYYY [Bank Transfer / nothing.] Balance £Z.ZZ in credit"  -- standalone
#                               credit-only balance lines that appear when
#                               the customer's overall balance is in credit
#                               and there is no transaction for the period.
#
# Pre-fix (#15): the Balance clause hard-required "in debit", silently
# dropping "in credit" rows.  This was a real, reproducible data loss.
#
# Each regex matches the trailing "Balance £X in (debit|credit)" with a
# non-grouping alternation so existing group numbers are preserved.


# ---------------------------------------------------------------------------
# Reconciliation statement detector + multi-row extractor
# ---------------------------------------------------------------------------
# A consolidated reconciliation statement PDF (e.g. EDF's "Bill reference:
# <N> (<date>) / Account number: A-<N> / Balance on your last bill £X /
# Charges... / Payments... / Your new balance £Y") lists many individual
# charge, reversal, late-payment, and payment rows under a single statement
# header.  Without this extractor it was parsed as a single row carrying only
# the "Balance on your last bill £37,301.48" line -- losing ~50 underlying
# rows.  Each emitted row is written through ``_add_record`` so it takes part
# in the same downstream dedup/analyser pipeline as a standalone PDF.

# _RECON_STATEMENT_RE helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.

# _RECON_CHARGE_RE helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.

# _RECON_REVERSAL_RE helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.

# _RECON_REVERSAL_PERIOD_RE helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.

# _RECON_LATE_PAYMENT_RE helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.

# _RECON_PAYMENT_RE helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.

# _RECON_BALANCE_LAST_RE helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.

# _RECON_NEW_BALANCE_RE helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.

_RECON_MONTH_MAP = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}


# _recon_to_iso helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.

# _recon_money helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.

# detect_reconciliation_statement moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); see re-export block at top of file.

# extract_reconciliation_statement_rows moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); see re-export block at top of file.

_PST_PR_ATTACH_LONG_FILENAME = 0x3707
_PST_PR_ATTACH_FILENAME = 0x3704


__all__ = [
    "build_evidence_index",
    "infer_contracts",
    "match_sap_events_to_edf",
]
