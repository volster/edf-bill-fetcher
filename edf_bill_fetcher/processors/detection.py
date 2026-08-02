"""Pure-pandas detectors for back-billing, rebilling, meter rollovers, PDF format, and reconciliation statements.

Extracted from ``edf_collector.py`` as part of the modularization refactor
(Task 5 - Phase 4).  Each detector returns a derived DataFrame; no LLM,
no external service.

Compat re-exports live in ``edf_collector.py`` so callers using
``from edf_collector import detect_back_billing`` continue to work;
stripped by Task 7.
"""

from __future__ import annotations

import re

import pandas as pd

from edf_bill_fetcher.helpers.date_utils import _safe_to_datetime, parse_to_display_date
from edf_bill_fetcher.io.adapters.pdf import legal_context as legal_context_fn  # noqa: E402,F401

# KI / KCR invoice-format presence regexes (still defined in edf_collector.py
# for backward compat).  Re-declare locally so detection.py is self-contained.
_KI_PRESENCE_RE = re.compile(r"invoice number:\s*KI-", re.IGNORECASE)
_KCR_PRESENCE_RE = re.compile(r"credit note number:\s*KCR-", re.IGNORECASE)

# Reconciliation statement regexes (kept local; private to this module).
_RECON_STATEMENT_RE = re.compile(
    r"Statement\s+reference:?\s*([A-Z0-9-]+).*?Statement\s+date:?\s*(\d{1,2}\s+\w+\s+\d{4})",
    re.IGNORECASE | re.DOTALL,
)
_RECON_BALANCE_LAST_RE = re.compile(
    r"Balance\s+(?:brought\s+forward|last\s+bill)\s*£([\d,]+\.\d{2})",
    re.IGNORECASE,
)
_RECON_NEW_BALANCE_RE = re.compile(
    r"Your\s+new\s+balance\s*£([\d,]+\.\d{2})",
    re.IGNORECASE,
)
_RECON_CHARGE_RE = re.compile(
    r"Electricity\s+charge[^\n]*?(\d{1,2}\s+\w+\s+\d{4})\s*[-–]\s*(\d{1,2}\s+\w+\s+\d{4})[^\n]*?£([\d,]+\.\d{2})",
    re.IGNORECASE,
)
_RECON_REVERSAL_RE = re.compile(
    r"Reversed\s+electricity\s+charge[^\n]*?(\d{1,2}\s+\w+\s+\d{4})[^\n]*?£([\d,]+\.\d{2})",
    re.IGNORECASE,
)
_RECON_REVERSAL_PERIOD_RE = re.compile(
    r"for\s+the\s+period\s+(\d{1,2}\s+\w+\s+\d{4})\s*[-–]\s*(\d{1,2}\s+\w+\s+\d{4})",
    re.IGNORECASE,
)
_RECON_LATE_PAYMENT_RE = re.compile(
    r"Late\s+payment[^\n]*?£([\d,]+\.\d{2})",
    re.IGNORECASE,
)
_RECON_PAYMENT_RE = re.compile(
    r"(\d{1,2}\s+\w+\s+\d{4})[^\n]*?£([\d,]+\.\d{2})",
    re.IGNORECASE,
)


# Reconciliation statement helpers (kept local; private to this module).
def _recon_to_iso(s: str) -> str:
    """Convert "DD Mon YYYY" to ISO "YYYY-MM-DD"; return "N/A" on failure."""
    try:
        return parse_to_display_date(s.strip())  # type: ignore[no-any-return]
    except Exception:
        return "N/A"


def _recon_money(s: str) -> float:
    """Parse a "1,234.56" string to float; return 0.0 on failure."""
    try:
        return float(s.replace(",", "").replace("£", "").strip())
    except (ValueError, AttributeError):
        return 0.0


# Helper used by detect_back_billing.  Stays local to keep module self-contained.
def _assess_reason(
    invoice: str,
    days: int,
    admitted: bool,
    period_from: pd.Timestamp,
    period_to: pd.Timestamp,
) -> str:
    """Return a short, deterministic narrative for the Reason Assessment

    column of the Back-billing sheet. Template-driven (no LLM).
    """
    pf = period_from.strftime("%d %b %Y")
    pt = period_to.strftime("%d %b %Y")
    excess = days - 365
    if admitted:
        head = (
            f"Invoice {invoice} billed {days} days ({pf} to {pt}), "
            f"{excess} days past the 12-month back-billing limit. "
            "EDF's cover page admits a cancellation/reversal, which is "
            "direct evidence the bill is a back-billing remedy."
        )
    else:
        head = (
            f"Invoice {invoice} billed {days} days ({pf} to {pt}), "
            f"{excess} days past the 12-month back-billing limit. No "
            "admit-phrase was found on the cover page."
        )
    return head


# Legal-context block for back-billing tab. Delegates to the canonical
# helper in ``io.adapters.pdf`` so any wording updates (and the test
# ``tests/test_legal_context.py`` that pins the first line) reach
# the back-billing sheet without further wiring.
legal_context = legal_context_fn


def detect_pdf_format(text):
    """Return 'new_invoice', 'new_credit', or 'old' based on document markers."""
    if _KI_PRESENCE_RE.search(text):
        return "new_invoice"
    if _KCR_PRESENCE_RE.search(text):
        return "new_credit"
    return "old"


# extract_new_invoice_fields moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); see re-export block at top of file.

# extract_new_credit_fields moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); see re-export block at top of file.

# _SAP_HEADER_RE helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.

# _SAP_CONTRACT_COLS helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.

# _SAP_METER_COLS helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.

# _SAP_FINANCIAL_COLS helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.

_SAP_DDMMYYYY_RE = re.compile(r"\b(\d{2})-(\d{2})-(\d{4})\b")


# ---------------------------------------------------------------------------
# Multi-regex fallback chain (Stream P3 / Task 5)
# ---------------------------------------------------------------------------
# Each fallback chain scans the input text in a fixed precedence order and
# returns ``(value, regex_name)`` so the Source Excerpt column can show the
# technical trace ("inv_num via _COVER_BLOCK_INV_RE; period via ..."). This
# reduces the N/A count on the analyser tabs (Back-billing, Rebilling,
# Meter Readings, Contract History) since many invoice PDFs sidestep the
# canonical "Invoice number: KI-<n>" / "Your charges: <from> - <to>" markers
# but still surface the data under alternative phrasings on the cover sheet.


def detect_back_billing(df: pd.DataFrame) -> pd.DataFrame:
    """Return invoices whose billing period exceeds 12 months.

    Back-billing (Ofgem / Electricity Act 1989 s.84B) bars suppliers
    from charging a domestic customer for energy supplied more than
    12 months before the bill that first raised the charge. This
    detector surfaces any single invoice whose ``Period From`` ->>
    ``Period To`` window exceeds 365 days, alongside whether the
    cover page admits a cancellation/reversal (the
    ``Cancel/Rebill Admitted`` column populated earlier in the
    pipeline by :func:`extract_admit_phrase`).

    The function tolerates a missing ``Cancel/Rebill Admitted``
    column (treated as ``False``).

    Output columns:
        Invoice #, Bill Date, Period From, Period To, Days Billed,
        Net Charge (£), 12-Month Limit (days), Excess Days,
        Cancel/Rebill Admitted, Reason Assessment.

    Rows with unparseable ``Period From``/``Period To`` are skipped
    silently. Output is sorted by ``Bill Date`` and re-indexed.

    Architectural note (SAP cross-feeding):
    This detector takes only the inferred-evidence dataframe. SAP
    data-dump rows (Contract-and-Product-Change-History,
    Meter-Read-History, Financial-Transactions) are surfaced in
    their own tabs (SAP Contract History / SAP Meter Readings /
    SAP Financial Transactions) plus the cross-source
    Reconciliation tab; they are NOT joined back into
    ``detect_back_billing`` because:

      * SAP financial transactions carry a Document No. (e.g.
        ``531000424090``) not an Invoice #, and their Transaction
        Text is the generic ledger description
        (``Dr- Consum Billing Receivable`` etc.) -- they cannot
        be unambiguously matched to an inferred invoice.
      * SAP records have no ``Period From`` / ``Period To`` span
        (only Posting Date / Document Date) so they cannot
        independently drive a back-billing judgement.
      * The Reconciliation sheet is the proper place to surface
        agreements and disagreements between the inferred and
        SAP samples; naively joining SAP amounts into the
        backbilling tab would mislead the reviewer.
    If a future resource joins the two sources by a higher-fidelity
    key (e.g. PDF receipt number + SAP Document No. mapping
    table), wire the intersection through ``run_analysers`` here.
    """
    columns = [
        "Invoice #",
        "Bill Date",
        "Period From",
        "Period To",
        "Days Billed",
        "Net Charge (£)",
        "12-Month Limit (days)",
        "Excess Days",
        "Cancel/Rebill Admitted",
        "Reason Assessment",
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=columns)
    has_admit = "Cancel/Rebill Admitted" in df.columns
    rows = []
    for _, r in df.iterrows():
        pf = _safe_to_datetime(r.get("Period From"))
        pt = _safe_to_datetime(r.get("Period To"))
        if pd.isna(pf) or pd.isna(pt):
            continue
        days = int((pt - pf).days)
        if days <= 365:
            continue
        net_raw = r.get("Amount (£)", 0)
        try:
            net = float(net_raw)
        except (TypeError, ValueError):
            net = 0.0
        admitted = bool(r.get("Cancel/Rebill Admitted")) if has_admit else False
        bill_date_raw = r.get("Date", "")
        bill_date_dt = _safe_to_datetime(bill_date_raw)
        rows.append(
            {
                "Invoice #": r.get("Invoice #", ""),
                "Bill Date": bill_date_raw,
                "_bill_date_sort": bill_date_dt if not pd.isna(bill_date_dt) else pd.Timestamp.max,
                "Period From": pf,
                "Period To": pt,
                "Days Billed": days,
                "Net Charge (£)": net,
                "12-Month Limit (days)": 365,
                "Excess Days": days - 365,
                "Cancel/Rebill Admitted": admitted,
                "Reason Assessment": _assess_reason(r.get("Invoice #", ""), days, admitted, pf, pt),
            }
        )
    out = pd.DataFrame(rows)
    if out.empty:
        return pd.DataFrame(columns=columns)
    sort_key = out["_bill_date_sort"]
    out = out.drop(columns=["_bill_date_sort"])
    # Reorder rows by the sort key (parsed Bill Date, ascending).
    out = out.loc[sort_key.sort_values().index].reset_index(drop=True)
    return out[columns]


def _disclosed_label(
    admitted: bool,
    overlaps: bool,
) -> str:
    """Return the human-readable value of the 'Cancel/Rebill Disclosed'

    cell used on the Back-billing and Rebilling tabs.

    The disclosed column joins two independent signals:
      * admit-phrase (the cover-page wording 'we've recently
        cancelled some charges for you'), captured as a bool on the
        record; and
      * period overlap, flagged by :func:`detect_rebilling`.
    """
    if admitted and overlaps:
        return "Admitted + overlap"
    if admitted:
        return "Admitted phrase"
    if overlaps:
        return "Period overlap"
    return ""


def _reversal_match(
    evidence_df: pd.DataFrame | None,
    killed_inv: str,
    killed_amount: float | None,
    killed_pf: pd.Timestamp,
    killed_pt: pd.Timestamp,
) -> bool:
    """Return whether a reversal-credit row in *evidence_df* matches the

    killed invoice well enough to count as rebilling evidence.

    Spec ref: 2026-07-16 §11. A reversal credit accepts the killed
    invoice when its amount is within ±£0.50 AND either its period
    overlaps the killed period by ≥ 30 days OR its period is
    unparseable (so we accept on amount alone, Entry Type == Credit).
    """
    if evidence_df is None or evidence_df.empty:
        return False
    if "Entry Type" not in evidence_df.columns:
        return False
    try:
        amount = abs(float(killed_amount or 0.0))
    except (TypeError, ValueError):
        return False
    matching = evidence_df[evidence_df["Entry Type"].isin(["Credit", "Payment"])]
    for _, row in matching.iterrows():
        try:
            row_amt = abs(float(row.get("Amount (£)", 0) or 0))
        except (TypeError, ValueError):
            continue
        if abs(row_amt - amount) > 0.50:
            continue
        rpf = _safe_to_datetime(row.get("Period From"))
        rpt = _safe_to_datetime(row.get("Period To"))
        if pd.isna(rpf) or pd.isna(rpt):
            return True
        overlap = (min(killed_pt, rpt) - max(killed_pf, rpf)).days
        if overlap >= 30:
            return True
    return False


def detect_rebilling(
    df: pd.DataFrame,
    *,
    evidence_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """Return cancel-and-repost pairs identified by the rebilling

    heuristic (spec §11, tightened gate).

    For each ordered pair ``(Killer, Killed)`` where ``Killer.Date``
    is strictly later than ``Killed.Date``, emit a row IFF ALL hold:

    1. ``Killer.Period From ≤ Killed.Period From AND Killer.Period To ≥
       Killed.Period To`` -- the killer's billing window fully contains
       the killed's billing window.
    2. ANY of these signals also fires:
       - ``Killer.Days Billed ≥ 365`` (wholesale cancel-and-repost of a
         long period),
       - the killer invoice has ``Cancel/Rebill Admitted = True``
         (an admission phrase like ``corrected`` / ``amended`` was
         detected on the source PDF), OR
       - a reversal credit row in ``evidence_df`` matches the killed
         invoice's amount within ±£0.50 and period overlap ≥ 30 days
         (or its period is unparseable, in which case amount alone
         suffices).

    Output columns:
        Killer Invoice, Killed Invoice, Killer Date, Killed Date,
        Period Overlap (days), Jump-back (days), Trigger Reason,
        Cancel/Rebill Admitted (Killer).

    ``Cancel/Rebill Admitted (Killer)`` is the admit-phrase flag
    lifted from the killer invoice.

    ``evidence_df`` is optional -- when omitted, the reversal-credit
    check is skipped and only the long-period / admit-phrase signals
    fire. ``run_analysers`` passes the evidence DataFrame so the
    reversal signal participates in normal pipeline use.
    """
    columns = [
        "Killer Invoice",
        "Killed Invoice",
        "Killer Date",
        "Killed Date",
        "Period Overlap (days)",
        "Jump-back (days)",
        "Trigger Reason",
        "Cancel/Rebill Admitted (Killer)",
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=columns)
    has_admit = "Cancel/Rebill Admitted" in df.columns
    rows = []
    parsed = []
    for _, r in df.iterrows():
        pf = _safe_to_datetime(r.get("Period From"))
        pt = _safe_to_datetime(r.get("Period To"))
        bd = _safe_to_datetime(r.get("Date"))
        if pd.isna(pf) or pd.isna(pt) or pd.isna(bd):
            continue
        try:
            amount = float(r.get("Amount (£)", 0) or 0)
        except (TypeError, ValueError):
            amount = None
        admitted = bool(r.get("Cancel/Rebill Admitted")) if has_admit else False
        parsed.append(
            {
                "Invoice #": r.get("Invoice #", ""),
                "Date_raw": r.get("Date", ""),
                "Date": bd,
                "Period From": pf,
                "Period To": pt,
                "Days Billed": int((pt - pf).days),
                "Amount": amount,
                "Admitted": admitted,
            }
        )
    if len(parsed) < 2:
        return pd.DataFrame(columns=columns)
    parsed.sort(key=lambda x: x["Date"])
    for i, killer in enumerate(parsed):
        for killed in parsed[:i]:
            # Containment -- the only structural requirement.
            if not (
                killer["Period From"] <= killed["Period From"]
                and killer["Period To"] >= killed["Period To"]
            ):
                continue
            triggers: list[str] = []
            if killer["Days Billed"] >= 365:
                triggers.append("killer period \u2265 365d")
            admitted = killer["Admitted"]
            if admitted:
                triggers.append("admit-phrase on killer")
            reversal_match = _reversal_match(
                evidence_df,
                killed["Invoice #"],
                killed["Amount"],
                killed["Period From"],
                killed["Period To"],
            )
            if reversal_match:
                triggers.append("reversal credit row matches killed")
            if not triggers:
                continue
            trigger_reason = "; ".join(triggers)
            overlap_d = max(
                0,
                (
                    min(killer["Period To"], killed["Period To"])
                    - max(killer["Period From"], killed["Period From"])
                ).days,
            )
            jumpback_d = (killed["Period From"] - killer["Period From"]).days
            rows.append(
                {
                    "Killer Invoice": killer["Invoice #"],
                    "Killed Invoice": killed["Invoice #"],
                    "Killer Date": killer["Date_raw"],
                    "Killed Date": killed["Date_raw"],
                    "Period Overlap (days)": overlap_d,
                    "Jump-back (days)": max(0, jumpback_d),
                    "Trigger Reason": trigger_reason,
                    "Cancel/Rebill Admitted (Killer)": admitted,
                }
            )
    if not rows:
        return pd.DataFrame(columns=columns)
    out = pd.DataFrame(rows, columns=columns)
    out["_k_sort"] = _safe_to_datetime(out["Killer Date"])
    out["_d_sort"] = _safe_to_datetime(out["Killed Date"])
    sort_idx = out.sort_values(["_k_sort", "_d_sort"]).index
    out = out.loc[sort_idx].drop(columns=["_k_sort", "_d_sort"]).reset_index(drop=True)
    return out[columns]


# Default 99,999 - 5,000 rollover threshold per spec \u00a73.3.
_DEFAULT_ROLLOVER_THRESHOLD = 99999 - 5000


def detect_meter_rollover(
    df: pd.DataFrame, rollover_threshold: int = _DEFAULT_ROLLOVER_THRESHOLD
) -> pd.DataFrame:
    """Return meter-rollover candidate events (spec \u00a73.3).

    Walks the rows of *df* keeping only ones tagged ``Actual'' or
    ``Smart'' in the ``Reading`` column (supplier-confirmed readings
    only -- ``Estimated``/``Unknown`` rows don't count). For each
    consecutive (actual-or-smart, actual-or-smart) pair, computes
    delta = (curr Units (kWh)) - (prev Units (kWh)) -- i.e. the
    change in per-period kWh consumption -- and emits a row when the
    delta is negative AND its magnitude exceeds
    ``rollover_threshold`` (default 99,999 - 5,000 = 94,999).

    Output columns:
        Date, Invoice #, Prev Units (kWh), Curr Units (kWh),
        Delta, Reading Type, Notes.

    Rows with unparseable ``Units (kWh)`` or ``Date`` are skipped
    silently.
    """
    columns = [
        "Date",
        "Invoice #",
        "Prev Units (kWh)",
        "Curr Units (kWh)",
        "Delta",
        "Reading Type",
        "Notes",
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=columns)
    # Restrict to Actual/Smart only.
    mask = df.get("Reading", pd.Series(dtype=str)).isin(["Actual", "Smart"])
    candidates = df[mask].copy()
    if candidates.empty:
        return pd.DataFrame(columns=columns)
    # Parse dates so we can sort.
    candidates["_date_dt"] = _safe_to_datetime(candidates["Date"])
    candidates = candidates.dropna(subset=["_date_dt"])
    candidates = candidates.sort_values("_date_dt")
    rows = []
    prev_units: float | None = None
    prev_invoice = ""
    prev_date_raw = ""
    for _, r in candidates.iterrows():
        u_raw = r.get("Units (kWh)", "N/A")
        try:
            u = float(u_raw)
        except (TypeError, ValueError):
            prev_units = None
            continue
        if prev_units is not None:
            delta = u - prev_units
            if delta < 0 and abs(delta) > rollover_threshold:
                rows.append(
                    {
                        "Date": r.get("Date", ""),
                        "Invoice #": r.get("Invoice #", ""),
                        "Prev Units (kWh)": prev_units,
                        "Curr Units (kWh)": u,
                        "Delta": int(delta),
                        "Reading Type": r.get("Reading", ""),
                        "Notes": (
                            f"Negative jump of {abs(int(delta))} kWh between "
                            f"{prev_invoice} ({prev_date_raw}) and "
                            f"{r.get('Invoice #', '')} ({r.get('Date', '')}) -- "
                            "consistent with a meter rollover near the "
                            f"{rollover_threshold + 5000}-rollover cap."
                        ),
                    }
                )
        prev_units = u
        prev_invoice = r.get("Invoice #", "")
        prev_date_raw = r.get("Date", "")
    if not rows:
        return pd.DataFrame(columns=columns)
    out = pd.DataFrame(rows, columns=columns)
    sort_idx = _safe_to_datetime(out["Date"]).sort_values().index
    out = out.loc[sort_idx].reset_index(drop=True)
    return out[columns]


def detect_reconciliation_statement(text: str) -> bool:
    return bool(_RECON_STATEMENT_RE.search(text[:2000]))


def extract_reconciliation_statement_rows(text: str, attachment_name: str) -> list[dict]:
    """Extract every charge, reversal, late-payment, payment + one meta row

    from a consolidation reconciliation statement PDF.
    """
    rows: list[dict] = []
    src = "Statement Reconciliation"

    def _excerpt_around(m: re.Match, window: int = 400) -> str:
        """Return up to ``window`` chars around the regex match."""
        start = max(0, m.start(0) - 20)
        end = min(len(text), m.end(0) + window)
        return text[start:end]

    bill_ref = ""
    bill_date_display = "N/A"
    bill_ref_match = _RECON_STATEMENT_RE.search(text)
    if bill_ref_match:
        bill_ref = bill_ref_match.group(1)
        bill_date_display = _recon_to_iso(bill_ref_match.group(2))

    bal_last: object = "N/A"
    bal_last_match = _RECON_BALANCE_LAST_RE.search(text)
    if bal_last_match:
        bal_last = _recon_money(bal_last_match.group(1))

    new_bal: object = "N/A"
    new_bal_match = _RECON_NEW_BALANCE_RE.search(text)
    if new_bal_match:
        new_bal = _recon_money(new_bal_match.group(1))

    # Charge rows
    for m in _RECON_CHARGE_RE.finditer(text):
        rows.append(
            {
                "Source": src,
                "Sender": "",
                "Date": bill_date_display,
                "Period From": _recon_to_iso(m.group(1)),
                "Period To": _recon_to_iso(m.group(2)),
                "Invoice #": bill_ref or "N/A",
                "Amount (£)": _recon_money(m.group(3)),
                "Period Charge (£)": _recon_money(m.group(3)),
                "Entry Type": "Charge",
                "Reading": "N/A",
                "Units (kWh)": "N/A",
                "Standing Chg (p/day)": "N/A",
                "Tariff": "N/A",
                "Attachment Name": attachment_name,
                "Details": "Electricity charge (reconciliation statement)",
                "Logic Used": "Reconciliation Statement Charge",
                "Balance Last Bill (£)": bal_last,
                "Source PDF Text": _excerpt_around(m),
                "_regex_trace": "recon _RECON_CHARGE_RE",
            }
        )

    # Reversed-electricity-charge rows
    for m in _RECON_REVERSAL_RE.finditer(text):
        date_iso = _recon_to_iso(m.group(1))
        amount = _recon_money(m.group(2))
        # Look for a parenthetical period on the next non-empty line.
        details = "Reversed electricity charge"
        tail = text[m.end() : m.end() + 400]
        period_match = _RECON_REVERSAL_PERIOD_RE.search(tail)
        if period_match:
            details = (
                f"Reversed electricity charge ({period_match.group(1)} - {period_match.group(2)})"
            )
        rows.append(
            {
                "Source": src,
                "Sender": "",
                "Date": date_iso,
                "Period From": "N/A",
                "Period To": "N/A",
                "Invoice #": bill_ref or "N/A",
                "Amount (£)": -abs(amount),
                "Period Charge (£)": -abs(amount),
                "Entry Type": "Credit",
                "Reading": "N/A",
                "Units (kWh)": "N/A",
                "Standing Chg (p/day)": "N/A",
                "Tariff": "N/A",
                "Attachment Name": attachment_name,
                "Details": details,
                "Logic Used": "Reconciliation Statement Reversal",
                "Balance Last Bill (£)": bal_last,
                "Source PDF Text": _excerpt_around(m),
                "_regex_trace": "recon _RECON_REVERSAL_RE",
            }
        )

    # Late payment rows
    for m in _RECON_LATE_PAYMENT_RE.finditer(text):
        amount = _recon_money(m.group(1))
        rows.append(
            {
                "Source": src,
                "Sender": "",
                "Date": bill_date_display,
                "Period From": "N/A",
                "Period To": "N/A",
                "Invoice #": bill_ref or "N/A",
                "Amount (£)": amount,
                "Period Charge (£)": amount,
                "Entry Type": "Late Payment",
                "Reading": "N/A",
                "Units (kWh)": "N/A",
                "Standing Chg (p/day)": "N/A",
                "Tariff": "N/A",
                "Attachment Name": attachment_name,
                "Details": "Late Payment Charge (reconciliation statement)",
                "Logic Used": "Reconciliation Statement Late Payment",
                "Balance Last Bill (£)": bal_last,
                "Source PDF Text": _excerpt_around(m),
                "_regex_trace": "recon _RECON_LATE_PAYMENT_RE",
            }
        )

    # Payment rows -- scoped to the section starting "Payments" through either
    # "Your new balance" or end-of-text. EDF lists payments with a date column
    # then a £ column.
    payments_block = ""
    pay_section_match = re.search(r"Payments\s*\n", text, re.IGNORECASE)
    if pay_section_match:
        block_start = pay_section_match.end()
        # End payment block at "Your new balance" or end-of-text.
        end_match = re.search(r"Your\s+new\s+balance", text[block_start:], re.IGNORECASE)
        block_end = block_start + end_match.start() if end_match else len(text)
        payments_block = text[block_start:block_end]

    if payments_block:
        for m in _RECON_PAYMENT_RE.finditer(payments_block):
            rows.append(
                {
                    "Source": src,
                    "Sender": "",
                    "Date": _recon_to_iso(m.group(1)),
                    "Period From": "N/A",
                    "Period To": "N/A",
                    "Invoice #": bill_ref or "N/A",
                    "Amount (£)": _recon_money(m.group(2)),
                    "Period Charge (£)": "N/A",
                    "Entry Type": "Payment",
                    "Reading": "N/A",
                    "Units (kWh)": "N/A",
                    "Standing Chg (p/day)": "N/A",
                    "Tariff": "N/A",
                    "Attachment Name": attachment_name,
                    "Details": "Payment received (reconciliation statement)",
                    "Logic Used": "Reconciliation Statement Payment",
                    "Balance Last Bill (£)": bal_last,
                    "Source PDF Text": _excerpt_around(m),
                    "_regex_trace": "recon _RECON_PAYMENT_RE",
                }
            )

    # Always emit one meta row carrying the statement-level context.
    rows.append(
        {
            "Source": src,
            "Sender": "",
            "Date": bill_date_display,
            "Period From": "N/A",
            "Period To": "N/A",
            "Invoice #": bill_ref or "N/A",
            "Amount (£)": new_bal,
            "Period Charge (£)": "N/A",
            "Entry Type": "Statement Reconciliation",
            "Reading": "N/A",
            "Units (kWh)": "N/A",
            "Standing Chg (p/day)": "N/A",
            "Tariff": "N/A",
            "Attachment Name": attachment_name,
            "Balance Last Bill (£)": bal_last,
            "Details": f"Statement reconciliation: bill ref {bill_ref}",
            "Logic Used": "Reconciliation Statement Meta",
            # The meta row carries the statement-level context
            # (bill ref + balances); there is no single regex match
            # to excerpt. Provide the first 600 chars of the statement
            # so a reviewer sees the statement header context.
            "Source PDF Text": text[:600],
            "_regex_trace": "recon meta",
        }
    )
    return rows


__all__ = [
    "detect_pdf_format",
    "detect_back_billing",
    "detect_rebilling",
    "detect_meter_rollover",
    "detect_reconciliation_statement",
    "extract_reconciliation_statement_rows",
]
