"""SLC-aware compensation estimator (Wave 6d, Task 6).

Pure-pandas, deterministic (no LLM) estimate of indicative compensation
claims derived from the deduplicated evidence records:

  * ``back_billing_excess``  -- the day-ratio slice of the Period Charge
    attributable to the SLC 7A Excess Days, computed from
    :func:`edf_bill_fetcher.processors.detection.detect_back_billing`
    output (``Period Charge x min(excess_days, days_billed) /
    days_billed`` -- the same capped ratio the back-billing sheet uses,
    so the claim never exceeds the Period Charge).
  * ``credit_hold_interest`` -- interest on a credit balance
    (``Amount < 0``) held by the supplier beyond ``credit_hold_days``
    (default 90), at ``credit_interest_rate`` (default 0.02 = 2%, the
    current Ofgem credit-balance interest reference).
  * ``late_credit_interest``  -- interest from the statement ``Date`` to
    ``as_of`` on a credit balance that was never refunded (no later
    positive record within GBP0.50 of the balance magnitude).

Each output row carries ``{category, invoice_ref, date, base_amount,
days, rate, indicative_amount, legal_basis, disclaimer}``.  The
``legal_basis`` strings reuse the established statutory wording of the
back-billing sheet (Electricity Act 1989 s.84B / SLC 7A) and the
credit-balance condition SLC 21BA already referenced in
:mod:`edf_bill_fetcher.processors.detection`; no new law is cited.

No framework imports at module scope (processors-layer rule) -- only
pandas and sibling processor/helper modules.
"""

from __future__ import annotations

import re
from typing import Any

import pandas as pd

from edf_bill_fetcher.helpers.date_utils import _safe_to_datetime
from edf_bill_fetcher.processors.detection import detect_back_billing

DISCLAIMER: str = (
    "Indicative figures only -- computed deterministically from extracted "
    "records and not legal advice. Verify against original documents before "
    "use in any formal dispute."
)

# Established statutory wording, mirroring the back-billing sheet's legal
# context block (io/adapters/pdf.py).  No new law is cited.
_BACK_BILLING_BASIS: str = (
    "Electricity Act 1989 s.84B / Ofgem -- Standard Licence Condition 7A "
    "(SLC 7A) 12-month back-billing bar"
)
_CREDIT_INTEREST_BASIS: str = (
    "Ofgem credit-balance interest -- Standard Licence Condition 21BA "
    "(SLC 21BA) refund of credit balances"
)

_DEFAULT_CREDIT_HOLD_DAYS: int = 90
_DEFAULT_CREDIT_INTEREST_RATE: float = 0.02  # 2% -- Ofgem credit-balance interest reference
_REFUND_TOLERANCE: float = 0.50  # same GBP0.50 matching tolerance as _reversal_match


_ISO_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def _parse_date(value: object) -> pd.Timestamp | None:
    """Parse *value* to a Timestamp; None on failure.

    ISO ``YYYY-MM-DD`` strings are parsed as ISO explicitly (the
    dayfirst heuristic in ``_safe_to_datetime`` would otherwise read
    them day-first, e.g. ``2026-06-01`` -> 6 Jan 2026) -- matching the
    ``parse_to_sort_date`` convention in :mod:`helpers.date_utils`.
    """
    if isinstance(value, str):
        s = value.strip()
        if _ISO_DATE_RE.match(s):
            iso = pd.to_datetime(s, format="%Y-%m-%d", errors="coerce")
            if not pd.isna(iso):
                return iso
    dt = _safe_to_datetime(value)
    if isinstance(dt, pd.Timestamp) and not pd.isna(dt):
        return dt
    return None


def _as_of(config: dict[str, Any]) -> pd.Timestamp:
    """Resolve the valuation date from config, defaulting to today."""
    parsed = _parse_date(config.get("as_of"))
    return parsed if parsed is not None else pd.Timestamp.today().normalize()


def estimate_compensation(
    records_df: pd.DataFrame | None,
    config: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Estimate indicative compensation claims from evidence records.

    Deterministic, pure-pandas.  ``config`` may override:

    * ``credit_hold_days`` (int, default 90) -- the supplier's free
      credit-hold window; hold interest accrues only beyond it.
    * ``credit_interest_rate`` (float, default 0.02) -- annual rate for
      credit-balance interest.
    * ``as_of`` (date string or Timestamp, default today) -- valuation
      date for interest elapsed so far.

    Rows are emitted in a stable order: back-billing excess rows first
    (detector bill-date order), then credit rows (record date order).
    """
    cfg = config or {}
    try:
        hold_days = int(cfg.get("credit_hold_days", _DEFAULT_CREDIT_HOLD_DAYS))
    except (TypeError, ValueError):
        hold_days = _DEFAULT_CREDIT_HOLD_DAYS
    try:
        rate = float(cfg.get("credit_interest_rate", _DEFAULT_CREDIT_INTEREST_RATE))
    except (TypeError, ValueError):
        rate = _DEFAULT_CREDIT_INTEREST_RATE
    as_of = _as_of(cfg)

    rows: list[dict[str, Any]] = []
    if records_df is None or records_df.empty:
        return rows

    # --- (a) back-billing excess: reuse the canonical detector output ---
    bb = detect_back_billing(records_df)
    for _, b in bb.iterrows():
        excess = int(b.get("Excess Days", 0) or 0)
        days_billed = int(b.get("Days Billed", 0) or 0)
        charge = float(b.get("Period Charge (£)", 0.0) or 0.0)
        if excess <= 0 or days_billed <= 0 or charge <= 0:
            continue
        bill_date = _parse_date(b.get("Bill Date"))
        if bill_date is None:
            continue
        # Same capped day-ratio as the back-billing sheet: the claim can
        # never exceed the Period Charge even when excess > days billed.
        ratio = min(excess, days_billed) / days_billed
        rows.append(
            {
                "category": "back_billing_excess",
                "invoice_ref": str(b.get("Invoice #", "")),
                "date": bill_date.strftime("%Y-%m-%d"),
                "base_amount": round(charge, 2),
                "days": excess,
                "rate": None,
                "indicative_amount": round(charge * ratio, 2),
                "legal_basis": _BACK_BILLING_BASIS,
                "disclaimer": DISCLAIMER,
            }
        )

    # --- (b) credit-hold interest / (c) late-credit interest ---
    all_records: list[tuple[pd.Timestamp, float, str]] = []
    for _, r in records_df.iterrows():
        try:
            amount = float(r.get("Amount (£)", 0) or 0)
        except (TypeError, ValueError):
            continue
        date = _parse_date(r.get("Date"))
        if date is None:
            continue
        all_records.append((date, amount, str(r.get("Invoice #", ""))))
    all_records.sort(key=lambda c: c[0])
    credits = [(d, round(abs(a), 2), inv) for d, a, inv in all_records if a < 0]

    for date, balance, invoice in credits:
        # Deterministic "refunded" proxy: a later record with a positive
        # amount within GBP0.50 of the balance magnitude (mirrors the
        # _reversal_match tolerance used by the rebilling detector).
        refund_date: pd.Timestamp | None = None
        for later_date, later_amount, _ in all_records:
            if later_date <= date or later_amount <= 0:
                continue
            if abs(later_amount - balance) > _REFUND_TOLERANCE:
                continue
            if later_date > as_of:
                continue
            refund_date = later_date
            break

        hold_end = refund_date if refund_date is not None else as_of
        held_beyond = (hold_end - date).days - hold_days
        if held_beyond > 0:
            rows.append(
                {
                    "category": "credit_hold_interest",
                    "invoice_ref": invoice,
                    "date": date.strftime("%Y-%m-%d"),
                    "base_amount": balance,
                    "days": held_beyond,
                    "rate": rate,
                    "indicative_amount": round(rate * balance * held_beyond / 365, 2),
                    "legal_basis": _CREDIT_INTEREST_BASIS,
                    "disclaimer": DISCLAIMER,
                }
            )

        if refund_date is None:
            late_days = (as_of - date).days
            if late_days > 0:
                rows.append(
                    {
                        "category": "late_credit_interest",
                        "invoice_ref": invoice,
                        "date": date.strftime("%Y-%m-%d"),
                        "base_amount": balance,
                        "days": late_days,
                        "rate": rate,
                        "indicative_amount": round(rate * balance * late_days / 365, 2),
                        "legal_basis": _CREDIT_INTEREST_BASIS,
                        "disclaimer": DISCLAIMER,
                    }
                )

    return rows


__all__ = ["DISCLAIMER", "estimate_compensation"]
