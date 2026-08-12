"""Shared multi-regex fallback extractors — invoice number, billing period from/to, amount.

Single source of truth for the fallback-extractor chain that recovers
``inv_num`` / ``period_from`` / ``period_to`` / ``period_charge`` when the
canonical extractors miss, shared by ``collectors.engine`` and
``processors.extraction``.  Extracted during the modularization refactor so
the two sites cannot drift apart.

Each function returns ``(value, regex_name)`` so the Source Excerpt column
can show a regex-trace ("inv_num via _COVER_BLOCK_INV_RE; period via ...").

- ``fallback_inv_num`` — multi-regex invoice-number fallback chain
  (canonical ``_INV_NUMBER_RE`` → ``_CREDIT_NUMBER_RE`` → cover-body
  ``_COVER_BLOCK_INV_RE`` → loose bare-token ``_FALLBACK_INV_RE``).
- ``fallback_period_from`` / ``fallback_period_to`` — billing-period
  fallback chain (canonical ``_BILLING_PERIOD_RE`` → cover-body
  ``_COVER_BLOCK_PERIOD_RE``).
- ``fallback_amount`` — amount fallback chain (period-charge
  ``_PERIOD_CHARGE_RE`` → credit-total ``_CREDIT_TOTAL_RE`` →
  pound-amount ``_POUND_AMOUNT_FALLBACK_RE``).

Dependency regexes live in :mod:`edf_bill_fetcher.processors.patterns` so
the package is self-contained (no circular import back into the consumers).
"""

from __future__ import annotations

from edf_bill_fetcher.processors.patterns import (
    _BILLING_PERIOD_RE,
    _COVER_BLOCK_INV_RE,
    _COVER_BLOCK_PERIOD_RE,
    _CREDIT_NUMBER_RE,
    _CREDIT_TOTAL_RE,
    _FALLBACK_INV_RE,
    _INV_NUMBER_RE,
    _PERIOD_CHARGE_RE,
    _POUND_AMOUNT_FALLBACK_RE,
)

__all__ = [
    "fallback_amount",
    "fallback_inv_num",
    "fallback_period_from",
    "fallback_period_to",
]


def fallback_inv_num(text: str) -> tuple[str | None, str]:
    """Try invoice-number regexes in priority order and return the first hit.

    Iterates over the canonical invoice-number regex, then the cover-body
    regex, then a loose bare-token regex. Returns ``(value, regex_name)``
    or ``(None, "")`` when no pattern matches.
    """
    for label, pat in (
        ("_INV_NUMBER_RE", _INV_NUMBER_RE),
        ("_CREDIT_NUMBER_RE", _CREDIT_NUMBER_RE),
        ("_COVER_BLOCK_INV_RE", _COVER_BLOCK_INV_RE),
        ("_FALLBACK_INV_RE", _FALLBACK_INV_RE),
    ):
        m = pat.search(text[:3000])
        if m:
            val = m.group(1).strip() if m.lastindex else m.group(0)
            return val, label
    return None, ""


def fallback_period_from(text: str) -> tuple[str | None, str]:
    """Return (period_from_str, regex_name)."""
    m = _BILLING_PERIOD_RE.search(text[:3000])
    if m:
        return m.group(1).strip(), "_BILLING_PERIOD_RE"
    m = _COVER_BLOCK_PERIOD_RE.search(text[:3000])
    if m:
        return m.group(1).strip(), "_COVER_BLOCK_PERIOD_RE"
    return None, ""


def fallback_period_to(text: str) -> tuple[str | None, str]:
    """Return (period_to_str, regex_name)."""
    m = _BILLING_PERIOD_RE.search(text[:3000])
    if m:
        return m.group(2).strip(), "_BILLING_PERIOD_RE"
    m = _COVER_BLOCK_PERIOD_RE.search(text[:3000])
    if m:
        return m.group(2).strip(), "_COVER_BLOCK_PERIOD_RE"
    return None, ""


def fallback_amount(text: str) -> tuple[float | None, str]:
    """Return (amount, regex_name) or (None, "")."""
    m = _PERIOD_CHARGE_RE.search(text[:3000])
    if m:
        return float(m.group(1).replace(",", "")), "_PERIOD_CHARGE_RE"
    m = _CREDIT_TOTAL_RE.search(text[:3000])
    if m:
        return float(m.group(1).replace(",", "")), "_CREDIT_TOTAL_RE"
    m = _POUND_AMOUNT_FALLBACK_RE.search(text[:3000])
    if m:
        return float(m.group(1).replace(",", "")), "_POUND_AMOUNT_FALLBACK_RE"
    return None, ""
