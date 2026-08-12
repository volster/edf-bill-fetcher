"""Fallback extractor functions and PST/OST helpers extracted from.

``edf_collector.py``.

This module is the single source of truth for:

- ``_fallback_inv_num`` — multi-regex invoice-number fallback chain
  (canonical → cover-body → loose bare-token).
- ``_fallback_period_from`` / ``_fallback_period_to`` — billing-period
  fallback chain (canonical → cover-body).
- ``_fallback_amount`` — amount fallback chain (period-charge →
  credit-total → pound-amount).
- ``_pst_attachment_filename`` / ``_extract_sender_email`` — PST/OST
  helpers re-exported from :mod:`edf_bill_fetcher.helpers.pst_resources`
  (the shared single source of truth); the underscore aliases keep the
  module's existing import surface stable.
- ``_matches_domain_filter`` — checks whether a sender email matches
  a comma-separated domain filter string.

Dependency regexes live in :mod:`edf_bill_fetcher.processors.patterns`
so the package is self-contained (no circular import back into
``edf_collector``).

Compat re-exports live in ``edf_collector.py`` so callers using
``from edf_collector import _fallback_amount`` continue to work.
"""

from __future__ import annotations

from edf_bill_fetcher.helpers.domain_filter import matches_domain_filter
from edf_bill_fetcher.helpers.pst_resources import (
    extract_sender_email,
    pst_attachment_filename,
)
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


def _fallback_inv_num(text: str) -> tuple[str | None, str]:
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


def _fallback_period_from(text: str) -> tuple[str | None, str]:
    """Return (period_from_str, regex_name)."""
    m = _BILLING_PERIOD_RE.search(text[:3000])
    if m:
        return m.group(1).strip(), "_BILLING_PERIOD_RE"
    m = _COVER_BLOCK_PERIOD_RE.search(text[:3000])
    if m:
        return m.group(1).strip(), "_COVER_BLOCK_PERIOD_RE"
    return None, ""


def _fallback_period_to(text: str) -> tuple[str | None, str]:
    """Return (period_to_str, regex_name)."""
    m = _BILLING_PERIOD_RE.search(text[:3000])
    if m:
        return m.group(2).strip(), "_BILLING_PERIOD_RE"
    m = _COVER_BLOCK_PERIOD_RE.search(text[:3000])
    if m:
        return m.group(2).strip(), "_COVER_BLOCK_PERIOD_RE"
    return None, ""


def _fallback_amount(text: str) -> tuple[float | None, str]:
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


_pst_attachment_filename = pst_attachment_filename
_extract_sender_email = extract_sender_email


_matches_domain_filter = matches_domain_filter


__all__ = [
    "_fallback_amount",
    "_fallback_inv_num",
    "_fallback_period_from",
    "_fallback_period_to",
    "_extract_sender_email",
    "_matches_domain_filter",
    "_pst_attachment_filename",
]
