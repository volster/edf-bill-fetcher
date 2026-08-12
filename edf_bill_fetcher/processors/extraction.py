"""Fallback extractor functions and PST/OST helpers extracted from.

``edf_collector.py``.

The four fallback extractors — ``_fallback_inv_num`` (multi-regex
invoice-number fallback chain: canonical → cover-body → loose bare-token),
``_fallback_period_from`` / ``_fallback_period_to`` (billing-period fallback
chain: canonical → cover-body), and ``_fallback_amount`` (amount fallback
chain: period-charge → credit-total → pound-amount) — live in
:mod:`edf_bill_fetcher.helpers.fallback_extractors` (the shared single source
of truth); this module re-exports them via underscore aliases so the existing
import surface stays stable.

- ``_pst_attachment_filename`` / ``_extract_sender_email`` — PST/OST
  helpers re-exported from :mod:`edf_bill_fetcher.helpers.pst_resources`
  (the shared single source of truth); the underscore aliases keep the
  module's existing import surface stable.
- ``_matches_domain_filter`` — checks whether a sender email matches
  a comma-separated domain filter string.

Dependency regexes live in :mod:`edf_bill_fetcher.processors.patterns`
so the package is self-contained (no circular import back into
``edf_collector``).
"""

from __future__ import annotations

from edf_bill_fetcher.helpers.domain_filter import matches_domain_filter
from edf_bill_fetcher.helpers.fallback_extractors import (
    fallback_amount,
    fallback_inv_num,
    fallback_period_from,
    fallback_period_to,
)
from edf_bill_fetcher.helpers.pst_resources import (
    extract_sender_email,
    pst_attachment_filename,
)

_pst_attachment_filename = pst_attachment_filename
_extract_sender_email = extract_sender_email

_fallback_inv_num = fallback_inv_num
_fallback_period_from = fallback_period_from
_fallback_period_to = fallback_period_to
_fallback_amount = fallback_amount


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
