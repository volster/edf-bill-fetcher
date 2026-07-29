"""Processors — regex patterns and extraction helpers extracted from edf_collector.py.

Submodules:
- ``patterns`` — pre-compiled regex constants used by the amount/reading/period
  extractors and the multi-regex fallback chain.
- ``extraction`` — fallback extractor functions (``_fallback_inv_num``,
  ``_fallback_period_from``, ``_fallback_period_to``, ``_fallback_amount``)
  and PST/OST helpers (``_pst_attachment_filename``, ``_extract_sender_email``,
  ``_matches_domain_filter``).

Compat re-exports live in ``edf_collector.py`` so callers using
``from edf_collector import AMOUNT_PATTERNS`` continue to work.
"""
