"""Tests that fallback extractor functions are importable from the processors.extraction submodule.

All tests are RED at Phase 0 because ``edf_bill_fetcher.processors.extraction``
does not yet exist.
"""

from __future__ import annotations


def test_fallback_amount_importable() -> None:
    from edf_bill_fetcher.processors.extraction import _fallback_amount

    assert _fallback_amount is not None


def test_fallback_inv_num_importable() -> None:
    from edf_bill_fetcher.processors.extraction import _fallback_inv_num

    assert _fallback_inv_num is not None


def test_fallback_period_from_importable() -> None:
    from edf_bill_fetcher.processors.extraction import _fallback_period_from

    assert _fallback_period_from is not None


def test_fallback_period_to_importable() -> None:
    from edf_bill_fetcher.processors.extraction import _fallback_period_to

    assert _fallback_period_to is not None


def test_extract_sender_email_importable() -> None:
    from edf_bill_fetcher.processors.extraction import _extract_sender_email

    assert _extract_sender_email is not None


def test_matches_domain_filter_importable() -> None:
    from edf_bill_fetcher.processors.extraction import _matches_domain_filter

    assert _matches_domain_filter is not None


def test_pst_attachment_filename_importable() -> None:
    from edf_bill_fetcher.processors.extraction import _pst_attachment_filename

    assert _pst_attachment_filename is not None
