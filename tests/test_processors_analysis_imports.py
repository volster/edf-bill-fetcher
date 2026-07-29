"""Tests that analysis helper functions are importable from the processors.analysis submodule.

All tests are RED at Phase 0 because ``edf_bill_fetcher.processors.analysis``
does not yet exist.
"""

from __future__ import annotations


def test_disclosed_label_importable() -> None:
    from edf_bill_fetcher.processors.analysis import _disclosed_label

    assert _disclosed_label is not None


def test_reading_type_to_aem_importable() -> None:
    from edf_bill_fetcher.processors.analysis import _reading_type_to_aem

    assert _reading_type_to_aem is not None


def test_reversal_match_importable() -> None:
    from edf_bill_fetcher.processors.analysis import _reversal_match

    assert _reversal_match is not None


def test_compute_dispute_flags_importable() -> None:
    from edf_bill_fetcher.processors.analysis import compute_dispute_flags

    assert compute_dispute_flags is not None


def test_data_quality_report_importable() -> None:
    from edf_bill_fetcher.processors.analysis import _data_quality_report

    assert _data_quality_report is not None


def test_detect_payment_patterns_importable() -> None:
    from edf_bill_fetcher.processors.analysis import _detect_payment_patterns

    assert _detect_payment_patterns is not None


def test_analyze_tariff_impact_importable() -> None:
    from edf_bill_fetcher.processors.analysis import _analyze_tariff_impact

    assert _analyze_tariff_impact is not None
