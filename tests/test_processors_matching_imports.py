"""Tests that matching functions are importable from the processors.matching submodule.

All tests are RED at Phase 0 because ``edf_bill_fetcher.processors.matching``
does not yet exist.
"""

from __future__ import annotations


def test_build_evidence_index_importable() -> None:
    from edf_bill_fetcher.processors.matching import build_evidence_index

    assert build_evidence_index is not None


def test_infer_contracts_importable() -> None:
    from edf_bill_fetcher.processors.matching import infer_contracts

    assert infer_contracts is not None


def test_match_sap_events_to_edf_importable() -> None:
    from edf_bill_fetcher.processors.matching import match_sap_events_to_edf

    assert match_sap_events_to_edf is not None
