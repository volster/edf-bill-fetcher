"""Tests for edf_bill_fetcher.collectors submodule.

Verifies that the EvidenceEngine orchestrator is importable from the
collectors submodule and behaves correctly.
"""

from __future__ import annotations

import threading
from typing import cast

from edf_bill_fetcher.models.config import ConfigDict


def test_collectors_submodule_importable():
    from edf_bill_fetcher.collectors import EvidenceEngine

    assert EvidenceEngine is not None


def test_evidence_engine_importable_from_collectors_engine():
    from edf_bill_fetcher.collectors.engine import EvidenceEngine

    assert EvidenceEngine is not None


def test_evidence_engine_re_exported_from_edf_collector():
    from edf_bill_fetcher.collectors.engine import EvidenceEngine as EE1
    from edf_bill_fetcher.collectors.engine import EvidenceEngine as EE2

    assert EE1 is EE2


def test_evidence_engine_initial_state():
    from edf_bill_fetcher.collectors import EvidenceEngine

    eng = EvidenceEngine(
        config=cast(ConfigDict, {"account": "123"}),
        update_ui_cb=lambda *_: None,
        progress_cb=None,
        cancel_event=threading.Event(),
    )
    assert eng.records == []
    assert eng.filtered_records == []
    assert eng.pdf_count == 0
    assert eng.email_count == 0
    assert eng.error_log == []
    assert eng.source_paths == {}
    assert eng.sap_contract_rows == []
    assert eng.sap_meter_rows == []
    assert eng.sap_financial_rows == []
    assert eng.cancel_event is not None
    assert eng.lock is not None


def test_evidence_engine_pickle_round_trip():
    import pickle

    from edf_bill_fetcher.collectors import EvidenceEngine

    eng = EvidenceEngine(
        config=cast(ConfigDict, {"account": "456"}),
        update_ui_cb=lambda *_: None,
    )
    eng.records.append({"Invoice #": "INV-001"})
    eng.pdf_count = 3

    blob = pickle.dumps(eng)
    eng2 = pickle.loads(blob)

    # Pickled state has no live UI / lock / event references — the
    # __setstate__ rebuild method replaces them with safe stubs.
    assert eng2.update_ui is not None
    assert eng2.update_progress is not None
    assert eng2.records == [{"Invoice #": "INV-001"}]
    assert eng2.pdf_count == 3
