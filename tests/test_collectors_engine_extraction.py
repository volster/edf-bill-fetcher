"""Tests for the EvidenceEngine extraction from the collectors submodule.

Verifies that EvidenceEngine is importable from
``edf_bill_fetcher.collectors.engine`` and that its core
methods work correctly.
"""

from __future__ import annotations

import base64
import os
import tempfile


def test_evidence_engine_instantiate() -> None:
    from edf_bill_fetcher.collectors.engine import EvidenceEngine

    engine = EvidenceEngine(config={}, update_ui_cb=lambda *args: None)
    assert engine is not None
    assert engine.records == []


def test_evidence_engine_process_pdf_file() -> None:
    from edf_bill_fetcher.collectors.engine import EvidenceEngine

    engine = EvidenceEngine(config={}, update_ui_cb=lambda *args: None)
    # A minimal inline base64-encoded PDF (1x1 pixel white image
    # wrapped in a valid PDF structure) is not a real EDF bill,
    # so process_pdf_file should either extract nothing or log
    # an error — but it must not raise ImportError or AttributeError.
    minimal_pdf = base64.b64encode(
        b"%PDF-1.0\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n"
        b"2 0 obj<</Type/Pages/Kids[]/Count 0>>endobj\n"
        b"xref\n0 3\n0000000000 65535 f \n0000000009 00000 n \n"
        b"0000000058 00000 n \ntrailer<</Size 3/Root 1 0 R>>\n"
        b"%%EOF\n"
    )

    # process_pdf_file expects a file path, not base64. Write the
    # bytes to a temp file, pass the path, then clean up.
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    try:
        tmp.write(minimal_pdf)
        tmp.close()
        engine.process_pdf_file(
            tmp.name,
            source_label="test",
            detail_label="test.pdf",
            fallback_date="2024-01-01",
        )
    finally:
        os.unlink(tmp.name)

    assert isinstance(engine.records, list)


def test_evidence_engine_pickle_roundtrip() -> None:
    from edf_bill_fetcher.collectors.engine import EvidenceEngine

    engine = EvidenceEngine(config={}, update_ui_cb=lambda *args: None)
    engine.records = [{"Invoice #": "T-001", "Amount": 100.0}]

    snapshot = engine.__getstate__()
    assert "records" in snapshot

    restored = EvidenceEngine.__new__(EvidenceEngine)
    restored.__setstate__(snapshot)

    assert restored.records == engine.records
    assert restored.config == engine.config
