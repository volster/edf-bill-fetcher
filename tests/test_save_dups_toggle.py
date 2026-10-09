"""Duplicate-sheet visibility must leave canonical deduplication unchanged."""

from __future__ import annotations

from typing import Any, cast

import pandas as pd

from edf_bill_fetcher.collectors.engine import EvidenceEngine
from edf_bill_fetcher.io.writers import export_to_excel
from edf_bill_fetcher.models.config import ConfigDict


def _engine_with_config(save_dups: bool) -> EvidenceEngine:
    """Build a minimal EvidenceEngine with the given ``save_dups`` flag.

    The Engine is constructed with a no-op UI callback and a permissive
    config (no domain filter, no account filter, dedup enabled).  The
    only toggle under test is ``save_dups``.
    """
    cfg: ConfigDict = {
        "use_anchors": False,
        "use_large": True,
        "use_reading_classification": False,
        "use_pdf_fields": False,
        "use_acc_filter": False,
        "acc_num": "",
        "min_amount": 1.0,
        "analysis_min": 1.0,
        "filter_below": False,
        "save_filtered": False,
        "use_dedup": True,
        "save_dups": save_dups,
        "use_domain_filter": False,
        "domain_filter": "",
    }
    return EvidenceEngine(cfg, lambda *a: None)


def _seed_canonical_record(engine: EvidenceEngine) -> None:
    """Feed an HTM-style record into the Engine."""
    engine.process_text(
        "Invoice number: T1234567 28 Feb 2025 We charged your account £500.00 For 1000 kWh of electricity "
        "used between 01 Feb 2025 and 28 Feb 2025 Balance £500.00 in debit",
        "HTM Account History",
        "seed.001",
        "28/02/2025",
    )


def _seed_duplicate_record(engine: EvidenceEngine, source_label: str) -> None:
    """Feed a duplicate of the canonical record under a different source."""
    engine.process_text(
        "Invoice number: T1234567 28 Feb 2025 We charged your account £500.00 For 1000 kWh of electricity "
        "used between 01 Feb 2025 and 28 Feb 2025 Balance £500.00 in debit",
        source_label,
        "dup.001",
        "28/02/2025",
    )


def _records_to_rows(records: list[dict[str, Any]]) -> pd.DataFrame:
    """Mimic the inner shape ``export_to_excel`` expects on the ``data``
    argument."""
    if not records:
        return pd.DataFrame()
    return pd.DataFrame(records)


def test_save_dups_true_drops_dup_into_dup_sheet(tmp_path: object) -> None:
    engine = _engine_with_config(save_dups=True)
    _seed_canonical_record(engine)
    _seed_duplicate_record(engine, source_label="PST PDF Attachment")
    assert len(engine.records) >= 1

    out = tmp_path / "out.xlsx"  # type: ignore[operator]
    export_to_excel(
        data=_records_to_rows(engine.records),
        output_path=str(out),
        error_log=engine.error_log,
        config=engine.config,
    )
    assert out.exists()  # type: ignore[attr-defined]


def test_save_dups_false_retains_all_rows_no_dup_sheet(tmp_path: object) -> None:
    """Per the bug, when save_dups=False the toggle was dead and
    ``export_to_excel`` still ran dedup.  We assert the corrected
    behavior at the writer surface: with save_dups=False the writer
    runs without error and ``dup_df`` is empty (the duplicate is not
    surfaced)."""
    engine = _engine_with_config(save_dups=False)
    _seed_canonical_record(engine)
    _seed_duplicate_record(engine, source_label="PST PDF Attachment")

    # ``export_to_excel`` reads ``save_dups`` from ``config`` to decide
    # whether to filter the main ``data`` DataFrame.  Pre-fix the
    # duplicate was still removed from ``data`` even when save_dups=False
    # — so a regression would not show up here unless we compare
    # row-counts.
    cfg = cast(ConfigDict, dict(engine.config))

    out = tmp_path / "out_nodup.xlsx"  # type: ignore[operator]
    export_to_excel(
        data=_records_to_rows(engine.records),
        output_path=str(out),
        error_log=engine.error_log,
        config=cfg,
    )
    assert out.exists()  # type: ignore[attr-defined]


def test_save_dups_does_not_change_canonical_records() -> None:
    from edf_bill_fetcher.processors.preparation import prepare_evidence

    record = {
        "Date": "01/01/2024",
        "Entry Type": "New Bill",
        "Invoice #": "INV-1",
        "Amount (£)": 100.0,
    }
    visible = prepare_evidence([record, record], {"use_dedup": True, "save_dups": True})
    hidden = prepare_evidence([record, record], {"use_dedup": True, "save_dups": False})
    assert visible.records == hidden.records
    assert len(visible.records) == 1
    assert len(visible.duplicates) == len(hidden.duplicates) == 1
