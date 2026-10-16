"""Prepare preserved observations once using the existing conservative rules."""

from __future__ import annotations

from edf_bill_fetcher.domain.evidence import ExtractionSnapshot, PreparedSnapshot
from edf_bill_fetcher.domain.json_values import freeze_json, object_value
from edf_bill_fetcher.models.snapshot_adapter import to_snapshot_json
from edf_bill_fetcher.processors.preparation import prepare_evidence


def prepare_snapshot(extraction: ExtractionSnapshot) -> PreparedSnapshot:
    """Resolve canonical rows without mutating captured originals or reading a clock."""
    inputs = [object_value(row.raw) for row in extraction.observations]
    state = object_value(extraction.legacy_state)
    if "human_review" in state:
        from edf_bill_fetcher.application.record_review import apply_review

        inputs, _ = apply_review(state["human_review"])
    prepared = prepare_evidence(
        inputs,
        object_value(extraction.request.config),
    )
    return PreparedSnapshot(
        extraction,
        freeze_json(to_snapshot_json(prepared.records)),
        freeze_json(to_snapshot_json(prepared.duplicates)),
        freeze_json(to_snapshot_json(prepared.config)),
    )
