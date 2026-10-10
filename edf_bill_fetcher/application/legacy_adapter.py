"""Capture legacy record/state data without retaining runtime engine primitives."""

from __future__ import annotations

import hashlib
from typing import Any

from edf_bill_fetcher.domain.evidence import ExtractionSnapshot, Observation
from edf_bill_fetcher.domain.json_values import canonical_text, freeze_json
from edf_bill_fetcher.domain.requests import RunRequest
from edf_bill_fetcher.models.snapshot_adapter import to_snapshot_json


def capture_legacy(
    records: list[dict[str, Any]], state: dict[str, Any], request: RunRequest
) -> ExtractionSnapshot:
    """Detach original records and all data fields into an immutable extraction."""
    if not isinstance(records, list) or any(not isinstance(row, dict) for row in records):
        raise ValueError("Original records must be an array of objects")
    diagnostics: list[dict] = []
    raw_rows = to_snapshot_json(records, path="$.original_records", diagnostics=diagnostics)
    data = to_snapshot_json(state, path="$.legacy_state", diagnostics=diagnostics)
    for key in ("original_records", "records"):
        if key in data and data[key] != raw_rows:
            if key == "records" and data.get("original_records") == raw_rows:
                # A preserved v2 state may legitimately contain both views.
                continue
            raise ValueError(f"Supplied {key} conflicts with original evidence")
    if diagnostics:
        data.setdefault("snapshot_diagnostics", []).extend(diagnostics)
    observations = tuple(
        Observation(
            hashlib.sha256(canonical_text([request.run_id, ordinal, row]).encode()).hexdigest(),
            ordinal,
            freeze_json(row),
        )
        for ordinal, row in enumerate(raw_rows)
    )
    return ExtractionSnapshot(request, observations, freeze_json(data))
