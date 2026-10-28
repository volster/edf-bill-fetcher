"""Prepare preserved observations once using the existing conservative rules."""

from __future__ import annotations

from dataclasses import replace

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
    schedule = state.get("dispute_schedule", [])
    if schedule:
        active_ids = {
            row.get("Review Row ID")
            for row in prepared.records + state.get("review_excluded_records", [])
        }
        duplicate_ids = {row.get("Review Row ID") for row in prepared.duplicates}
        absorbed_positions = [
            position for position in schedule if position["row_id"] not in active_ids
        ]
        if any(position["row_id"] not in duplicate_ids for position in absorbed_positions):
            raise ValueError("Dispute position refers to an unknown record")
        if absorbed_positions:
            state.setdefault("archived_dispute_schedule", []).extend(absorbed_positions)
            state["dispute_schedule"] = [
                position for position in schedule if position["row_id"] in active_ids
            ]
            for position in absorbed_positions:
                state.setdefault("review_history", []).append(
                    {
                        "at": extraction.request.created_at.isoformat(),
                        "base_fingerprint": state.get("human_review", {}).get("base_fingerprint"),
                        "row_id": position["row_id"],
                        "invoice": None,
                        "field": "Dispute position",
                        "before": str(position),
                        "after": "archived_after_deduplication",
                        "reason": "Record merged by duplicate resolution; its proposal is archived, not added to the retained record's challenge.",
                    }
                )
            extraction = replace(extraction, legacy_state=freeze_json(state))
    claims = state.get("claim_reconciliation")
    if claims:
        active_ids = {
            row.get("Review Row ID")
            for row in prepared.records + state.get("review_excluded_records", [])
        }
        duplicate_ids = {row.get("Review Row ID") for row in prepared.duplicates}
        archived = []
        for group in claims["groups"]:
            references = set(group["members"]) | {credit["row_id"] for credit in group["credits"]}
            missing = references - active_ids
            if missing and not missing <= duplicate_ids:
                raise ValueError("Claim group refers to an unknown record")
            if missing:
                archived.append(group)
        if archived:
            archived_ids = {group["id"] for group in archived}
            state.setdefault("archived_claim_groups", []).extend(archived)
            claims["groups"] = [
                group for group in claims["groups"] if group["id"] not in archived_ids
            ]
            removed = [
                decision
                for decision in claims["separations"]
                if set(decision["groups"]) & archived_ids
            ]
            state.setdefault("archived_claim_separations", []).extend(removed)
            claims["separations"] = [
                decision
                for decision in claims["separations"]
                if not set(decision["groups"]) & archived_ids
            ]
            state.setdefault("review_history", []).append(
                {
                    "at": extraction.request.created_at.isoformat(),
                    "base_fingerprint": state.get("human_review", {}).get("base_fingerprint"),
                    "row_id": "",
                    "invoice": None,
                    "field": "Claim reconciliation",
                    "before": str(sorted(archived_ids)),
                    "after": "archived_after_deduplication",
                    "reason": "A group member or allocated credit merged during duplicate resolution; the proposal is archived, not transferred or added to another claim.",
                }
            )
            extraction = replace(extraction, legacy_state=freeze_json(state))
    statuses = state.get("human_review_statuses", {})
    if statuses:
        # Correcting an invoice can merge two previously distinct rows. Keep
        # the absorbed row's verification as history, rather than attaching
        # it to a different record or leaving an unreopenable active status.
        active_ids = {
            row.get("Review Row ID")
            for row in prepared.records + state.get("review_excluded_records", [])
        }
        absorbed = set(statuses) - active_ids
        duplicate_ids = {row.get("Review Row ID") for row in prepared.duplicates}
        if not absorbed <= duplicate_ids:
            raise ValueError("Review status refers to an unknown record")
        if absorbed:
            archived = state.setdefault("archived_review_statuses", {})
            history = state.setdefault("review_history", [])
            for row_id in sorted(absorbed):
                annotation = statuses.pop(row_id)
                archived[row_id] = annotation
                history.append(
                    {
                        "at": extraction.request.created_at.isoformat(),
                        "base_fingerprint": state.get("human_review", {}).get("base_fingerprint"),
                        "row_id": row_id,
                        "invoice": next(
                            (
                                row.get("Invoice #")
                                for row in inputs
                                if row.get("Review Row ID") == row_id
                            ),
                            None,
                        ),
                        "field": "Review status",
                        "before": annotation["status"],
                        "after": "archived_after_deduplication",
                        "reason": "Record merged by duplicate resolution; previous verification retained in archived_review_statuses.",
                    }
                )
            extraction = replace(extraction, legacy_state=freeze_json(state))
    return PreparedSnapshot(
        extraction,
        freeze_json(to_snapshot_json(prepared.records)),
        freeze_json(to_snapshot_json(prepared.duplicates)),
        freeze_json(to_snapshot_json(prepared.config)),
    )
