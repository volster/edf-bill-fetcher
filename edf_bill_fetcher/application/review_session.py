"""Resumable human review drafts, with shared validation and bounded undo/redo."""

from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path
from typing import Any

from edf_bill_fetcher.application.record_review import (
    REVIEW_FIELDS,
    apply_review,
    corrected_record,
    input_value,
    review_inputs,
    review_rows,
    validate_review_statuses,
)
from edf_bill_fetcher.domain.analysis import AnalysisSnapshot
from edf_bill_fetcher.domain.json_values import canonical_text, object_value, strict_loads
from edf_bill_fetcher.helpers.artifacts import write_json
from edf_bill_fetcher.io.persistence.result_codec import encode_result
from edf_bill_fetcher.processors.rebilling_ledger import (
    build_ledger,
    schedule_digest,
    validate_schedule,
)
from edf_bill_fetcher.processors.revisions import build_revisions, validate_decisions

MAX_DRAFT_BYTES = 8 * 1024 * 1024


class ReviewSession:
    """Keep edits separate from the accepted snapshot; publish only through regeneration."""

    def __init__(self, snapshot: AnalysisSnapshot) -> None:
        """Capture stable row identities and bind drafts to the entire accepted result."""
        self.snapshot = snapshot
        self.rows = review_rows(snapshot)
        self.index = {row["id"]: row for row in self.rows}
        self.base_digest = hashlib.sha256(
            canonical_text(encode_result(snapshot, "review-base")).encode()
        ).hexdigest()
        state = object_value(snapshot.prepared.extraction.legacy_state)
        self.decisions: dict[str, dict] = {}
        self.statuses: dict[str, dict] = copy.deepcopy(state.get("human_review_statuses", {}))
        self.revisions: dict[tuple[str, str], dict] = {
            (d["older"], d["newer"]): dict(d)
            for d in validate_decisions(self.rows, state.get("bill_revision_decisions", []))
        }
        self.schedule: dict[str, dict] = {
            d["row_id"]: d for d in validate_schedule(self.rows, state.get("dispute_schedule", []))
        }
        self.undo_stack: list[list[tuple]] = []
        self.redo_stack: list[list[tuple]] = []
        self._validate_statuses(self.statuses)

    def decision(self, row_id: str) -> dict:
        """Return a detached complete decision for one original canonical row."""
        if row_id not in self.index:
            raise ValueError("Unknown review row")
        return copy.deepcopy(
            self.decisions.get(
                row_id,
                {
                    "row_id": row_id,
                    "include": self.index[row_id]["included"],
                    "changes": {},
                    "reason": "",
                },
            )
        )

    def record(self, row_id: str) -> dict:
        """Return the effective draft record, preserving original raw observations."""
        return corrected_record(self.index[row_id]["record"], self.decision(row_id)["changes"])

    def items(self) -> list[dict]:
        """Expose identities, inclusion and effective fields for relationship validation."""
        return [
            {
                "id": row["id"],
                "included": self.decision(row["id"])["include"],
                "record": self.record(row["id"]) if row["id"] in self.decisions else row["record"],
            }
            for row in self.rows
        ]

    def update(self, row_id: str, changes: dict, include: bool, reason: str, status: str) -> None:
        """Validate a record edit and its review state as one undoable command."""
        decision = self.decision(row_id)
        if not isinstance(changes, dict) or not set(changes) <= REVIEW_FIELDS.keys():
            raise ValueError("Unsupported review fields")
        for field, value in changes.items():
            normalized = input_value(field, value)
            original = self.index[row_id]["record"].get(field)
            try:
                original = input_value(field, original)
            except ValueError:
                pass
            if normalized == original:
                decision["changes"].pop(field, None)
            else:
                decision["changes"][field] = normalized
        decision.update(include=include, reason=reason)
        apply_review(
            {
                "version": 1,
                "base_fingerprint": self.snapshot.input_fingerprint,
                "rows": [self.index[row_id]],
                "decisions": [decision],
            }
        )
        annotation = {"status": status, "reason": reason}
        self._validate_statuses({row_id: annotation})
        effective = self.items()
        for item in effective:
            if item["id"] == row_id:
                item.update(
                    record=corrected_record(self.index[row_id]["record"], decision["changes"]),
                    included=include,
                )
        validate_decisions(effective, list(self.revisions.values()))
        validate_schedule(effective, list(self.schedule.values()))
        self._commit([("decisions", row_id, decision), ("statuses", row_id, annotation)])

    def set_revision(self, older: str, newer: str, kind: str, status: str, reason: str) -> None:
        """Record a confirmed/rejected relationship, or clear a previous decision."""
        pair = older, newer
        candidate = {**self.revisions}
        value = (
            None
            if status == "suggested"
            else {"older": older, "newer": newer, "kind": kind, "status": status, "reason": reason}
        )
        if value is None:
            candidate.pop(pair, None)
        else:
            candidate[pair] = value
        validate_decisions(self.items(), list(candidate.values()))
        self._commit([("revisions", pair, value)])

    def set_position(self, row_id: str, **fields: Any) -> None:
        """Save a dispute position against current evidence as one undoable decision."""
        if row_id not in self.index:
            raise ValueError("Unknown dispute schedule row")
        value = validate_schedule(
            self.items(),
            [
                {
                    "row_id": row_id,
                    "source_digest": schedule_digest(
                        self.items(), list(self.revisions.values()), row_id
                    ),
                    **fields,
                }
            ],
        )[0]
        self._commit([("schedule", row_id, value)])

    def clear_position(self, row_id: str) -> None:
        """Clear a position as an undoable command; regeneration preserves its audit trail."""
        if row_id not in self.index:
            raise ValueError("Unknown dispute schedule row")
        self._commit([("schedule", row_id, None)])

    def ledger_model(self) -> dict:
        """Preview full history and stale positions without changing accepted totals."""
        return build_ledger(self.items(), self.revision_model(), list(self.schedule.values()))

    def _commit(self, updates: list[tuple]) -> None:
        command = []
        for section, key, after in updates:
            before = getattr(self, section).get(key)
            if before != after:
                command.append((section, key, copy.deepcopy(before), copy.deepcopy(after)))
        if command:
            self._apply(command, 3)
            self.undo_stack.append(command)
            self.undo_stack = self.undo_stack[-100:]
            self.redo_stack.clear()

    def _apply(self, command: list[tuple], position: int) -> None:
        for change in command:
            mapping, key, value = getattr(self, change[0]), change[1], change[position]
            if value is None:
                mapping.pop(key, None)
            else:
                mapping[key] = copy.deepcopy(value)

    def undo(self) -> None:
        """Undo one accepted edit without changing the source analysis."""
        if self.undo_stack:
            command = self.undo_stack.pop()
            self._apply(command, 2)
            self.redo_stack.append(command)

    def redo(self) -> None:
        """Reapply an undone command."""
        if self.redo_stack:
            command = self.redo_stack.pop()
            self._apply(command, 3)
            self.undo_stack.append(command)

    def _validate_statuses(self, statuses: object) -> None:
        validate_review_statuses(set(self.index), statuses)

    def payload(self) -> dict:
        """Detach a compact draft; it never replaces or embeds original source data."""
        return copy.deepcopy(
            {
                "version": 1,
                "base_fingerprint": self.snapshot.input_fingerprint,
                "base_digest": self.base_digest,
                "decisions": [self.decisions[key] for key in sorted(self.decisions)],
                "statuses": self.statuses,
                "revisions": [self.revisions[key] for key in sorted(self.revisions)],
                "schedule": [self.schedule[key] for key in sorted(self.schedule)],
            }
        )

    @classmethod
    def restore(cls, snapshot: AnalysisSnapshot, payload: object) -> ReviewSession:
        """Reject stale, foreign, malformed or unsupported draft data before use."""
        session = cls(snapshot)
        if (
            not isinstance(payload, dict)
            or set(payload)
            not in (
                {
                    "version",
                    "base_fingerprint",
                    "base_digest",
                    "decisions",
                    "statuses",
                    "revisions",
                },
                {
                    "version",
                    "base_fingerprint",
                    "base_digest",
                    "decisions",
                    "statuses",
                    "revisions",
                    "schedule",
                },
            )
            or type(payload["version"]) is not int
            or payload["version"] != 1
            or payload["base_fingerprint"] != snapshot.input_fingerprint
            or payload["base_digest"] != session.base_digest
        ):
            raise ValueError("This draft belongs to a different or changed saved analysis")
        decisions = payload["decisions"]
        if not isinstance(decisions, list) or any(not isinstance(d, dict) for d in decisions):
            raise ValueError("Invalid draft decisions")
        ids = [d.get("row_id") for d in decisions]
        if any(not isinstance(key, str) or key not in session.index for key in ids) or len(
            set(ids)
        ) != len(ids):
            raise ValueError("Unknown or duplicate draft rows")
        for decision in decisions:
            apply_review(
                {
                    "version": 1,
                    "base_fingerprint": snapshot.input_fingerprint,
                    "rows": [session.index[decision["row_id"]]],
                    "decisions": [decision],
                }
            )
        session.decisions = {
            d["row_id"]: {
                **copy.deepcopy(d),
                "changes": {
                    field: input_value(field, value) for field, value in d["changes"].items()
                },
            }
            for d in decisions
        }
        session._validate_statuses(payload["statuses"])
        session.statuses = copy.deepcopy(payload["statuses"])
        session.revisions = {
            (d["older"], d["newer"]): d
            for d in validate_decisions(session.items(), payload["revisions"])
        }
        session.schedule = {
            d["row_id"]: d for d in validate_schedule(session.items(), payload.get("schedule", []))
        }
        return session

    def save(self, path: str | Path) -> None:
        """Atomically save review progress separately from the immutable result."""
        payload = self.payload()
        # Match write_json's indentation and ASCII escaping, so every saved
        # draft also fits the bounded reader (including non-ASCII reasons).
        if len(json.dumps(payload, indent=2, allow_nan=False).encode()) > MAX_DRAFT_BYTES:
            raise ValueError("Review draft exceeds the 8 MiB limit")
        write_json(payload, str(path))

    @classmethod
    def load(cls, snapshot: AnalysisSnapshot, path: str | Path) -> ReviewSession:
        """Bound and strictly decode a saved human-review draft."""
        with Path(path).open("rb") as source:
            raw = source.read(MAX_DRAFT_BYTES + 1)
        if len(raw) > MAX_DRAFT_BYTES:
            raise ValueError("Review draft exceeds the 8 MiB limit")
        return cls.restore(snapshot, strict_loads(raw.decode("utf-8")))

    def inputs(self, input_path: str) -> tuple[list[dict], dict, str]:
        """Use the same complete decision set and audit builder as Excel."""
        return review_inputs(
            self.snapshot,
            [self.decision(row["id"]) for row in self.rows],
            input_path=input_path,
            origin="reviewed_in_app",
            statuses=self.statuses,
            revisions=[self.revisions[key] for key in sorted(self.revisions)],
            schedule=[self.schedule[key] for key in sorted(self.schedule)],
        )

    def revision_model(self) -> dict:
        """Suggest relationships for the draft without changing financial analysis."""
        return build_revisions(
            self.items(), [self.revisions[key] for key in sorted(self.revisions)]
        )
