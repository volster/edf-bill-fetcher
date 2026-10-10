"""Detached evidence observations and immutable preparation results."""

from __future__ import annotations

import re
from dataclasses import dataclass

from edf_bill_fetcher.domain.json_values import FrozenJSON, object_value, rows_value
from edf_bill_fetcher.domain.requests import RunRequest


@dataclass(frozen=True)
class Observation:
    """An original row with a recorded extraction-local identity."""

    observation_id: str
    ordinal: int
    raw: FrozenJSON

    def __post_init__(self) -> None:
        """Validate the observation identity and raw object shape."""
        if not isinstance(self.observation_id, str) or not re.fullmatch(
            r"[0-9a-f]{64}", self.observation_id
        ):
            raise ValueError("Invalid observation identity")
        if type(self.ordinal) is not int or self.ordinal < 0:
            raise ValueError("Observation ordinal must be nonnegative")
        object_value(self.raw)


@dataclass(frozen=True)
class ExtractionSnapshot:
    """Preserved originals and data-only legacy metadata for one captured request."""

    request: RunRequest
    observations: tuple[Observation, ...]
    legacy_state: FrozenJSON

    def __post_init__(self) -> None:
        """Reject mutable membership and duplicate observation identities/ordinals."""
        if type(self.observations) is not tuple or any(
            not isinstance(row, Observation) for row in self.observations
        ):
            raise ValueError("Observations must be immutable")
        ids = [row.observation_id for row in self.observations]
        ordinals = [row.ordinal for row in self.observations]
        if len(ids) != len(set(ids)) or ordinals != list(range(len(ordinals))):
            raise ValueError("Observation membership is duplicated or out of order")
        object_value(self.legacy_state)


@dataclass(frozen=True)
class PreparedSnapshot:
    """Canonical records and retained duplicate decisions without mutable frames."""

    extraction: ExtractionSnapshot
    records: FrozenJSON
    duplicates: FrozenJSON
    effective_config: FrozenJSON

    def __post_init__(self) -> None:
        """Validate row shapes, duplicate links and the captured valuation date."""
        records = rows_value(self.records)
        for duplicate in rows_value(self.duplicates):
            index = duplicate.get("_matches_kept_idx")
            if type(index) is not int or not 0 <= index < len(records):
                raise ValueError("Duplicate does not reference a kept record")
        if (
            object_value(self.effective_config).get("as_of")
            != self.extraction.request.as_of.isoformat()
        ):
            raise ValueError("Prepared configuration changed the valuation date")
