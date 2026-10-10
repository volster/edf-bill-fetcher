"""Validated run inputs captured independently of presentation and live UI state."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta

from edf_bill_fetcher.domain.json_values import FrozenJSON, object_value

SOURCE_KINDS = frozenset({"pdf_dir", "htm", "pst", "ost", "eml_dir", "msg_dir", "records_json"})


@dataclass(frozen=True)
class SourceSelection:
    """One captured source path; filesystem access belongs to an I/O adapter."""

    kind: str
    path: str

    def __post_init__(self) -> None:
        """Reject unsupported source kinds and empty paths."""
        if self.kind not in SOURCE_KINDS or not isinstance(self.path, str) or not self.path.strip():
            raise ValueError("Invalid source selection")


@dataclass(frozen=True)
class RunRequest:
    """Pinned analytical request with immutable configuration."""

    run_id: str
    created_at: datetime
    as_of: date
    sources: tuple[SourceSelection, ...]
    config: FrozenJSON

    def __post_init__(self) -> None:
        """Validate identifiers, UTC timestamps, pinned date and unique sources."""
        if not isinstance(self.run_id, str) or not self.run_id.strip():
            raise ValueError("Run ID must be nonempty")
        if not isinstance(self.created_at, datetime) or self.created_at.utcoffset() != timedelta(0):
            raise ValueError("Creation timestamp must be aware UTC")
        if type(self.as_of) is not date:
            raise ValueError("Valuation date must be a calendar date")
        if type(self.sources) is not tuple or any(
            not isinstance(source, SourceSelection) for source in self.sources
        ):
            raise ValueError("Sources must be immutable selections")
        keys = [(source.kind, source.path) for source in self.sources]
        if len(keys) != len(set(keys)):
            raise ValueError("Duplicate source selection")
        if object_value(self.config).get("as_of") != self.as_of.isoformat():
            raise ValueError("Configuration valuation date disagrees with request")
