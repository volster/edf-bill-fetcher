"""Read-only evidence views and explicit source relationships for the desktop."""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, replace
from pathlib import Path

from edf_bill_fetcher.application.run_service import check_cancelled
from edf_bill_fetcher.domain.analysis import AnalysisSnapshot
from edf_bill_fetcher.domain.json_values import FrozenJSON, freeze_json, object_value, rows_value
from edf_bill_fetcher.io.persistence.run_manifest import artifact_record
from edf_bill_fetcher.io.writers.evidence_bundle import resolve_bundle_paths
from edf_bill_fetcher.models.report_view import RenderSelection, build_report_view


@dataclass(frozen=True)
class ReviewSource:
    """A source occurrence; page numbers refer to the whole original PDF."""

    title: str
    path: str | None
    content_hash: str | None
    pages: tuple[int, int] | None
    status: str
    source_id: str
    invalid_pages: bool = False


@dataclass(frozen=True)
class ReviewEntry:
    """One preserved row or finding, with detached fields and source choices."""

    category: str
    fields: FrozenJSON
    sources: tuple[ReviewSource, ...]
    relation_note: str = ""


@dataclass(frozen=True)
class EvidenceReview:
    """A presentation of saved calculations with runtime evidence availability."""

    snapshot: AnalysisSnapshot
    directory: str
    entries: tuple[ReviewEntry, ...]
    notes: tuple[str, ...]


def text(value: object) -> str:
    """Display unavailable fields without inventing a zero or date."""
    if value is None:
        return ""
    if isinstance(value, (dict, list)):
        return (
            "; ".join(f"{key}: {item}" for key, item in value.items())
            if isinstance(value, dict)
            else "; ".join(map(str, value))
        )
    return str(value)


def page_range(row: dict) -> tuple[int, int] | None:
    """Accept only positive, ordered, integral source-page metadata."""
    start, end = row.get("Page From"), row.get("Page To")
    if (
        isinstance(start, (int, float))
        and not isinstance(start, bool)
        and isinstance(end, (int, float))
        and not isinstance(end, bool)
        and 0 < start <= end <= 2_147_483_647
        and (isinstance(start, int) or start.is_integer())
        and (isinstance(end, int) or end.is_integer())
    ):
        return int(start), int(end)
    return None


def build_evidence_review(
    snapshot: AnalysisSnapshot,
    directory: str,
    cancelled: Callable[[], bool] = lambda: False,
) -> EvidenceReview:
    """Resolve evidence and expose accepted models without preparing or analysing again."""
    check_cancelled(cancelled)
    state = object_value(snapshot.prepared.extraction.legacy_state)
    paths, warnings = resolve_bundle_paths(state, directory)
    check_cancelled(cancelled)
    manifest = state.get("source_manifest", {})
    manifest = manifest if isinstance(manifest, dict) else {}
    notes = list(warnings)
    errors = state.get("error_log", [])
    notes.extend(map(str, errors if isinstance(errors, list) else [errors]))
    statuses: dict[str, str] = {}
    # Sources deliberately not bundled can still be reviewed in their original
    # location. A missing/changed bundled copy must never fall back silently.
    for source_id, metadata in manifest.items():
        check_cancelled(cancelled)
        if not isinstance(metadata, dict):
            continue
        if source_id in paths:
            statuses[source_id] = (
                "Verified saved source"
                if metadata.get("Content Hash")
                else "Saved source · hash not verified at load"
            )
        elif metadata.get("Saved Path"):
            statuses[source_id] = "Saved source unavailable or changed"
        else:
            original, expected = metadata.get("Original Path"), metadata.get("Content Hash")
            if not isinstance(original, str) or not expected:
                statuses[source_id] = "Original source unavailable"
                continue
            try:
                if artifact_record(original)["sha256"] == expected:
                    paths[source_id] = str(Path(original).resolve())
                    statuses[source_id] = "Verified original source · not copied into this pack"
                else:
                    statuses[source_id] = "Original source changed"
            except (OSError, ValueError, RuntimeError):
                statuses[source_id] = "Original source unavailable"
            if source_id not in paths:
                notes.append(
                    f"{statuses[source_id]}: {metadata.get('Original Filename', source_id)}"
                )

    def sources(row: dict) -> tuple[ReviewSource, ...]:
        primary = text(row.get("Source ID"))
        extra = row.get("Source IDs", [])
        ids = [primary] if primary else []
        if isinstance(extra, list):
            ids.extend(value for value in extra if isinstance(value, str) and value)
        attachment = text(row.get("Attachment Name"))
        if attachment.strip().casefold() in {"n/a", "nan", "none", "null"}:
            attachment = ""
        if not ids and attachment:
            ids = [
                sid
                for sid, meta in manifest.items()
                if isinstance(meta, dict) and meta.get("Attachment Name") == attachment
            ]
            if len(ids) > 1:
                return (
                    ReviewSource(
                        attachment, None, None, page_range(row), "Ambiguous filename mapping", ""
                    ),
                )
        if not ids:
            if not attachment:
                return ()
            ids = [attachment]
        result = []
        pages = page_range(row)
        invalid_pages = pages is None and any(
            text(row.get(key)).strip().casefold() not in {"", "n/a", "nan", "none", "null"}
            for key in ("Page From", "Page To")
        )
        for source_id in dict.fromkeys(ids):
            metadata = manifest.get(source_id, {})
            metadata = metadata if isinstance(metadata, dict) else {}
            path = paths.get(source_id)
            # An authoritative identity must not be replaced by a stale name.
            if not manifest and path is None:
                path = paths.get(attachment)
            name = text(
                metadata.get("Original Filename")
                or row.get("Original Filename")
                or attachment
                or source_id
            )
            result.append(
                ReviewSource(
                    name,
                    path,
                    text(metadata.get("Content Hash") or row.get("Content Hash")) or None,
                    pages,
                    statuses.get(
                        source_id,
                        "Legacy source · integrity unverified at load"
                        if path
                        else "Source unavailable",
                    ),
                    source_id,
                    invalid_pages,
                )
            )
        return tuple(result)

    records = rows_value(snapshot.prepared.records)
    entries = []
    for row in records:
        check_cancelled(cancelled)
        kind = row.get("Entry Type")
        category = (
            "Payments & credits"
            if kind in {"Payment", "Credit", "Refund"}
            else "Bills & charges"
            if kind in {"New Bill", "Ongoing Balance", "Charge", "Late Payment"}
            else "Other records"
        )
        entries.append(ReviewEntry(category, freeze_json(row), sources(row)))
    view = build_report_view(snapshot, RenderSelection())
    canonical_entries = tuple(entries)
    for flag, when, amount, detail, severity in view.model("dispute").flags:
        # Saved legacy flags carry dates, not source/account identities. Keep
        # date matches explicit and require the user to choose their context.
        related: list[ReviewSource] = []
        for entry in canonical_entries:
            row = object_value(entry.fields)
            if text(when) and text(row.get("Date")) == text(when):
                prefix = " · ".join(
                    text(row.get(k))
                    for k in ("Invoice #", "Account", "Fuel", "Entry Type")
                    if text(row.get(k))
                )
                related.extend(
                    replace(source, title=prefix + " · " + source.title) for source in entry.sources
                )
        entries.append(
            ReviewEntry(
                "Findings",
                freeze_json(
                    {
                        "Finding": flag,
                        "Date": when,
                        "Amount (£)": amount,
                        "Details": detail,
                        "Severity": severity,
                    }
                ),
                tuple(related),
                "Related records match this date only; the saved finding does not identify a source.",
            )
        )
    for row in rows_value(snapshot.prepared.duplicates):
        entries.append(
            ReviewEntry(
                "Duplicates",
                freeze_json(row),
                sources(row),
                "Excluded duplicate; retained record index: " + text(row.get("_matches_kept_idx")),
            )
        )
    filtered = state.get("filtered_records", [])
    if isinstance(filtered, list):
        for row in filtered:
            if isinstance(row, dict):
                entries.append(
                    ReviewEntry(
                        "Filtered out",
                        freeze_json(row),
                        sources(row),
                        "Excluded by the captured extraction filters.",
                    )
                )
    check_cancelled(cancelled)
    for source_id, metadata in manifest.items():
        if isinstance(metadata, dict):
            row = {**metadata, "Source ID": source_id}
            entries.append(
                ReviewEntry(
                    "Source files",
                    freeze_json(row),
                    sources(row),
                    "Original file inventory. Structured SAP histories remain available in the exported workbook.",
                )
            )
    return EvidenceReview(
        snapshot,
        str(Path(directory).resolve()),
        tuple(entries),
        tuple(dict.fromkeys(note for note in notes if note)),
    )
