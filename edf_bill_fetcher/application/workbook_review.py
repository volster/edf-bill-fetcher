"""Regenerate a new evidence pack from validated Excel review decisions."""

from __future__ import annotations

import os
import tempfile
from datetime import date, datetime, timezone
from pathlib import Path
from uuid import uuid4

import pandas as pd

from edf_bill_fetcher.application.run_service import RunOutcome, run
from edf_bill_fetcher.domain.evidence import PreparedSnapshot
from edf_bill_fetcher.domain.json_values import freeze_json, object_value, rows_value
from edf_bill_fetcher.domain.requests import RunRequest, SourceSelection
from edf_bill_fetcher.io.persistence.run_manifest import artifact_record
from edf_bill_fetcher.io.review_workbook import load_review_workbook
from edf_bill_fetcher.io.writers.evidence_bundle import (
    build_bundle_index,
    build_bundle_manifest,
    resolve_bundle_paths,
    save_evidence_files,
)
from edf_bill_fetcher.models.report_view import RenderSelection


def _identities(records: list[dict]) -> set[str]:
    result: set[str] = set()
    for record in records:
        identities = record.get("Source IDs", [])
        if isinstance(identities, list):
            result.update(identity for identity in identities if isinstance(identity, str))
        identity = record.get("Source ID")
        if isinstance(identity, str) and identity:
            result.add(identity)
    return result


def bundle_review_sources(prepared: PreparedSnapshot, directory: str, workbook: str) -> dict:
    """Copy verified retained sources; exclude files used only by excluded records."""
    state = object_value(prepared.extraction.legacy_state)
    manifest = {
        **state.get("review_excluded_source_manifest", {}),
        **state.get("source_manifest", {}),
    }
    records = (
        rows_value(prepared.records)
        + rows_value(prepared.duplicates)
        + state.get("filtered_records", [])
    )
    excluded_ids = _identities(state.get("review_excluded_records", [])) - _identities(records)
    # Older sources have attachment names but no occurrence IDs. Only exclude a
    # name when no retained row references it; never guess across ambiguous IDs.
    retained_names = {row.get("Attachment Name") for row in records}
    excluded_names = {
        row.get("Attachment Name") for row in state.get("review_excluded_records", [])
    } - retained_names
    selected = {
        identity: metadata
        for identity, metadata in manifest.items()
        if identity not in excluded_ids
        and not (
            not metadata.get("Source ID") and metadata.get("Attachment Name") in excluded_names
        )
    }
    paths: dict[str, str] = {}
    roots = [str(Path(workbook).resolve().parent)]
    old_root = state.get("source_bundle_path")
    if isinstance(old_root, str) and old_root and old_root not in roots:
        roots.append(old_root)
    for root in roots:
        resolved, _ = resolve_bundle_paths({**state, "source_manifest": manifest}, root)
        for identity, path in resolved.items():
            paths.setdefault(identity, path)
    for identity, metadata in selected.items():
        # A source deliberately not bundled may still be available at its
        # recorded original location. Never use this for a missing saved copy.
        original, expected = metadata.get("Original Path"), metadata.get("Content Hash")
        if (
            identity in paths
            or metadata.get("Saved Path")
            or not isinstance(original, str)
            or not expected
        ):
            continue
        try:
            if artifact_record(original)["sha256"] == expected:
                paths[identity] = original
        except (OSError, ValueError, RuntimeError):
            continue
    # Excluded files stay in earlier packs. A reinstatement can recover the
    # exact earlier copy while that pack exists, with its content hash checked.
    for identity, path in state.get("review_source_paths", {}).items():
        expected = manifest.get(identity, {}).get("Content Hash")
        try:
            if identity not in paths and expected and artifact_record(path)["sha256"] == expected:
                paths[identity] = path
        except (OSError, ValueError, TypeError):
            continue
    bundle_records = pd.DataFrame(
        records + [{**metadata, "Source ID": identity} for identity, metadata in selected.items()]
    )
    notes: list[str] = []
    saved = save_evidence_files(
        bundle_records, paths, str(Path(directory) / "evidence_files"), log=notes.append
    )
    build_bundle_manifest(bundle_records, saved, directory)
    build_bundle_index(
        bundle_records,
        saved,
        str(Path(directory) / "evidence_index.docx"),
        account=str(object_value(prepared.effective_config).get("acc_num", "")),
    )
    aliases = dict(saved)
    excluded_manifest = {}
    for identity, metadata in manifest.items():
        copied = saved.get(metadata.get("Attachment Name")) if identity in selected else None
        if identity not in selected:
            excluded_manifest[identity] = metadata
            continue
        metadata["Saved Path"] = (
            os.path.relpath(copied, directory).replace(os.sep, "/") if copied else None
        )
        if copied:
            aliases[identity] = copied
    # JSON-only legacy evidence may not have a source manifest. The same copy
    # helper reports unavailable originals rather than manufacturing provenance.
    state.update(
        {
            "source_manifest": selected,
            "review_excluded_source_manifest": excluded_manifest,
            "review_source_paths": {
                identity: path for identity, path in paths.items() if identity in manifest
            },
            "saved_evidence_paths": aliases,
            "source_bundle_path": str(Path(directory).resolve()),
            "source_paths": paths,
        }
    )
    state["error_log"] = [*state.get("error_log", []), *notes]
    return state


def regenerate_workbook(
    workbook: str,
    output_dir: str,
    *,
    formats: tuple[str, ...] = ("xlsx", "pdf", "docx", "html"),
    sections: tuple[str, ...] | None = None,
) -> tuple[str, RunOutcome]:
    """Validate before reserving a fresh run folder, then recompute all selected outputs."""
    imported = load_review_workbook(workbook)
    if imported is None:
        raise ValueError(
            "This workbook predates the review workflow. Export a new review workbook from its saved result.json first."
        )
    records, state, as_of = imported
    config = state["review_config"]
    request = RunRequest(
        uuid4().hex,
        datetime.now(timezone.utc),
        date.fromisoformat(as_of),
        (SourceSelection("records_json", str(Path(workbook).resolve())),),
        freeze_json(config),
    )
    # Always include a new editable workbook for the next review round.
    formats = ("xlsx", *(fmt for fmt in formats if fmt != "xlsx"))
    selection = RenderSelection(formats, sections)
    root = Path(output_dir)
    root.mkdir(parents=True, exist_ok=True)
    directory = tempfile.mkdtemp(prefix="reviewed_", dir=root)
    outputs = {
        fmt: str(
            Path(directory)
            / ("Reviewed_Evidence" + ("_Report" if fmt != "xlsx" else "") + "." + fmt)
        )
        for fmt in formats
    }
    outcome = run(
        request,
        outputs,
        result_path=str(Path(directory) / "result.json"),
        manifest_path=str(Path(directory) / "run.json"),
        collect=lambda: (records, state),
        copy_sources=lambda prepared: bundle_review_sources(prepared, directory, workbook),
        selection=selection,
    )
    return directory, outcome
