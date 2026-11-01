"""Scriptable access to the same record, revision, position and claim commands."""

from __future__ import annotations

import argparse
import json
import tempfile
from pathlib import Path

from edf_bill_fetcher.application.review_journal import ReviewJournal
from edf_bill_fetcher.application.review_session import ReviewSession
from edf_bill_fetcher.application.run_service import check_source_destinations
from edf_bill_fetcher.domain.json_values import object_value
from edf_bill_fetcher.helpers.artifacts import write_json
from edf_bill_fetcher.io.persistence.result_store import load_result


def main(argv: list[str]) -> None:
    """Apply one validated command and atomically preserve a portable review draft."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--result", required=True)
    parser.add_argument("--draft", required=True)
    parser.add_argument("--command", required=True, help="JSON file with operation and arguments")
    parser.add_argument("--publish", help="Create a fresh reviewed pack beneath this output folder")
    parser.add_argument("--preview", help="Save a recalculated preview result JSON")
    parser.add_argument(
        "--formats",
        nargs="+",
        choices=("xlsx", "pdf", "docx", "html"),
        default=["xlsx", "pdf", "docx", "html"],
    )
    options = parser.parse_args(argv)
    snapshot = load_result(options.result)
    draft = Path(options.draft)
    session = ReviewSession.load(snapshot, draft) if draft.exists() else ReviewSession(snapshot)
    command = json.loads(Path(options.command).read_text(encoding="utf-8"))
    if not isinstance(command, dict) or set(command) != {"operation", "arguments"}:
        raise ValueError("Supply operation and arguments")
    history_path = draft.with_name(draft.name + ".history.json")
    destinations = [
        draft.resolve(),
        history_path.resolve(),
        *([Path(options.preview).resolve()] if options.preview else []),
    ]
    if len(set(destinations)) != len(destinations) or set(destinations) & {
        Path(options.result).resolve(),
        Path(options.command).resolve(),
    }:
        raise ValueError("Review destination would overwrite an input")
    check_source_destinations(
        snapshot.prepared.extraction.request,
        destinations,
        object_value(snapshot.prepared.extraction.legacy_state),
    )
    history, position = [session.payload()], 0
    if history_path.is_file():
        previous = json.loads(history_path.read_text(encoding="utf-8"))
        journal = ReviewJournal(snapshot, previous["history"], previous["position"])
        if journal.session.payload() == session.payload():
            history, position = journal.history, journal.position
        elif command["operation"] in {"undo", "redo"}:
            raise ValueError("Review history no longer matches this draft")
    journal = ReviewJournal(snapshot, history, position)
    journal.apply(command["operation"], command["arguments"])
    session = journal.session
    session.save(draft)
    write_json({"history": journal.history, "position": journal.position}, str(history_path))
    if options.preview:
        from edf_bill_fetcher.application.review_execution import preview_review
        from edf_bill_fetcher.io.persistence.result_store import save_result

        save_result(
            preview_review(options.result, session.payload()),
            options.preview,
            result_id="review-preview",
        )
    if options.publish:
        from edf_bill_fetcher.application.review_execution import publish_review

        parent = Path(options.publish)
        parent.mkdir(parents=True, exist_ok=True)
        directory = tempfile.mkdtemp(prefix="reviewed_", dir=parent)
        outcome = publish_review(
            options.result, session.payload(), directory, formats=tuple(options.formats)
        )
        if outcome.status not in {"complete", "review_required", "empty"}:
            raise ValueError(outcome.error or outcome.status)
