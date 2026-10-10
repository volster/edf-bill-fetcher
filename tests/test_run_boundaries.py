"""Failure and cancellation contracts at external run and archive boundaries."""

import hashlib
import json
from collections.abc import Callable
from dataclasses import replace
from datetime import date, datetime, timezone
from pathlib import Path
from threading import Event
from types import SimpleNamespace
from typing import Any
from unittest.mock import MagicMock

import pytest

from edf_bill_fetcher.application.run_service import RunCancelled, collect_engine, run
from edf_bill_fetcher.domain.json_values import freeze_json, object_value
from edf_bill_fetcher.domain.requests import RunRequest, SourceSelection
from edf_bill_fetcher.io.persistence.result_store import load_result
from edf_bill_fetcher.io.persistence.run_manifest import (
    artifact_record,
    load_manifest,
    save_manifest,
)
from edf_bill_fetcher.models.report_view import RenderSelection


@pytest.fixture
def captured_request() -> RunRequest:
    return RunRequest(
        "boundaries",
        datetime(2024, 4, 1, tzinfo=timezone.utc),
        date(2024, 4, 1),
        (),
        freeze_json({"as_of": "2024-04-01", "analysis_min": 0}),
    )


def records():
    return [
        {
            "Date": "31/01/2024",
            "Account": "A-TEST",
            "Fuel": "Electricity",
            "Entry Type": "New Bill",
            "Amount (£)": 250,
            "Period Charge (£)": 100,
        }
    ]


def render(view, path):
    Path(path).write_bytes(b"complete artifact")


@pytest.mark.parametrize(
    "stage", ["collecting", "copying_sources", "saving_result", "rendering_html"]
)
def test_stage_failure_records_where_it_stopped_and_always_cleans_up(
    captured_request: RunRequest, tmp_path: Path, stage: str
) -> None:
    cleanup = MagicMock()

    def fail(*args, **kwargs):
        persisted = json.loads((tmp_path / "run.json").read_text())
        assert persisted["stage"] == stage
        assert persisted["status"] == "running"
        raise OSError("injected " + stage)

    kwargs: dict[str, Any] = {"collect": lambda: (records(), {}), "renderers": {"html": render}}
    if stage == "collecting":
        kwargs["collect"] = fail
    elif stage == "copying_sources":
        kwargs["copy_sources"] = fail
    elif stage == "saving_result":
        kwargs["store"] = fail
    else:
        kwargs["renderers"] = {"html": fail}
    outcome = run(
        captured_request,
        {"html": str(tmp_path / "report.html")},
        result_path=str(tmp_path / "result.json"),
        manifest_path=str(tmp_path / "run.json"),
        cleanup=cleanup,
        **kwargs,
    )
    manifest = load_manifest(str(tmp_path / "run.json"))
    assert outcome.status == manifest["status"] == "failed"
    assert manifest["stage"] == stage
    assert "injected " + stage in manifest["error"]
    assert manifest["artifacts"] == {}
    if stage == "rendering_html":
        assert load_result(tmp_path / "result.json") == outcome.snapshot
        assert manifest["result"] == artifact_record(str(tmp_path / "result.json"))
    else:
        assert manifest["result"] is None
        assert not (tmp_path / "result.json").exists()
    cleanup.assert_called_once()


@pytest.mark.parametrize("when", ["before_collection", "after_collection", "after_first_render"])
def test_cancel_boundaries_preserve_only_published_artifacts(
    captured_request: RunRequest, tmp_path: Path, when: str
) -> None:
    cancelled = Event()
    cleanup = MagicMock()
    collect = MagicMock(return_value=(records(), {}))
    if when == "before_collection":
        cancelled.set()
    elif when == "after_collection":

        def collect_and_cancel():
            cancelled.set()
            return records(), {}

        collect.side_effect = collect_and_cancel

    def render_and_cancel(view, path):
        render(view, path)
        cancelled.set()

    outcome = run(
        captured_request,
        {"html": str(tmp_path / "first.html"), "pdf": str(tmp_path / "second.pdf")},
        result_path=str(tmp_path / "result.json"),
        manifest_path=str(tmp_path / "run.json"),
        collect=collect,
        cancelled=cancelled.is_set,
        cleanup=cleanup,
        renderers={"html": render_and_cancel, "pdf": render},
    )
    manifest = load_manifest(str(tmp_path / "run.json"))
    assert outcome.status == manifest["status"] == "cancelled"
    assert not (tmp_path / "second.pdf").exists()
    if when == "after_first_render":
        assert set(manifest["artifacts"]) == {"html"}
        assert load_result(tmp_path / "result.json") == outcome.snapshot
        assert (
            manifest["artifacts"]["html"]["sha256"]
            == hashlib.sha256(b"complete artifact").hexdigest()
        )
    else:
        assert manifest["artifacts"] == {}
        assert manifest["result"] is None
    if when == "before_collection":
        collect.assert_not_called()
    else:
        collect.assert_called_once()
    cleanup.assert_called_once()


def test_cleanup_failure_cannot_claim_success(captured_request: RunRequest, tmp_path: Path) -> None:
    cleanup = MagicMock(side_effect=OSError("resource still open"))
    outcome = run(
        captured_request,
        {"html": str(tmp_path / "report.html")},
        result_path=str(tmp_path / "result.json"),
        manifest_path=str(tmp_path / "run.json"),
        collect=lambda: (records(), {}),
        renderers={"html": render},
        cleanup=cleanup,
    )
    assert outcome.status == "failed"
    manifest = load_manifest(str(tmp_path / "run.json"))
    assert "Cleanup failed: resource still open" in manifest["error"]
    assert set(manifest["artifacts"]) == {"html"}
    assert load_result(tmp_path / "result.json") == outcome.snapshot


@pytest.mark.parametrize("invalid", ["no_collector", "unknown_format", "selection_disagrees"])
def test_invalid_run_aborts_before_collection(
    captured_request: RunRequest, tmp_path: Path, invalid: str
) -> None:
    collect = MagicMock(return_value=(records(), {}))
    output = {"html": str(tmp_path / "report.html")}
    selection = None
    collector: Callable | None = collect
    if invalid == "no_collector":
        collector = None
    elif invalid == "unknown_format":
        output = {"csv": str(tmp_path / "report.csv")}
    else:
        selection = RenderSelection(("pdf",))
    outcome = run(
        captured_request,
        output,
        result_path=str(tmp_path / "result.json"),
        manifest_path=str(tmp_path / "run.json"),
        collect=collector,
        selection=selection,
    )
    assert outcome.status == "failed"
    collect.assert_not_called()
    assert not (tmp_path / "result.json").exists()
    assert object_value(outcome.manifest)["artifacts"] == {}


@pytest.mark.parametrize("failure", ["open", "crawl"])
def test_archive_is_closed_when_ingestion_fails(captured_request: RunRequest, failure: str) -> None:
    archive = MagicMock()
    engine = SimpleNamespace(cancel_event=Event(), records=[], crawl_pst=MagicMock())
    if failure == "open":
        archive.open.side_effect = OSError("archive unreadable")
    else:
        engine.crawl_pst.side_effect = OSError("folder unreadable")
    captured_request = replace(captured_request, sources=(SourceSelection("pst", "synthetic.pst"),))
    with pytest.raises(OSError, match="unreadable"):
        collect_engine(engine, captured_request, pff_module=SimpleNamespace(file=lambda: archive))
    archive.close.assert_called_once()
    if failure == "open":
        engine.crawl_pst.assert_not_called()
    else:
        engine.crawl_pst.assert_called_once()


def test_collection_checks_cancel_before_opening_archive(captured_request: RunRequest) -> None:
    event = Event()
    event.set()
    constructor = MagicMock()
    captured_request = replace(captured_request, sources=(SourceSelection("ost", "synthetic.ost"),))
    with pytest.raises(RunCancelled):
        collect_engine(
            SimpleNamespace(cancel_event=event),
            captured_request,
            pff_module=SimpleNamespace(file=constructor),
        )
    constructor.assert_not_called()


def test_unimplemented_source_cannot_silently_be_ignored(captured_request: RunRequest) -> None:
    captured_request = replace(
        captured_request, sources=(SourceSelection("records_json", "synthetic.json"),)
    )
    with pytest.raises(ValueError, match="Unsupported collection source"):
        collect_engine(SimpleNamespace(cancel_event=Event()), captured_request)


@pytest.mark.parametrize("status", ["running", "captured", "interrupted"])
def test_nonterminal_manifest_is_recognised_as_interrupted(tmp_path: Path, status: str) -> None:
    path = tmp_path / "run.json"
    saved = {
        "manifest_version": 1,
        "status": status,
        "stage": "rendering_pdf",
        "artifacts": {"html": {"sha256": "kept"}},
    }
    save_manifest(saved, str(path))
    original = path.read_bytes()
    restored = load_manifest(str(path))
    assert restored["status"] == "interrupted"
    assert restored["stage"] == "rendering_pdf"
    assert restored["artifacts"] == saved["artifacts"]
    assert path.read_bytes() == original


@pytest.mark.parametrize("payload", [[], {"manifest_version": 2}])
def test_unsupported_manifest_cannot_be_accepted(tmp_path: Path, payload: object) -> None:
    path = tmp_path / "run.json"
    path.write_text(json.dumps(payload))
    with pytest.raises(ValueError, match="Unsupported run manifest"):
        load_manifest(str(path))


@pytest.mark.parametrize("status", ["complete", "review_required"])
@pytest.mark.parametrize("problem", ["missing_artifact", "error"])
def test_success_manifest_requires_all_artifacts_and_no_error(
    tmp_path: Path, status: str, problem: str
) -> None:
    path = tmp_path / "run.json"
    path.write_bytes(b"previous manifest")
    manifest: dict = {
        "status": status,
        "requested": {"html": "report.html"},
        "artifacts": {} if problem == "missing_artifact" else {"html": {}},
        "error": "renderer failed" if problem == "error" else None,
    }
    with pytest.raises(ValueError, match="cannot be complete"):
        save_manifest(manifest, str(path))
    assert path.read_bytes() == b"previous manifest"
    assert list(tmp_path.iterdir()) == [path]
