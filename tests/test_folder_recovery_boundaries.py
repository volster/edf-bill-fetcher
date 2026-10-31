"""Discovery failures must stay visible without losing readable sibling evidence."""

import errno
import os
from datetime import date, datetime, timezone
from email.message import EmailMessage
from unittest.mock import Mock

import pytest

from edf_bill_fetcher.application.run_service import run
from edf_bill_fetcher.collectors import engine as module
from edf_bill_fetcher.collectors.engine import EvidenceEngine
from edf_bill_fetcher.domain.json_values import freeze_json, object_value
from edf_bill_fetcher.domain.requests import RunRequest
from edf_bill_fetcher.io.persistence.result_store import load_result


@pytest.fixture
def collector():
    value = EvidenceEngine({"filter_below": False}, lambda _: None)
    yield value
    value.cleanup()


def scan(collector, folder, kind):
    if kind == "pdf":
        collector.crawl_local_pdfs(str(folder))
    else:
        collector._process_message_folder(str(folder), ".eml", module.parse_eml_message, kind)


def write_bill(folder):
    message = EmailMessage()
    message["Subject"] = "EDF bill"
    message["From"] = "billing@edfenergy.com"
    message.set_content("Total charges for this period £120.00 in debit")
    (folder / "good.eml").write_bytes(message.as_bytes())
    (folder / "good.pdf").write_bytes(b"PDF placeholder")


@pytest.mark.parametrize("kind", ["pdf", "EML", "MSG"])
@pytest.mark.parametrize("deny_root", [False, True])
def test_unreadable_directory_is_reported_and_siblings_survive(
    tmp_path, monkeypatch, collector, kind, deny_root
):
    write_bill(tmp_path)
    denied = tmp_path if deny_root else tmp_path / "denied"
    denied.mkdir(exist_ok=True)
    original = os.scandir

    def scandir(path):
        if os.fspath(path) == str(denied):
            raise PermissionError(errno.EACCES, "Access denied", str(denied))
        return original(path)

    monkeypatch.setattr(module.os, "scandir", scandir)
    process = Mock()
    monkeypatch.setattr(collector, "process_pdf_file", process)
    scan(collector, tmp_path, kind)
    assert len(collector.error_log) == 1
    assert str(denied) in collector.error_log[0]
    assert "Access denied" in collector.error_log[0]
    if kind == "pdf":
        assert process.call_count == (0 if deny_root else 1)
    else:
        assert len(collector.records) == (0 if deny_root else 1)
        if not deny_root:
            assert collector.records[0]["Period Charge (£)"] == 120


@pytest.mark.parametrize("kind", ["pdf", "EML", "MSG"])
def test_cancel_during_discovery_stops_walking_before_processing(
    tmp_path, monkeypatch, collector, kind
):
    write_bill(tmp_path)
    visited = []

    def walk(path, **kwargs):
        visited.append("root")
        collector.cancel_event.set()
        yield str(tmp_path), ["later"], ["good.pdf", "good.eml"]
        visited.append("later")
        yield str(tmp_path / "later"), [], []

    monkeypatch.setattr(module.os, "walk", walk)
    process = Mock()
    monkeypatch.setattr(collector, "process_pdf_file", process)
    scan(collector, tmp_path, kind)
    assert visited == ["root"]
    process.assert_not_called()
    assert collector.records == []
    assert collector.error_log == []


@pytest.mark.parametrize("failure", [FileNotFoundError, PermissionError])
def test_unavailable_pdf_metadata_does_not_abort_later_files(
    tmp_path, monkeypatch, collector, failure
):
    missing = tmp_path / "a-unavailable.pdf"
    good = tmp_path / "b-readable.pdf"
    missing.touch()
    good.touch()
    original = os.path.getmtime

    def getmtime(path):
        if os.fspath(path) == str(missing):
            raise failure("PDF became unavailable")
        return original(path)

    monkeypatch.setattr(module.os.path, "getmtime", getmtime)
    process = Mock()
    monkeypatch.setattr(collector, "process_pdf_file", process)
    collector.crawl_local_pdfs(str(tmp_path))
    assert process.call_count == 1
    assert process.call_args.args[0] == str(good)
    assert collector.pdf_count == 1
    assert len(collector.error_log) == 1
    assert str(missing) in collector.error_log[0]
    assert "PDF became unavailable" in collector.error_log[0]


def test_partial_folder_run_persists_warning_and_marks_review_required(
    tmp_path, monkeypatch, collector
):
    sources = tmp_path / "sources"
    sources.mkdir()
    write_bill(sources)
    denied = sources / "denied"
    denied.mkdir()
    original = os.scandir

    def scandir(path):
        if os.fspath(path) == str(denied):
            raise PermissionError(errno.EACCES, "Access denied", str(denied))
        return original(path)

    monkeypatch.setattr(module.os, "scandir", scandir)
    request = RunRequest(
        "partial-folder",
        datetime(2025, 3, 1, tzinfo=timezone.utc),
        date(2025, 3, 1),
        (),
        freeze_json({"as_of": "2025-03-01", "analysis_min": 0}),
    )

    def collect():
        collector.process_eml_folder(str(sources))
        return collector.records, collector.__getstate__()

    outcome = run(
        request,
        {},
        result_path=str(tmp_path / "result.json"),
        manifest_path=str(tmp_path / "run.json"),
        collect=collect,
    )
    assert outcome.status == "review_required", outcome.error
    assert object_value(outcome.manifest)["status"] == "review_required"
    restored = load_result(tmp_path / "result.json")
    state = object_value(restored.prepared.extraction.legacy_state)
    assert any(str(denied) in warning for warning in state["error_log"])
    assert len(restored.prepared.extraction.observations) == 1
