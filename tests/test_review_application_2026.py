"""Protect evidence originals and honest presentation across application boundaries."""

import hashlib
import json
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any
from unittest.mock import MagicMock

import pytest

from edf_bill_fetcher.application.analysis_service import analyse_prepared
from edf_bill_fetcher.application.evidence_review import build_evidence_review
from edf_bill_fetcher.application.legacy_adapter import capture_legacy
from edf_bill_fetcher.application.preparation_service import prepare_snapshot
from edf_bill_fetcher.application.run_service import run
from edf_bill_fetcher.domain.json_values import freeze_json
from edf_bill_fetcher.domain.requests import RunRequest, SourceSelection
from edf_bill_fetcher.io.cli import run_cli_extract, run_cli_pdf_report
from edf_bill_fetcher.io.persistence.result_store import save_result
from edf_bill_fetcher.models.report_view import RenderSelection, build_report_view


def request(sources=()):
    return RunRequest(
        "review-boundaries",
        datetime(2024, 4, 1, tzinfo=timezone.utc),
        date(2024, 4, 1),
        sources,
        freeze_json({"as_of": "2024-04-01", "analysis_min": 0}),
    )


def bill(**extra):
    return {
        "Date": "01/01/2024",
        "Account": "A",
        "Entry Type": "New Bill",
        "Amount (£)": 500,
        "Period Charge (£)": 100,
        "Invoice #": "B1",
        **extra,
    }


def snapshot(records, state=None, captured=None):
    return analyse_prepared(
        prepare_snapshot(capture_legacy(records, state or {}, captured or request()))
    )


@pytest.mark.parametrize("destination", ["report", "result", "manifest"])
def test_original_pdf_inside_selected_folder_cannot_be_overwritten(tmp_path, destination):
    original = tmp_path / "original.pdf"
    original.write_bytes(b"original evidence")
    paths = {
        "report": tmp_path / "report.pdf",
        "result": tmp_path / "result.json",
        "manifest": tmp_path / "run.json",
    }
    paths[destination] = original
    collector = MagicMock(return_value=([bill()], {}))
    outcome = run(
        request((SourceSelection("pdf_dir", str(tmp_path)),)),
        {"pdf": str(paths["report"])},
        result_path=str(paths["result"]),
        manifest_path=str(paths["manifest"]),
        collect=collector,
    )
    assert outcome.status == "failed"
    assert outcome.error
    assert "overwrite" in outcome.error
    collector.assert_not_called()
    assert original.read_bytes() == b"original evidence"


@pytest.mark.parametrize("mapping", ["source_paths", "saved_evidence_paths", "source_manifest"])
def test_saved_snapshot_protects_originals_outside_selected_paths(tmp_path, mapping):
    original = tmp_path / "original.pdf"
    original.write_bytes(b"preserved evidence")
    state: dict[str, Any]
    if mapping == "source_manifest":
        state = {mapping: {"sid": {"Original Path": str(original)}}}
    else:
        state = {mapping: {"sid": str(original)}}
    renderer = MagicMock()
    outcome = run(
        request(),
        {"pdf": str(original)},
        result_path=str(tmp_path / "result.json"),
        manifest_path=str(tmp_path / "run.json"),
        accepted_result=snapshot([bill()], state),
        renderers={"pdf": renderer},
    )
    assert outcome.status == "failed"
    renderer.assert_not_called()
    assert not (tmp_path / "result.json").exists()
    assert not (tmp_path / "run.json").exists()
    assert original.read_bytes() == b"preserved evidence"


@pytest.mark.parametrize("destination", ["report", "result", "manifest"])
def test_discovered_source_protection_runs_before_copy_or_render(tmp_path, destination):
    original = tmp_path / "original.pdf"
    original.write_bytes(b"source discovered during collection")
    copier = MagicMock()
    paths = {
        "report": tmp_path / "report.pdf",
        "result": tmp_path / "result.json",
        "manifest": tmp_path / "run.json",
    }
    paths[destination] = original
    outcome = run(
        request(),
        {"pdf": str(paths["report"])},
        result_path=str(paths["result"]),
        manifest_path=str(paths["manifest"]),
        collect=lambda: ([bill()], {"source_paths": {"sid": str(original)}}),
        copy_sources=copier,
    )
    assert outcome.status == "failed"
    copier.assert_not_called()
    assert original.read_bytes() == b"source discovered during collection"


def test_failure_to_save_run_status_retains_available_snapshot(tmp_path, monkeypatch):
    def fail_save(*args):
        raise OSError("disk full")

    monkeypatch.setattr("edf_bill_fetcher.application.run_service.save_manifest", fail_save)
    accepted = snapshot([bill()])
    cleanup = MagicMock()
    outcome = run(
        request(),
        {},
        result_path=str(tmp_path / "result.json"),
        manifest_path=str(tmp_path / "run.json"),
        accepted_result=accepted,
        cleanup=cleanup,
    )
    assert outcome.status == "failed"
    assert outcome.snapshot == accepted
    assert outcome.error
    assert "Run status could not be saved: disk full" in outcome.error
    cleanup.assert_called_once()


def test_observation_timeline_does_not_claim_payments_or_missing_bills():
    view = build_report_view(
        snapshot(
            [
                bill(),
                bill(Date="15/03/2024", **{"Invoice #": "B2", "Amount (£)": 100}),
            ]
        ),
        RenderSelection(),
    )
    timeline = view.model("sap")["timeline"]
    labels = {event[1] for event in timeline}
    assert {
        "FIRST OBSERVED BALANCE",
        "LATEST OBSERVED BALANCE",
        "BALANCE REDUCTION",
        "EVIDENCE GAP",
    } <= labels
    assert "PAYMENT/CREDIT" not in labels
    assert "ACCOUNT START" not in labels
    assert "BILLING GAP" not in labels
    assert "does not establish whether bills were issued" in next(
        event[2] for event in timeline if event[1] == "EVIDENCE GAP"
    )


def test_duplicate_invoice_sources_keep_their_own_page_ranges(tmp_path):
    files = [tmp_path / "first.pdf", tmp_path / "second.pdf"]
    for index, path in enumerate(files):
        path.write_bytes(f"distinct original {index}".encode())
    hashes = [hashlib.sha256(path.read_bytes()).hexdigest() for path in files]
    manifest = {
        sid: {
            "Attachment Name": path.name,
            "Original Filename": path.name,
            "Saved Path": path.name,
            "Content Hash": digest,
        }
        for sid, path, digest in zip(("sid1", "sid2"), files, hashes, strict=True)
    }
    preserved = snapshot(
        [
            bill(
                **{
                    "Source": "Local PDF Folder",
                    "Source ID": "sid1",
                    "Content Hash": hashes[0],
                    "Attachment Name": files[0].name,
                    "Page From": 1,
                    "Page To": 2,
                }
            ),
            bill(
                **{
                    "Source": "EML PDF Attachment",
                    "Source ID": "sid2",
                    "Content Hash": hashes[1],
                    "Attachment Name": files[1].name,
                    "Page From": 7,
                    "Page To": 8,
                }
            ),
        ],
        {"source_manifest": manifest},
    )
    review = build_evidence_review(preserved, str(tmp_path))
    entry = next(row for row in review.entries if row.category == "Bills & charges")
    assert {source.source_id: source.pages for source in entry.sources} == {
        "sid1": (1, 2),
        "sid2": (7, 8),
    }


def test_legacy_other_original_with_unknown_pages_does_not_inherit_primary_range(tmp_path):
    first, second = tmp_path / "first.pdf", tmp_path / "second.pdf"
    first.write_bytes(b"one")
    second.write_bytes(b"different PDF")
    hashes = {
        "sid1": hashlib.sha256(first.read_bytes()).hexdigest(),
        "sid2": hashlib.sha256(second.read_bytes()).hexdigest(),
    }
    manifest = {
        sid: {"Attachment Name": path.name, "Saved Path": path.name, "Content Hash": hashes[sid]}
        for sid, path in (("sid1", first), ("sid2", second))
    }
    preserved = snapshot(
        [
            bill(
                **{
                    "Source ID": "sid1",
                    "Source IDs": ["sid1", "sid2"],
                    "Page From": 1,
                    "Page To": 2,
                    "Content Hash": hashes["sid1"],
                }
            )
        ],
        {"source_manifest": manifest},
    )
    entry = next(
        row
        for row in build_evidence_review(preserved, str(tmp_path)).entries
        if row.category == "Bills & charges"
    )
    assert {source.source_id: source.pages for source in entry.sources} == {
        "sid1": (1, 2),
        "sid2": None,
    }


def test_missing_authoritative_id_cannot_open_legacy_filename_mapping(tmp_path):
    wrong = tmp_path / "bill.pdf"
    wrong.write_bytes(b"unrelated original")
    preserved = snapshot(
        [bill(**{"Source ID": "missing", "Attachment Name": "bill.pdf"})],
        {"saved_evidence_paths": {"bill.pdf": str(wrong)}},
    )
    entry = next(
        row
        for row in build_evidence_review(preserved, str(tmp_path)).entries
        if row.category == "Bills & charges"
    )
    assert entry.sources[0].path is None


def test_unique_legacy_filename_retains_its_recorded_pages(tmp_path):
    original = tmp_path / "bill.pdf"
    original.write_bytes(b"legacy original")
    preserved = snapshot(
        [bill(**{"Attachment Name": "bill.pdf", "Page From": 3, "Page To": 4})],
        {
            "source_manifest": {
                "sid": {
                    "Attachment Name": "bill.pdf",
                    "Saved Path": "bill.pdf",
                    "Content Hash": hashlib.sha256(original.read_bytes()).hexdigest(),
                }
            }
        },
    )
    entry = next(
        row
        for row in build_evidence_review(preserved, str(tmp_path)).entries
        if row.category == "Bills & charges"
    )
    assert entry.sources[0].path == str(original)
    assert entry.sources[0].pages == (3, 4)


@pytest.mark.parametrize("changed", [False, True])
def test_cli_report_cannot_overwrite_relocated_source_copy(tmp_path, capsys, changed):
    pack = tmp_path / "moved-pack"
    pack.mkdir()
    original = pack / "bill.pdf"
    original.write_bytes(b"relocated original")
    preserved = snapshot(
        [bill()],
        {
            "source_manifest": {
                "sid": {
                    "Saved Path": "bill.pdf",
                    "Content Hash": hashlib.sha256(original.read_bytes()).hexdigest(),
                }
            }
        },
    )
    source = pack / "result.json"
    save_result(preserved, source, result_id="source")
    if changed:
        original.write_bytes(b"changed original still must be preserved")
    original_bytes = original.read_bytes()
    with pytest.raises(SystemExit) as error:
        run_cli_pdf_report(["-i", str(source), "-o", str(original)])
    assert error.value.code == 1
    assert "overwrite" in capsys.readouterr().err
    assert original.read_bytes() == original_bytes
    assert not Path(str(original) + ".result.json").exists()


@pytest.mark.parametrize("payload", [None, [], "not a configuration"])
def test_cli_config_requires_object_with_clear_error(tmp_path, payload, capsys):
    config = tmp_path / "config.json"
    config.write_text(json.dumps(payload), encoding="utf-8")
    with pytest.raises(SystemExit) as error:
        run_cli_extract(
            ["--pdf-dir", str(tmp_path), "--config", str(config), "-o", str(tmp_path / "out.xlsx")]
        )
    assert error.value.code == 1
    assert "Configuration must be a JSON object" in capsys.readouterr().err
