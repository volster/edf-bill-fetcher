"""Mixed evidence must stay source-backed, bounded and explicit about unprocessed files."""

import hashlib
import io
import json
from datetime import date, datetime, timezone
from email.message import EmailMessage
from pathlib import Path
from zipfile import ZipFile, ZipInfo

import pytest
from reportlab.pdfgen.canvas import Canvas

from edf_bill_fetcher.application.run_service import bundle_engine_sources, collect_engine, run
from edf_bill_fetcher.collectors.engine import EvidenceEngine
from edf_bill_fetcher.domain.json_values import freeze_json, object_value
from edf_bill_fetcher.domain.requests import RunRequest, SourceSelection
from edf_bill_fetcher.io.adapters.intake import IntakeLimits, detect, discover
from edf_bill_fetcher.models.config import ConfigDict


def bill(invoice, charge):
    output = io.BytesIO()
    canvas = Canvas(output)
    for index, text in enumerate(
        [
            f"Invoice number: {invoice}",
            "Account number: A-12345678",
            "Date issued: 15 February 2024",
            "Your charges: 01 Jan 2024 - 31 Jan 2024",
            f"Total charges for this period £{charge}.00 debit",
            f"Current balance £{charge}.00 debit",
        ]
    ):
        canvas.drawString(40, 780 - index * 30, text)
    canvas.save()
    return output.getvalue()


@pytest.mark.parametrize(
    ("name", "content", "kind"),
    [
        ("renamed.dat", b"%PDF-1.7\n", "pdf"),
        ("mailbox.ost", b"!BDN" + b"\0" * 10, "pst"),
        ("mail.MSG", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1", "msg"),
        ("unknown.ole", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1", "unsupported"),
        ("folder.zip", b"PK\x03\x04", "zip"),
        ("review.xlsx", b"PK\x03\x04", "workbook"),
        ("history.dat", b"<!DOCTYPE html><html>", "htm"),
        ("mail.dat", b"From: billing@edfenergy.com\nSubject: EDF bill\n\n", "eml"),
        ("result.json", b'{"schema_version":3}', "result"),
        ("image.png", b"\x89PNG\r\n", "unsupported"),
    ],
)
def test_detection_uses_content_without_turning_unknown_documents_into_bills(name, content, kind):
    assert detect(name, content) == kind


def mixed_folder(root):
    root.mkdir()
    (root / "first.pdf").write_bytes(bill("KI-10001", 100))
    message = EmailMessage()
    message["From"] = "billing@edfenergy.com"
    message["Subject"] = "EDF bill"
    message.set_content("Total charges for this period £30.00 debit")
    (root / "email.eml").write_bytes(message.as_bytes())
    (root / "photo.png").write_bytes(b"not parsed as billing text")
    archive = root / "more.zip"
    with ZipFile(archive, "w") as zipped:
        zipped.writestr("2024/second.pdf", bill("KI-10002", 50))
    return root


def test_mixed_folder_and_zip_share_normal_financial_pipeline_and_preserve_members(tmp_path):
    sources = mixed_folder(tmp_path / "inputs")
    destination = tmp_path / "pack"
    destination.mkdir()
    config: ConfigDict = {
        "as_of": "2025-03-01",
        "analysis_min": 0,
        "filter_below": False,
        "use_dedup": False,
        "use_domain_filter": False,
    }
    request = RunRequest(
        "mixed",
        datetime(2025, 3, 1, tzinfo=timezone.utc),
        date(2025, 3, 1),
        (SourceSelection("auto", str(sources)),),
        freeze_json(config),
    )
    engine = EvidenceEngine(config, lambda _: None)
    outcome = run(
        request,
        {},
        result_path=str(destination / "result.json"),
        manifest_path=str(destination / "run.json"),
        collect=lambda: collect_engine(engine, request),
        copy_sources=lambda prepared: bundle_engine_sources(engine, prepared, str(destination)),
        cleanup=engine.cleanup,
    )
    assert outcome.status == "review_required", outcome.error
    assert outcome.snapshot is not None
    assert outcome.snapshot.summary.charges == 180
    state = object_value(outcome.snapshot.prepared.extraction.legacy_state)
    assert any(
        item["status"] == "unsupported" and item["name"] == "photo.png"
        for item in state["intake_inventory"]
    )
    member = next(
        value
        for value in state["source_manifest"].values()
        if value.get("Member") == "2024/second.pdf"
    )
    assert (
        member["Container SHA256"]
        == hashlib.sha256((sources / "more.zip").read_bytes()).hexdigest()
    )
    with ZipFile(sources / "more.zip") as archive:
        assert (destination / member["Saved Path"]).read_bytes() == archive.read("2024/second.pdf")
    assert not Path(member["Original Path"]).exists()
    assert engine._attachment_dir is None


@pytest.mark.parametrize(
    "name", ["../escape.pdf", "/absolute.pdf", "C:/secret.pdf", "folder\\secret.pdf"]
)
def test_archive_traversal_is_rejected_before_any_member_is_written(tmp_path, name):
    path = tmp_path / "bad.zip"
    with ZipFile(path, "w") as archive:
        archive.writestr("good.pdf", bill("KI-10003", 20))
        info = ZipInfo(name)
        # Preserve wire spelling even where ZipInfo normalises local separators.
        info.filename = name
        archive.writestr(info, b"outside")
    rows = list(discover(path, tmp_path / "expanded"))
    assert rows[0]["status"] == "failed"
    assert not (tmp_path / "expanded").exists()


@pytest.mark.parametrize(
    "case", ["case_collision", "symlink", "expanded_bytes", "file_count", "depth"]
)
def test_archive_limits_have_explicit_failed_inventory_entries(tmp_path, case):
    source = tmp_path / "input.zip"
    with ZipFile(source, "w") as archive:
        archive.writestr("bill.pdf", b"%PDF-1.7")
        if case == "case_collision":
            archive.writestr("BILL.pdf", b"%PDF-1.7")
        if case == "symlink":
            info = ZipInfo("link.pdf")
            info.external_attr = 0o120777 << 16
            archive.writestr(info, b"/private")
    limits = IntakeLimits(
        files=1 if case == "file_count" else 100,
        expanded_bytes=1 if case == "expanded_bytes" else 1000,
        depth=0 if case == "depth" else 3,
    )
    rows = list(discover(source, tmp_path / "expanded", limits=limits))
    assert rows[0]["status"] == "failed"
    assert rows[0]["detail"]


def test_saved_pack_requires_import_choice_and_cancel_does_not_expand(tmp_path):
    source = tmp_path / "saved.zip"
    with ZipFile(source, "w") as archive:
        archive.writestr("old/result.json", json.dumps({"schema_version": 3}))
        archive.writestr("old/report.pdf", b"%PDF-1.7")
    rows = list(discover(source, tmp_path / "expanded"))
    assert len(rows) == 1
    assert rows[0]["kind"] == "bundle"
    assert rows[0]["status"] == "needs_choice"
    assert not (tmp_path / "expanded").exists()
    assert list(discover(source, tmp_path / "cancelled", cancelled=lambda: True)) == []


@pytest.mark.parametrize(
    "names",
    [
        ["folder", "folder/bill.pdf"],
        ["folder/bill.pdf", "folder"],
        ["folder/", "FOLDER"],
        ["../unsafe/", "bill.pdf"],
    ],
)
def test_archive_file_directory_conflicts_are_rejected_before_expansion(tmp_path, names):
    source = tmp_path / "bad.zip"
    with ZipFile(source, "w") as archive:
        for name in names:
            archive.writestr(name, b"%PDF-1.7")
    rows = list(discover(source, tmp_path / "expanded"))
    assert rows[0]["status"] == "failed"
    assert not (tmp_path / "expanded").exists()


def test_inspected_message_cannot_bypass_the_source_read_policy(tmp_path):
    from edf_bill_fetcher.application.source_intake import collect_mixed

    outside = tmp_path / "private.eml"
    outside.write_bytes(b"From: billing@edfenergy.com\nSubject: EDF\n\nSecret")
    engine = EvidenceEngine({}, lambda _: None, allowed_source=lambda path: False)
    try:
        collect_mixed(
            engine,
            str(outside),
            plan=[
                {
                    "path": str(outside),
                    "name": "private.eml",
                    "kind": "eml",
                    "status": "ready",
                    "detail": "",
                    "container": None,
                }
            ],
        )
        assert not engine.records
        assert any("outside the configured" in row for row in engine.error_log)
        assert engine._attachment_dir is not None
        assert not list(Path(engine._attachment_dir.name).rglob("*.eml"))
    finally:
        engine.cleanup()


def test_intake_reports_missing_folders_and_renamed_pdf_previews(tmp_path):
    from edf_bill_fetcher.io.adapters.intake import is_pdf_candidate

    rows = list(discover(tmp_path / "missing", tmp_path / "staging"))
    assert len(rows) == 1
    assert rows[0]["status"] == "failed"
    renamed = tmp_path / "renamed.dat"
    renamed.write_bytes(bill("KI-10003", 100))
    assert is_pdf_candidate(str(renamed))
    assert not is_pdf_candidate(str(tmp_path / "missing.dat"))
    renamed.write_bytes(b"unsupported content")
    assert not is_pdf_candidate(str(renamed))


def test_archive_expansion_handles_explicit_folders_and_cancellation(tmp_path):
    from threading import Event

    path = tmp_path / "input.zip"
    with ZipFile(path, "w") as archive:
        archive.writestr("folder/", b"")
        archive.writestr("folder/bill.pdf", bill("KI-10003", 100))
    rows = list(discover(path, tmp_path / "expanded"))
    assert [row["status"] for row in rows] == ["container", "ready"]
    cancel = Event()
    pending = discover(path, tmp_path / "cancelled", cancelled=cancel.is_set)
    assert next(pending)["status"] == "container"
    cancel.set()
    assert list(pending) == []
    assert not list((tmp_path / "cancelled").rglob("*.pdf"))


def test_archive_expansion_reports_insufficient_storage_without_publishing_members(
    tmp_path, monkeypatch
):
    from types import SimpleNamespace

    import edf_bill_fetcher.io.adapters.intake as intake_module

    path = tmp_path / "input.zip"
    with ZipFile(path, "w") as archive:
        archive.writestr("bill.pdf", bill("KI-10003", 100))
    monkeypatch.setattr(intake_module.shutil, "disk_usage", lambda path: SimpleNamespace(free=0))
    rows = list(discover(path, tmp_path / "expanded"))
    assert rows[0]["status"] == "failed"
    assert "Not enough free space" in rows[0]["detail"]
    assert not list((tmp_path / "expanded").rglob("*.pdf"))


def test_raw_zip_name_is_validated_before_windows_normalisation():
    from edf_bill_fetcher.io.adapters.intake import validate_zip_members

    info = ZipInfo("folder/secret.pdf")
    info.orig_filename = "folder\\secret.pdf"
    with pytest.raises(ValueError, match="Unsafe ZIP"):
        validate_zip_members([info])
    info.orig_filename = "secret\x00.pdf"
    with pytest.raises(ValueError, match="Unsafe ZIP"):
        validate_zip_members([info])
