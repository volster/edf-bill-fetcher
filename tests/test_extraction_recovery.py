"""Regression cases from nonempty revision results and damaged PST messages."""

from datetime import date, datetime, timezone
from io import BytesIO
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest
from reportlab.pdfgen.canvas import Canvas

from edf_bill_fetcher.application.run_service import run
from edf_bill_fetcher.collectors.engine import EvidenceEngine
from edf_bill_fetcher.domain.json_values import freeze_json, object_value
from edf_bill_fetcher.domain.requests import RunRequest
from edf_bill_fetcher.io.persistence.result_codec import (
    decode_result,
    encode_result,
    migrate_legacy,
)
from edf_bill_fetcher.io.persistence.result_store import load_result
from edf_bill_fetcher.processors.sap_parsers import _parse_sap_csv


def _request():
    return RunRequest(
        "extraction-recovery",
        datetime(2024, 4, 1, tzinfo=timezone.utc),
        date(2024, 4, 1),
        (),
        freeze_json({"as_of": "2024-04-01", "analysis_min": 0}),
    )


def _revision_records(expanded_period):
    return [
        {
            "Account": "A-123",
            "Entry Type": "New Bill",
            "Invoice #": invoice,
            "Date": when,
            "Period From": start,
            "Period To": "31/12/2020",
            "Period Charge (£)": 100,
            "Amount (£)": balance,
        }
        for invoice, when, balance, start in (
            ("OLD", "01/01/2023", 100, "01/01/2020"),
            ("NEW", "01/02/2023", 120, "01/12/2019" if expanded_period else "01/01/2020"),
        )
    ]


@pytest.mark.parametrize("expanded_period", [False, True])
def test_nonempty_superseded_result_survives_save_reload_and_all_report_formats(
    tmp_path, expanded_period
):
    outcome = run(
        _request(),
        {fmt: str(tmp_path / ("report." + fmt)) for fmt in ("xlsx", "pdf", "docx", "html")},
        result_path=str(tmp_path / "result.json"),
        manifest_path=str(tmp_path / "run.json"),
        collect=lambda: (_revision_records(expanded_period), {}),
    )
    assert outcome.status == "complete", outcome.error
    restored = load_result(tmp_path / "result.json")
    assert restored == outcome.snapshot
    assert object_value(restored.models)["sap"]["payload"]["domination"] == {"OLD": ["NEW", False]}
    assert set(object_value(outcome.manifest)["artifacts"]) == {"xlsx", "pdf", "docx", "html"}


@pytest.mark.parametrize("partial_overlap", [False, True])
def test_stored_boolean_overlap_flags_roundtrip_without_string_coercion(partial_overlap):
    accepted = migrate_legacy([_revision_records(False)[0]], _request())
    payload = encode_result(accepted, "valid")
    payload["analysis"]["models"]["sap"]["payload"]["domination"] = {
        "OLD": ["NEW", partial_overlap]
    }
    restored = decode_result(payload)
    assert encode_result(restored, "reloaded")["analysis"]["models"]["sap"]["payload"][
        "domination"
    ] == {"OLD": ["NEW", partial_overlap]}


@pytest.mark.parametrize(
    "mapping",
    [
        ["NEW", "False"],
        ["NEW", 0],
        ["NEW", None],
        ["", False],
        [7, False],
        ["NEW"],
        ["NEW", False, True],
    ],
)
def test_malformed_superseded_entries_still_fail_validation(mapping):
    accepted = migrate_legacy([_revision_records(False)[0]], _request())
    payload = encode_result(accepted, "valid")
    payload["analysis"]["models"]["sap"]["payload"]["domination"] = {"OLD": mapping}
    with pytest.raises(ValueError, match="superseded invoice mapping"):
        decode_result(payload)


def _message(amount=200):
    return MagicMock(
        get_subject=MagicMock(return_value="Synthetic EDF bill"),
        get_delivery_time=MagicMock(return_value=datetime(2024, 1, 1)),
        get_transport_headers=MagicMock(return_value="From: billing@edfenergy.com"),
        get_html_body=MagicMock(return_value=None),
        get_plain_text_body=MagicMock(
            return_value=f"Your new account balance £{amount}.00".encode()
        ),
        get_rtf_body=MagicMock(return_value=None),
        get_number_of_attachments=MagicMock(return_value=0),
    )


def _folder(messages, name="Inbox", children=()):
    return SimpleNamespace(
        get_name=lambda: name,
        get_number_of_sub_messages=lambda: len(messages),
        get_sub_message=lambda index: messages[index],
        get_number_of_sub_folders=lambda: len(children),
        get_sub_folder=lambda index: children[index],
    )


@pytest.fixture
def engine(monkeypatch):
    monkeypatch.setattr("edf_bill_fetcher.collectors.engine.HAS_PYPFF", True)
    value = EvidenceEngine({"filter_below": False, "use_domain_filter": False}, lambda text: None)
    yield value
    value.cleanup()


def _pdf():
    stream = BytesIO()
    canvas = Canvas(stream)
    canvas.drawString(40, 750, "Invoice number: KI-12345")
    canvas.drawString(40, 730, "Date issued: 15 January 2024")
    canvas.drawString(40, 710, "Total charges for this period £125.00 debit")
    canvas.drawString(40, 690, "Current balance £250.00 debit")
    canvas.save()
    return stream.getvalue()


def test_pst_attachment_table_failure_retains_body_and_processes_later_folders(engine):
    damaged = _message()
    native_error = "pypff_message_get_number_of_attachments: invalid local descriptors node: 1649"
    damaged.get_number_of_attachments.side_effect = OSError(native_error)
    engine.crawl_pst(_folder([damaged], children=(_folder([_message(300)], "Archive"),)))
    assert [row["Balance (£)"] for row in engine.records] == [200, 300]
    assert engine.email_count == 2
    assert len(engine.error_log) == 1
    diagnostic = engine.error_log[0]
    assert "PST attachments" in diagnostic
    assert "Inbox" in diagnostic
    assert "message index 0" in diagnostic
    assert "Synthetic EDF bill" in diagnostic
    assert "attachments may be missing" in diagnostic
    assert native_error in diagnostic


def test_pst_bad_html_falls_back_to_plain_and_still_extracts_pdf(engine):
    message = _message()
    message.get_html_body.side_effect = OSError("damaged HTML property")
    data = _pdf()
    message.get_number_of_attachments.return_value = 1
    message.get_attachment.return_value = MagicMock(
        get_size=lambda: len(data), read_buffer=lambda size: data
    )
    engine.crawl_pst(_folder([message]))
    assert {row["Balance (£)"] for row in engine.records} == {200, 250}
    assert engine.pdf_count == 1
    assert any("damaged HTML property" in entry for entry in engine.error_log)


def test_pst_html_does_not_read_unneeded_damaged_plain_body(engine):
    message = _message()
    message.get_html_body.return_value = "<p>Your new account balance £200.00</p>"
    message.get_plain_text_body.side_effect = OSError("unneeded damaged plain property")
    engine.crawl_pst(_folder([message]))
    assert engine.records[0]["Balance (£)"] == 200
    message.get_plain_text_body.assert_not_called()
    assert engine.error_log == []


def test_pst_unreadable_bodies_do_not_prevent_attachment_extraction(engine):
    message = _message()
    message.get_html_body.side_effect = OSError("damaged HTML")
    message.get_plain_text_body.side_effect = OSError("damaged plain text")
    data = _pdf()
    message.get_number_of_attachments.return_value = 1
    message.get_attachment.return_value = MagicMock(
        get_size=lambda: len(data), read_buffer=lambda size: data
    )
    engine.crawl_pst(_folder([message]))
    assert len(engine.records) == 1
    assert engine.records[0]["Invoice #"] == "KI-12345"
    assert any("No readable body" in entry for entry in engine.error_log)


@pytest.mark.parametrize(
    "name", ["image001.png", "ATT00001.txt", "signature.jpg", "unknown-attachment"]
)
def test_non_pdf_attachments_are_counted_as_skipped_without_parse_errors(engine, name):
    engine._process_attachment(
        b"ordinary non-PDF content",
        name,
        "PST PDF Attachment",
        "01/01/2024",
        "",
        "Synthetic EDF bill",
    )
    assert engine.attachment_count == 1
    assert engine.skipped_attachment_count == 1
    assert engine.error_log == []
    assert engine.records == []


@pytest.mark.parametrize("data", [b"damaged PDF", b""])
def test_attachment_named_pdf_with_invalid_bytes_remains_a_review_error(engine, data):
    engine._process_attachment(
        data, "invoice.PDF", "PST PDF Attachment", "01/01/2024", "", "Synthetic EDF bill"
    )
    assert engine.skipped_attachment_count == 1
    assert len(engine.error_log) == 1
    assert "PDF" in engine.error_log[0]
    assert "invoice.PDF" in engine.error_log[0]


def test_sap_wide_row_warning_identifies_two_affected_rows_without_aborting_other_rows():
    csv = "Reference,Amount\n" + "\n".join(
        f"TX-{index},100.00" + (",unexpected" if index in {3, 905} else "")
        for index in range(1, 906)
    )
    reports: list[str] = []
    rows = _parse_sap_csv(csv, reports.append)
    assert len(rows) == 905
    assert rows[2] == {"Reference": "TX-3", "Amount": "100.00"}
    assert rows[-1] == {"Reference": "TX-905", "Amount": "100.00"}
    assert len(reports) == 1
    assert "0 short row(s), 2 long row(s) out of 905 parsed" in reports[0]
    assert "long: 3, 905" in reports[0]
    assert "extra fields omitted" in reports[0]


def test_pst_ignored_attachments_do_not_force_an_otherwise_clean_run_into_review(engine, tmp_path):
    def collect():
        engine._process_attachment(
            b"logo", "image001.png", "PST PDF Attachment", "01/01/2024", "", "Synthetic EDF bill"
        )
        engine.crawl_pst(_folder([_message()]))
        return engine.records, engine.__getstate__()

    outcome = run(
        _request(),
        {"html": str(tmp_path / "report.html")},
        result_path=str(tmp_path / "result.json"),
        manifest_path=str(tmp_path / "run.json"),
        collect=collect,
    )
    assert outcome.status == "complete", outcome.error
    assert load_result(tmp_path / "result.json") == outcome.snapshot
    assert engine.skipped_attachment_count == 1


def test_pst_unreadable_attachment_table_marks_partial_pack_for_review(engine, tmp_path):
    damaged = _message()
    damaged.get_number_of_attachments.side_effect = OSError("invalid local descriptors node")

    def collect():
        engine.crawl_pst(_folder([damaged]))
        return engine.records, engine.__getstate__()

    outcome = run(
        _request(),
        {"html": str(tmp_path / "report.html")},
        result_path=str(tmp_path / "result.json"),
        manifest_path=str(tmp_path / "run.json"),
        collect=collect,
    )
    assert outcome.status == "review_required", outcome.error
    assert load_result(tmp_path / "result.json") == outcome.snapshot
    assert "html" in object_value(outcome.manifest)["artifacts"]
    assert engine.records[0]["Balance (£)"] == 200
