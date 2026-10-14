"""Review regressions for source boundaries and the original evidence figures."""

from __future__ import annotations

from email.message import EmailMessage
from pathlib import Path

import pytest
from reportlab.pdfgen.canvas import Canvas

from edf_bill_fetcher.collectors.engine import EvidenceEngine
from edf_bill_fetcher.io.adapters.eml import parse_eml_message
from edf_bill_fetcher.io.adapters.html import parse_htm_account_history
from edf_bill_fetcher.io.adapters.pdf import slice_pdf_pages


def _engine() -> EvidenceEngine:
    return EvidenceEngine({"filter_below": False, "use_domain_filter": False}, lambda message: None)


def _pdf(path: Path, pages: list[str]) -> None:
    canvas = Canvas(str(path))
    for page in pages:
        y = 750
        for line in page.splitlines():
            canvas.drawString(40, y, line)
            y -= 20
        canvas.showPage()
    canvas.save()


@pytest.mark.parametrize(
    ("verb", "entry_type", "transaction"),
    [
        ("We charged your account", "New Bill", 25.0),
        ("You paid us", "Payment", -25.0),
        ("Reversed account charge", "Credit", -25.0),
    ],
)
def test_missing_htm_balance_preserves_each_stated_transaction(
    tmp_path: Path, verb: str, entry_type: str, transaction: float
) -> None:
    text = (
        f"01 Jan 2024 {verb} £25.00 Balance unavailable "
        f"02 Jan 2024 {verb} £50.00 Balance £100.00 in debit"
    )
    source = tmp_path / "history.htm"
    source.write_text(text, encoding="utf-8")
    engine = _engine()
    engine.process_htm_file(str(source))
    assert len(engine.records) == 2
    first, second = engine.records
    assert first["Entry Type"] == entry_type
    assert first["Transaction Amount (£)"] == transaction
    assert first["Balance (£)"] is None
    assert first["Amount (£)"] is None
    assert first["Source PDF Text"] == f"01 Jan 2024 {verb} £25.00 Balance unavailable"
    assert second["Date"] == "02/01/2024"
    assert second["Transaction Amount (£)"] == transaction * 2
    assert second["Balance (£)"] == 100.0


def test_htm_events_keep_source_order_and_consumption_dates_are_not_boundaries() -> None:
    text = (
        "03 Jan 2024 You paid us £25.00 Balance £100.00 in debit "
        "02 Jan 2024 Reversed account charge £50.00 Balance £125.00 in debit "
        "01 Jan 2024 We charged your account £100.00 "
        "For 250 kWh of electricity used between 01 Dec 2023 and 31 Dec 2023 "
        "Balance £175.00 in debit "
        "30 Nov 2023 Balance £75.00 in credit"
    )
    records = parse_htm_account_history(text)
    assert [row["Date"] for row in records] == [
        "03/01/2024",
        "02/01/2024",
        "01/01/2024",
        "30/11/2023",
    ]
    assert [row["Entry Type"] for row in records] == ["Payment", "Credit", "New Bill", "Credit"]
    assert records[2]["Period From"] == "01/12/2023"
    assert records[2]["Period To"] == "31/12/2023"
    assert records[2]["Units (kWh)"] == "250"
    assert records[3]["Amount (£)"] == -75.0
    assert records[3]["Period Charge (£)"] == "N/A"


@pytest.mark.parametrize(
    ("encoding", "declaration"),
    [
        ("utf-8", ""),
        ("cp1252", ""),
        ("cp1252", '<meta charset="windows-1252">'),
        ("utf-16", ""),
        ("utf-8", '<meta charset="utf-8">'),
    ],
)
def test_browser_export_encoding_preserves_pound_amounts(
    tmp_path: Path, encoding: str, declaration: str
) -> None:
    source = tmp_path / "history.htm"
    source.write_bytes(
        (
            f"<html><head>{declaration}</head><body>"
            "01 Jan 2024 You paid us £25.00 Balance £100.00 in debit</body></html>"
        ).encode(encoding)
    )
    engine = _engine()
    engine.process_htm_file(str(source))
    assert len(engine.records) == 1
    assert engine.records[0]["Transaction Amount (£)"] == -25.0
    assert engine.records[0]["Balance (£)"] == 100.0
    assert len(engine.source_manifest) == 1
    if declaration or encoding != "cp1252":
        assert engine.error_log == []
    else:
        assert "read as windows-1252" in engine.error_log[0]


def test_repeated_invoice_header_keeps_continuation_charges_and_original_pages(
    tmp_path: Path,
) -> None:
    pages = [
        "Invoice number: KI-12345\nDate issued: 15 January 2024\n"
        "Your charges: 01 Jan 2024 - 31 Jan 2024\nCurrent balance £200.00 debit\nPage 1 of 2",
        "Invoice number: KI-12345\nTotal charges for this period £100.00 debit\nPage 2 of 2",
    ]
    assert slice_pdf_pages(pages) == [pages]
    source = tmp_path / "bill.pdf"
    _pdf(source, pages)
    engine = _engine()
    engine.process_pdf_file(str(source), "Local PDF Folder", source.name, "01/01/2024")
    assert engine.error_log == []
    assert len(engine.records) == 1
    row = engine.records[0]
    assert row["Balance (£)"] == 200.0
    assert row["Transaction Amount (£)"] == 100.0
    assert (row["Page From"], row["Page To"]) == (1, 2)


def test_merged_credit_notes_without_page_numbers_extract_each_credit(tmp_path: Path) -> None:
    pages = [
        "Credit note number: KCR-12345\nDate issued: 15 January 2024\nTotal credits for this bill £25.00",
        "Credit note number: KCR-67890\nDate issued: 16 January 2024\nTotal credits for this bill £50.00",
    ]
    assert slice_pdf_pages(pages) == [[pages[0]], [pages[1]]]
    source = tmp_path / "credits.pdf"
    _pdf(source, pages)
    engine = _engine()
    engine.process_pdf_file(str(source), "Local PDF Folder", source.name, "01/01/2024")
    assert engine.error_log == []
    assert [row["Invoice #"] for row in engine.records] == ["KCR-12345", "KCR-67890"]
    assert [row["Transaction Amount (£)"] for row in engine.records] == [-25.0, -50.0]
    assert [(row["Page From"], row["Page To"]) for row in engine.records] == [(1, 1), (2, 2)]


def test_repeated_credit_number_is_continuation_but_page_one_starts_an_occurrence() -> None:
    pages = [
        "Credit note number: KCR-1 Page 1 of 2",
        "Credit note number: KCR-1 Page 2 of 2",
        "Credit note number: KCR-1 Page 1 of 2",
        "Credit note number: KCR-1 Page 2 of 2",
    ]
    assert slice_pdf_pages(pages) == [pages[:2], pages[2:]]


def test_attached_email_html_cannot_override_current_message_evidence(tmp_path: Path) -> None:
    current = EmailMessage()
    current["From"] = "billing@edfenergy.com"
    current["Subject"] = "Your EDF Bill"
    current.set_content("Your new account balance £200.00")
    old = EmailMessage()
    old["Subject"] = "Old EDF Bill"
    old.set_content("<p>Your new account balance £1000.00</p>", subtype="html")
    current.add_attachment(old, filename="old.eml")
    source = tmp_path / "bill.eml"
    source.write_bytes(current.as_bytes())
    parsed = parse_eml_message(source)
    assert parsed["body_html"] == ""
    assert "£200.00" in parsed["body_text"]
    engine = _engine()
    engine.process_eml_folder(str(tmp_path))
    assert len(engine.records) == 1
    assert engine.records[0]["Balance (£)"] == 200.0
    assert len(engine.source_manifest) == 1


def test_unpadded_slash_dates_cannot_split_a_pdf() -> None:
    from edf_bill_fetcher.io.adapters.pdf import PAGE1_BOUNDARY_RE

    pages = [
        "Invoice number: KI-12345 Page 1 of 3",
        "Charges for 1/4/2024 and reading on 1/5/2024 Page 2 of 3",
        "Total charges for this period £100.00 Page 3 of 3",
    ]
    assert not PAGE1_BOUNDARY_RE.search("Charges for 1/4/2024")
    assert slice_pdf_pages(pages) == [pages]


def test_pdf_cancel_stops_at_page_boundary(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    from unittest.mock import MagicMock

    engine = _engine()
    source = tmp_path / "bill.pdf"
    source.write_bytes(b"%PDF synthetic test input")
    first, second = MagicMock(), MagicMock()

    def first_page() -> str:
        engine.cancel_event.set()
        return "Your new account balance £200.00"

    first.extract_text.side_effect = first_page
    second.extract_text.return_value = "Your new account balance £1000.00"
    document = MagicMock()
    document.pages = [first, second]
    document.__enter__.return_value = document
    monkeypatch.setattr(
        "edf_bill_fetcher.collectors.engine.pdfplumber.open", lambda stream: document
    )
    engine.process_pdf_file(str(source), "Local PDF Folder", source.name, "01/01/2024")
    first.extract_text.assert_called_once()
    second.extract_text.assert_not_called()
    assert engine.records == []
    assert engine.error_log == []


def test_first_invoice_number_on_continuation_stays_with_cover_page() -> None:
    pages = [
        "Your bill Page 1 of 3 Current balance £200.00",
        "Invoice number: KI-12345 Page 2 of 3 Total charges for this period £100.00",
        "Invoice number: KI-12345 Page 3 of 3 Your charges in detail",
        "Your next bill Page 1 of 2 Current balance £300.00",
        "Invoice number: KI-67890 Page 2 of 2 Total charges for this period £100.00",
    ]
    assert slice_pdf_pages(pages) == [pages[:3], pages[3:]]
