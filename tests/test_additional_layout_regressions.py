"""Synthetic checks for additional source layouts, including printed portal history."""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from types import SimpleNamespace

import pandas as pd
import pytest
from reportlab.pdfgen.canvas import Canvas

from edf_bill_fetcher.collectors.engine import EvidenceEngine
from edf_bill_fetcher.io.adapters.html import parse_pdf_account_history
from edf_bill_fetcher.io.adapters.ocr import OcrPage, PdfOcrReader, requires_ocr
from edf_bill_fetcher.io.adapters.pdf import join_sap_pdf_pages
from edf_bill_fetcher.processors.bill_facts import bill_facts, payment_history
from edf_bill_fetcher.processors.detection import detect_back_billing
from edf_bill_fetcher.processors.patterns import extract_sub_periods
from edf_bill_fetcher.processors.preparation import prepare_evidence
from edf_bill_fetcher.processors.sap_parsers import _parse_sap_csv

PORTAL = (
    "Payments and Invoices | EDF https://edfenergy.com/myaccount/accounts/A-12345678/property/123"
)
CHARGE = "30 Jun 2024\nWe charged your account £200.00\nFor 500 kWh of electricity used between 1 Jun 2024 and 30 Jun 2024\nBalance £1200.00 in debit"
PAYMENT = "2 Jul 2024\nYou paid us £100.00\nBank Transfer\nBalance £1100.00 in debit"


def _engine() -> EvidenceEngine:
    return EvidenceEngine({"filter_below": False}, lambda _: None)


def _pdf(path: Path, pages: list[str]) -> None:
    canvas = Canvas(str(path))
    for text in pages:
        for n, line in enumerate(text.splitlines()):
            canvas.drawString(20, 800 - n * 20, line)
        canvas.showPage()
    canvas.save()


def test_multi_section_invoice_sums_consumption_and_uses_latest_reading():
    text = """Electricity charges for meter D99T12345
1 May 2025 1000.0000 Estimated reading
13 May 2025 1653.0000 Estimated reading
653.0000 kWh @
Electricity used 40.000p/kWh £261.20
Government Climate 653.0 kWh @ 0.000 £/kWh £0.00
Total electricity charges for this period £300.00
Electricity charges for meter D99T12345
1 Apr 2025 200.0000 Estimated reading
1 May 2025 1000.0000 Estimated reading
Electricity used 800.0000 kWh @ 40.000p/kWh £320.00
Total electricity charges for this period £400.00
To avoid estimates, please provide an actual reading.
"""
    facts = bill_facts(text, [])
    assert facts["Units (kWh)"] == 1453
    assert facts["Gross Electricity Charges (£)"] == 700
    assert facts["Register Reading (kWh)"] == 1653
    assert facts["Reading Date"] == "13/05/2025"
    assert facts["Reading"] == facts["Closing Reading Type"] == "Estimated"


@pytest.mark.parametrize(
    "extra",
    [
        "Electricity charges for meter D88T98765\n1 Aug 2024 500 Customer reading",
        "1 Aug 2024 600 Actual reading",
    ],
)
def test_ambiguous_register_readings_are_not_collapsed(extra):
    facts = bill_facts(
        "Electricity charges for meter D99T12345\n1 Aug 2024 500 Customer reading\n" + extra, []
    )
    assert "Register Reading (kWh)" not in facts


def test_invalid_reading_dates_do_not_discard_valid_bill_facts():
    facts = bill_facts(
        "31 February 2024 200 Estimated reading\nElectricity used 50 kWh @ 30p/kWh £15.00", []
    )
    assert facts["Units (kWh)"] == 50
    assert "Reading Date" not in facts


def test_data_collector_reading_is_actual():
    facts = bill_facts("7 July 2025 1000.0000 Data collector reading", [])
    assert facts["Closing Reading Type"] == "Actual"


def test_wrapped_old_table_does_not_borrow_interest_periods():
    text = """Late Payment Interest for 14 Jun 20 - 28 Jun 20 @ 8.75% on £135.82 £0.49
Meter: D99T12345 Tariff: Example Product SME
13 Jun 20 - 1000 1500 500 kWh 25.00p £125.00
10 Jul 20 OUR READ OUR READ
"""
    periods = extract_sub_periods(text)
    assert len(periods) == 1
    assert periods[0]["period_from"] == "13/06/2020"
    assert periods[0]["period_to"] == "10/07/2020"
    facts = bill_facts(text, periods)
    assert facts["Register Reading (kWh)"] == 1500
    assert facts["Reading Date"] == "10/07/2020"
    assert facts["Tariff"] == "Example"


def test_old_table_year_wrap_and_missing_dash_keep_both_periods():
    text = """09 Dec 31 Dec 23 1000 ESTIMATED 1500 ESTIMATED 500 kWh 30.00p £150.00
23 -
01 Jan 24 10 Jan 24 1500 ESTIMATED 1800 ESTIMATED 300 kWh 40.00p £120.00
-
"""
    periods = extract_sub_periods(text)
    assert [(r["period_from"], r["period_to"]) for r in periods] == [
        ("09/12/2023", "31/12/2023"),
        ("01/01/2024", "10/01/2024"),
    ]
    assert bill_facts(text, periods)["Register Reading (kWh)"] == 1800


def test_parallel_usage_header_does_not_hide_payment_history():
    rows = payment_history(
        "Your payment history About your usage Page 3 of 3\n28 June 2020 £100.00\nFrequently asked questions\n1 July 2020 £200.00"
    )
    assert [(r["Date"], r["Transaction Amount (£)"]) for r in rows] == [("28/06/2020", -100)]


def test_renewal_notice_preserves_stated_balance_without_annual_quote():
    text = """18 January 2023
Important news about your energy contract
Your outstanding balance of £500.00.
Account number 601234567890
Your estimated annual bill £30000.00
Estimated Annual Consumption 100000 kWh
Please see your Terms and Conditions.
"""
    engine = _engine()
    engine.process_text(text, "Local PDF Folder", "renewal.pdf", None)
    row = engine.records[0]
    assert row["Entry Type"] == "Ongoing Balance"
    assert row["Document Role"] == "Contract notice"
    assert row["Date"] == "18/01/2023"
    assert row["Balance (£)"] == 500
    assert row["Transaction Amount (£)"] is None
    assert row["Units (kWh)"] == "N/A"


def test_renewal_notice_without_a_stated_balance_does_not_use_annual_quote():
    engine = _engine()
    engine.process_text(
        "Important news about your energy contract\nYour estimated annual bill £30000.00",
        "PDF",
        "renewal.pdf",
        None,
    )
    assert engine.records[0]["Entry Type"] == "Supporting Document"
    assert engine.records[0]["Balance (£)"] is None
    assert engine.records[0]["Transaction Amount (£)"] is None


@pytest.mark.parametrize(
    "title", ["THE PRINCIPAL TERMS", "Contract for the supply of energy", "Terms and Conditions"]
)
def test_contract_documents_do_not_become_charges(title):
    engine = _engine()
    engine.process_text(
        title + "\nAdministration fee £250.00; annual 100000 kWh", "PDF", "terms.pdf", "17/10/2026"
    )
    assert engine.records[0]["Entry Type"] == "Supporting Document"
    assert engine.records[0]["Transaction Amount (£)"] is None
    assert engine.records[0]["Date"] == "N/A"


def test_history_fee_is_not_a_consumption_bill():
    rows = parse_pdf_account_history(
        "25 Apr 2025 We charged your account £30.00 Late Payment Charge (Small Business) Balance £1000.00 in debit"
    )
    engine = _engine()
    engine._add_record(rows[0])
    fee = engine.records[0]
    assert fee["Entry Type"] == "Late Payment"
    assert fee["Date"] == "25/04/2025"
    assert fee["Transaction Amount (£)"] == 30
    assert fee["Period From"] == "N/A"


def test_ocr_detects_portal_vector_headings_even_with_native_balances():
    page = SimpleNamespace(images=[], width=650, height=900)
    assert requires_ocr(
        PORTAL
        + "\nFor 500 kWh of used between 1 Jun 2024 and 30 Jun 2024\nBalance £1200.00 in debit",
        page,
    )
    assert not requires_ocr(PORTAL + "\n" + CHARGE, page)
    assert not requires_ocr("Invoice number: KI-12345\nCurrent balance £1200.00 debit", page)


def test_browser_header_between_date_and_card_keeps_charge():
    rows = parse_pdf_account_history("30 Jun 2024\n" + PORTAL + "\n" + CHARGE.split("\n", 1)[1])
    assert len(rows) == 1
    assert rows[0]["Posting Date"] == "30/06/2024"
    assert rows[0]["Date"] == "N/A"
    assert rows[0]["First Billed Date"] == "N/A"
    assert rows[0]["Entry Type"] == "Charge"
    assert detect_back_billing(pd.DataFrame(rows)).empty


@pytest.mark.parametrize("ocr_error", [False, True])
def test_portal_pdf_preserves_native_identity_and_page_provenance(tmp_path, monkeypatch, ocr_error):
    path = tmp_path / "portal.pdf"
    _pdf(path, [PORTAL + "\nBalance £1200.00 in debit", PORTAL + "\n" + PAYMENT])
    original = path.read_bytes()
    calls = []

    def recognize(self, raw, index, cancelled):
        calls.append(index)
        if ocr_error:
            raise RuntimeError("OCR unavailable")
        return OcrPage(PORTAL.replace("12345678", "1234 5678") + "\n" + CHARGE, 88)

    monkeypatch.setattr(PdfOcrReader, "read_page", recognize)
    engine = _engine()
    engine.process_pdf_file(str(path), "Local PDF Folder", path.name, None)
    assert calls == [0]
    assert path.read_bytes() == original
    assert all(row["Account"] == "A-12345678" for row in engine.records)
    payment = next(r for r in engine.records if r["Entry Type"] == "Payment")
    assert (payment["Page From"], payment["Page To"]) == (2, 2)
    assert payment["Transaction Amount (£)"] == -100
    assert payment["OCR Review Required"] is False
    assert payment["OCR Minimum Confidence (%)"] is None
    if not ocr_error:
        charge = next(r for r in engine.records if r["Entry Type"] == "Charge")
        assert charge["Transaction Amount (£)"] == 200
        assert charge["Date"] == "N/A"
        assert charge["OCR Pages"] == [1]
        assert charge["OCR Minimum Confidence (%)"] == 88
    else:
        assert len(engine.records) == 1
        assert any("OCR unavailable" in error for error in engine.error_log)


def test_unreadable_portal_cannot_fall_back_to_a_fake_invoice(tmp_path, monkeypatch):
    path = tmp_path / "portal.pdf"
    _pdf(path, [PORTAL + "\nBalance £1200.00 in debit\nWebsite terms and conditions"])
    monkeypatch.setattr(PdfOcrReader, "read_page", lambda *args: OcrPage("", None))
    engine = _engine()
    engine.process_pdf_file(str(path), "Local PDF Folder", path.name, None)
    assert len(engine.records) == 1
    assert engine.records[0]["Document Role"] == "Account history: incomplete text"
    assert engine.records[0]["Transaction Amount (£)"] is None


def _histories(tmp_path):
    source = tmp_path / "history.htm"
    source.write_text(
        '<a href="https://edfenergy.com/myaccount/accounts/A-12345678">Account</a><p>'
        + CHARGE
        + "</p><p>"
        + PAYMENT
        + "</p>",
        encoding="utf-8",
    )
    engine = _engine()
    engine.process_htm_file(str(source))
    pdf_rows = parse_pdf_account_history(CHARGE + "\n" + PAYMENT)
    for i, row in enumerate(pdf_rows):
        row.update(
            {
                "Source ID": "printed-history",
                "Account": "A-12345678",
                "Fuel": "electricity",
                "Source Row": i + 1,
            }
        )
    return engine.records, pdf_rows


def test_matching_html_pdf_history_deduplicates_with_both_sources(tmp_path):
    html, pdf = _histories(tmp_path)
    assert html[0]["Account"] == "A-12345678"
    assert html[0]["Posting Date"] == "30/06/2024"
    assert html[0]["First Billed Date"] == "N/A"
    result = prepare_evidence(html + pdf)
    assert len(result.records) == 2
    assert len(result.duplicates) == 2
    assert all(len(r["Source IDs"]) == 2 for r in result.records)
    again = prepare_evidence(result.records + result.duplicates)
    assert len(again.records) == 2


@pytest.mark.parametrize(
    "change",
    [
        {"Account": "A-87654321"},
        {"Balance (£)": 999},
        {"Posting Date": "01/07/2024"},
        {"Fuel": "gas"},
    ],
)
def test_different_history_events_do_not_merge(tmp_path, change):
    html, pdf = _histories(tmp_path)
    pdf[0].update(change)
    assert len(prepare_evidence(html + pdf).records) == 3


def test_repeated_identical_payments_in_one_source_are_retained(tmp_path):
    html, pdf = _histories(tmp_path)
    repeated = deepcopy(html[1])
    repeated["Source Row"] = 3
    result = prepare_evidence([html[1], repeated, pdf[1]])
    assert len(result.records) == 3
    assert all("Repeated identical" in row["Reconciliation Review"] for row in result.records)


def test_sap_repeated_partial_page_tail_is_removed_without_losing_the_row():
    first = '"Account","Amount","Description","Status","Extra"\n"A-12345678",100,"No restriction","Cleared",""'
    second = 'restriction","Cleared",""\n"A-12345678",200,"Other","Open"'
    second += ',""'
    text, repairs = join_sap_pdf_pages([first, second])
    messages: list[str] = []
    rows = _parse_sap_csv(text, messages.append)
    assert repairs == [2]
    assert [row["Amount"] for row in rows] == ["100", "200"]
    assert messages == []


def test_sap_complete_repeated_rows_and_unverified_fragments_are_preserved():
    row = '"A-12345678",100,"Open"'
    assert join_sap_pdf_pages([row, row]) == (row + "\n" + row, [])
    assert join_sap_pdf_pages(["  " + row, "  " + row + "\n" + row])[1] == []
    assert join_sap_pdf_pages([row, 'unrelated","data",""\n' + row])[1] == []
