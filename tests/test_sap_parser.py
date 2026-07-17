"""Tests for the SAP dump detector and CSV-in-PDF parsers.

Verifies the three SAP-data-dump parsers (contract history, meter-read
history, financial transactions) plus the header-row detector. Real-world
data lives in
``scratch/input/pdfs/June 2026 download from ombudsman/``; tests use
synthetic CSV text shaped like the real dumps and one happy-path
``detect_sap_dump`` call against the actual PDFs (when available on
this environment).
"""

import os

import pytest

from edf_collector import (
    detect_sap_dump,
    parse_sap_contract_history,
    parse_sap_financial_transactions,
    parse_sap_meter_read_history,
)

# ---------------------------------------------------------------------------
# Synthetic CSV-with-quoted-fields — same shape as the real SAP dumps.
# ---------------------------------------------------------------------------

CONTRACT_CSV = (
    '"Kraken ID","SAP Account number","Business Partner","Contract",'
    '"Account Determination ID","Fuel type","Start Date","End Date",'
    '"Cancelled Flag","Product","Product Description","Contract Reason",'
    '"Replication lock","Contract End Reason","Channel",'
    '"Cancellation Party","Cancellation Reason","Creation Date & Time",'
    '"Created by","Changed Date & Time","Changed by"\n'
    '"A-31105244","671078701920","0159628206","2011040650","SME",'
    '"Electricity","03-06-2016","30-09-2017","",'
    '"ESC1_FIXED_1B","Fixed for Business 59 to Sep 17",'
    '"Acquisition (New)","","Broker","","","",'
    '"19-05-2016 08:42:49","KNOTT1G",'
    '"30-09-2017 21:14:45","SAPBATCH"\n'
    '"A-31105244","671078701920","0159628206","2011040650","SME",'
    '"Electricity","01-04-2022","31-03-2023","",'
    '"ESC1_RENEWAL_1_ELEC","EasyFix for Business 1 year",'
    '"Renewal","","","","","",'
    '"31-03-2022 19:22:11","SAPBATCH",'
    '"31-03-2023 21:10:52","SAPBATCH"\n'
)


METER_CSV = (
    '"Kraken ID","SAP Account Number","Contract","Fuel Type",'
    '"Meter Read Reason","Scheduled Meter Read Date","Register",'
    '"Meter Read","Unit Of Measurment","Meter Reading Active",'
    '"Meter Read Status","Scheduled Meter Read Category",'
    '"Meter Read Type","Meter Read Category","Indicator",'
    '"Meter Read Date","Created On","Changed On"\n'
    '"A-31105244","671078701920","2011040650","Elec",'
    '"Meter reading at move-in","03-06-2016","001",31264,"KWH","Active",'
    '"Released by Agent",'
    '"Meter reading by utility company",'
    '"Deemed (Settlement Registers) or Estimat",'
    '"Meter reading by utility company","Read validated by DC",'
    '"03-06-2016","13-06-2016","01-07-2016"\n'
    '"A-31105244","671078701920","2011040650","Elec",'
    '"Periodic Meter Reading","14-07-2016","001",34732,"KWH","Active",'
    '"Billable",'
    '"Meter reading by the customer",'
    '"Automatic estimation - SAP","Automatic estimation","NA",'
    '"14-07-2016","20-06-2016","16-07-2016"\n'
)


FINANCIAL_CSV = (
    '"Kraken ID","SAP Account Number","Business Partner",'
    '"Account Determination ID","Contract","Fuel Type","Document No.",'
    '"Item","Sub Item","Payment Method","Document Date","Posting Date ",'
    '"Net Due Date","Clearing Status","Main Transactions",'
    '"Sub Transactions","Transaction Text","Amount",'
    '"Down Payment Flag","Statistical Key Flag","Clearing Document",'
    '"Clearing Date","Clearing Reason","Clearing Posting Date",'
    '"Clearing Amount","Restriction","Document Type",'
    '"Document Type Description","Tax Code","Tax Code Description",'
    '"G/L Account","G/L Description","Deferral Date"\n'
    '"A-31105244","671078701920","0159628206","Non-residential customers",'
    '"2011040650","Electricity","551000421040","1","0","",'
    '"18-07-2016","18-07-2016","21-07-2016","Cleared Item","0100","0020",'
    '"Dr- Consum Billing Receivable",436,'
    '"No","NA","376001212905","26-03-2020","Automatic Clearing",'
    '"26-03-2020",436,"No restriction","IN","Energy Invoicing","A4",'
    '"Donations or payment for equity funds",'
    '"0000210251","Billed Debtor SME Elec",""\n'
    '"A-31105244","671078701920","0159628206","Non-residential customers",'
    '"2011040650","Electricity","307000019853","1","0","",'
    '"28-07-2016","28-07-2016","24-08-2016","Cleared Item","0010","0022",'
    '"Dr- Late Payment Charges",12,'
    '"No","NA","376001212905","26-03-2020","Automatic Clearing",'
    '"26-03-2020",12,"No restriction","DM","Miscellaneous Debits","AG",'
    '"Payable Only after Budget Billing Request",'
    '"0000210251","Billed Debtor SME Elec","17-10-2019"\n'
)


# ---------------------------------------------------------------------------
# detect_sap_dump
# ---------------------------------------------------------------------------


class TestDetectSapDump:
    def test_contract_text_returns_contract(self) -> None:
        assert detect_sap_dump(CONTRACT_CSV) == "contract"

    def test_meter_text_returns_meter_read(self) -> None:
        assert detect_sap_dump(METER_CSV) == "meter_read"

    def test_financial_text_returns_financial(self) -> None:
        assert detect_sap_dump(FINANCIAL_CSV) == "financial"

    def test_regular_invoice_text_returns_none(self) -> None:
        text = """
        Invoice number: KI-31105244-0001
        Account number: A-31105244
        Date issued: 15 Jan 2024
        Your charges: 01 Jan 2024 - 31 Jan 2024
        Current balance £1,234.56 debit
        Total charges for this period £500.00 debit
        """
        assert detect_sap_dump(text) is None

    def test_empty_text_returns_none(self) -> None:
        assert detect_sap_dump("") is None

    def test_text_missing_header_returns_none(self) -> None:
        assert detect_sap_dump("Some random CSV, columns, here\nrow1,row2") is None

    def test_header_present_but_unknown_columns_returns_none(self) -> None:
        # Header row exists, but none of the three column signatures match
        text = (
            '"Kraken ID","SAP Account Number","Other Column"\n"A-31105244","671078701920","value"\n'
        )
        assert detect_sap_dump(text) is None


# ---------------------------------------------------------------------------
# parse_sap_contract_history
# ---------------------------------------------------------------------------


class TestParseSapContractHistory:
    def test_returns_dicts_one_per_data_row(self) -> None:
        rows = parse_sap_contract_history(CONTRACT_CSV)
        assert len(rows) == 2

    def test_first_row_has_iso_dates(self) -> None:
        rows = parse_sap_contract_history(CONTRACT_CSV)
        assert rows[0]["Contract From"] == "2016-06-03"
        assert rows[0]["Contract To"] == "2017-09-30"

    def test_product_code_and_description_preserved(self) -> None:
        rows = parse_sap_contract_history(CONTRACT_CSV)
        assert rows[0]["Product Code"] == "ESC1_FIXED_1B"
        assert rows[0]["Product Description"] == "Fixed for Business 59 to Sep 17"

    def test_contract_reason_preserved(self) -> None:
        rows = parse_sap_contract_history(CONTRACT_CSV)
        assert rows[0]["Contract Reason"] == "Acquisition (New)"
        assert rows[1]["Contract Reason"] == "Renewal"

    def test_set_up_by_from_created_by_column(self) -> None:
        rows = parse_sap_contract_history(CONTRACT_CSV)
        assert rows[0]["Set Up By"] == "KNOTT1G"
        assert rows[1]["Set Up By"] == "SAPBATCH"

    def test_cancelled_flag_preserved(self) -> None:
        rows = parse_sap_contract_history(CONTRACT_CSV)
        assert rows[0]["Cancelled Flag"] == ""

    def test_source_file_propagated(self) -> None:
        rows = parse_sap_contract_history(CONTRACT_CSV, source_file="abc.pdf")
        assert rows[0]["Source File"] == "abc.pdf"

    def test_empty_text_returns_empty_list(self) -> None:
        assert parse_sap_contract_history("") == []


# ---------------------------------------------------------------------------
# parse_sap_meter_read_history
# ---------------------------------------------------------------------------


class TestParseSapMeterReadHistory:
    def test_returns_dicts_one_per_data_row(self) -> None:
        rows = parse_sap_meter_read_history(METER_CSV)
        assert len(rows) == 2

    def test_first_actual_reading_marks_A(self) -> None:
        rows = parse_sap_meter_read_history(METER_CSV)
        # Row 0 = "Released by Agent" → Read Type "A"
        assert rows[0]["Read Type"] == "A"
        assert rows[0]["Reading (kWh)"] == "31264"

    def test_second_estimated_reading_marks_E(self) -> None:
        rows = parse_sap_meter_read_history(METER_CSV)
        # Row 1 = "Billable" + "Automatic estimation - SAP" → Read Type "E"
        assert rows[1]["Read Type"] == "E"

    def test_scheduled_read_date_iso(self) -> None:
        rows = parse_sap_meter_read_history(METER_CSV)
        assert rows[0]["Scheduled Read Date"] == "2016-06-03"
        assert rows[1]["Scheduled Read Date"] == "2016-07-14"

    def test_meter_read_date_iso(self) -> None:
        rows = parse_sap_meter_read_history(METER_CSV)
        assert rows[0]["Meter Read Date"] == "2016-06-03"
        assert rows[1]["Meter Read Date"] == "2016-07-14"

    def test_register_preserved(self) -> None:
        rows = parse_sap_meter_read_history(METER_CSV)
        assert rows[0]["Register"] == "001"

    def test_source_file_propagated(self) -> None:
        rows = parse_sap_meter_read_history(METER_CSV, source_file="meter.pdf")
        assert rows[0]["Source File"] == "meter.pdf"


# ---------------------------------------------------------------------------
# parse_sap_financial_transactions
# ---------------------------------------------------------------------------


class TestParseSapFinancialTransactions:
    def test_returns_dicts_one_per_data_row(self) -> None:
        rows = parse_sap_financial_transactions(FINANCIAL_CSV)
        assert len(rows) == 2

    def test_document_no_preserved(self) -> None:
        rows = parse_sap_financial_transactions(FINANCIAL_CSV)
        assert rows[0]["Document No."] == "551000421040"
        assert rows[1]["Document No."] == "307000019853"

    def test_document_date_iso(self) -> None:
        rows = parse_sap_financial_transactions(FINANCIAL_CSV)
        assert rows[0]["Document Date"] == "2016-07-18"
        assert rows[1]["Document Date"] == "2016-07-28"

    def test_posting_date_iso_with_trailing_space_in_header(self) -> None:
        # The real header has "Posting Date " with a trailing space
        rows = parse_sap_financial_transactions(FINANCIAL_CSV)
        assert rows[0]["Posting Date"] == "2016-07-18"
        assert rows[1]["Posting Date"] == "2016-07-28"

    def test_amount_preserved_as_string(self) -> None:
        rows = parse_sap_financial_transactions(FINANCIAL_CSV)
        assert rows[0]["Amount"] == "436"
        assert rows[1]["Amount"] == "12"

    def test_main_transaction_preserved(self) -> None:
        rows = parse_sap_financial_transactions(FINANCIAL_CSV)
        assert rows[0]["Main Transaction"] == "0100"
        assert rows[1]["Main Transaction"] == "0010"

    def test_transaction_text_preserved(self) -> None:
        rows = parse_sap_financial_transactions(FINANCIAL_CSV)
        assert rows[0]["Transaction Text"] == "Dr- Consum Billing Receivable"
        assert rows[1]["Transaction Text"] == "Dr- Late Payment Charges"

    def test_clearing_status_preserved(self) -> None:
        rows = parse_sap_financial_transactions(FINANCIAL_CSV)
        assert rows[0]["Clearing Status"] == "Cleared Item"

    def test_clearing_date_iso(self) -> None:
        rows = parse_sap_financial_transactions(FINANCIAL_CSV)
        assert rows[0]["Clearing Date"] == "2020-03-26"

    def test_document_type_preserved(self) -> None:
        rows = parse_sap_financial_transactions(FINANCIAL_CSV)
        assert rows[0]["Document Type"] == "IN"
        assert rows[0]["Document Type Description"] == "Energy Invoicing"
        assert rows[1]["Document Type"] == "DM"
        assert rows[1]["Document Type Description"] == "Miscellaneous Debits"

    def test_source_file_propagated(self) -> None:
        rows = parse_sap_financial_transactions(FINANCIAL_CSV, source_file="fin.pdf")
        assert rows[0]["Source File"] == "fin.pdf"

    def test_empty_text_returns_empty_list(self) -> None:
        assert parse_sap_financial_transactions("") == []


# ---------------------------------------------------------------------------
# Real-world smoke test against the live ombudsman PDFs when available.
# This test is skipped automatically if the PDF files are missing.
# ---------------------------------------------------------------------------

OMBUDSMAN_DIR = (
    "/mnt/c/users/matthew/wsl/edf-bill-fetcher/scratch/input/pdfs/June 2026 download from ombudsman"
)


def _load_pdf_text(path: str) -> str:
    import pdfplumber

    parts: list[str] = []
    with pdfplumber.open(path) as pdf:
        for p in pdf.pages:
            parts.append(p.extract_text() or "")
    return "\n".join(parts)


@pytest.mark.skipif(
    not os.path.isdir(OMBUDSMAN_DIR),
    reason="Ombudsman PDFs not present in scratch/input/pdfs/",
)
class TestRealOmbudsmanPdfs:
    def test_detect_contract_pdf(self) -> None:
        path = os.path.join(OMBUDSMAN_DIR, "671078701920_Contract-and-Product-Change-History.pdf")
        if not os.path.exists(path):
            pytest.skip("missing file: " + path)
        text = _load_pdf_text(path)
        assert detect_sap_dump(text) == "contract"

    def test_detect_meter_pdf(self) -> None:
        path = os.path.join(OMBUDSMAN_DIR, "671078701920_Meter-Read-History.pdf")
        if not os.path.exists(path):
            pytest.skip("missing file: " + path)
        text = _load_pdf_text(path)
        assert detect_sap_dump(text) == "meter_read"

    def test_detect_financial_pdf(self) -> None:
        path = os.path.join(OMBUDSMAN_DIR, "671078701920_Financial-Transactions.pdf")
        if not os.path.exists(path):
            pytest.skip("missing file: " + path)
        text = _load_pdf_text(path)
        assert detect_sap_dump(text) == "financial"

    def test_parse_contract_pdf_finds_six_rows(self) -> None:
        path = os.path.join(OMBUDSMAN_DIR, "671078701920_Contract-and-Product-Change-History.pdf")
        if not os.path.exists(path):
            pytest.skip("missing file: " + path)
        text = _load_pdf_text(path)
        rows = parse_sap_contract_history(text, source_file="contract.pdf")
        # Real file has exactly 6 contracts (per spec exploration)
        assert len(rows) == 6
        assert rows[0]["Product Code"] == "ESC1_FIXED_1B"
        assert rows[-1]["Product Code"] == "ESC1_FREEDOM"

    def test_parse_meter_pdf_has_many_rows(self) -> None:
        path = os.path.join(OMBUDSMAN_DIR, "671078701920_Meter-Read-History.pdf")
        if not os.path.exists(path):
            pytest.skip("missing file: " + path)
        text = _load_pdf_text(path)
        rows = parse_sap_meter_read_history(text, source_file="meter.pdf")
        assert len(rows) >= 20, f"expected many rows, got {len(rows)}"
