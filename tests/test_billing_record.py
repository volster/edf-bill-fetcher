"""Tests for the typed ``BillingRecord`` dataclass and its ``to_dict()`` boundary.

``to_dict()`` is the single producer of the 19-column evidence schema shared
by every record builder: ``None`` string fields map to ``"N/A"``,
``cancel_rebill_admitted=None`` maps to ``False``, and an explicit ``"N/A"``
string passes through unchanged.
"""

from __future__ import annotations

from edf_bill_fetcher.models.records import BillingRecord

CANONICAL_KEYS = [
    "Source",
    "Sender",
    "Date",
    "Period From",
    "Period To",
    "Invoice #",
    "Amount (£)",
    "Period Charge (£)",
    "Entry Type",
    "Reading",
    "Units (kWh)",
    "Standing Chg (p/day)",
    "Tariff",
    "Attachment Name",
    "Details",
    "Logic Used",
    "Source PDF Text",
    "_regex_trace",
    "Cancel/Rebill Admitted",
]


def _full_shape_record() -> BillingRecord:
    return BillingRecord(
        source="Local PDF Folder",
        sender="edf.co.uk",
        date="01 Aug 2025",
        period_from="01 Jul 2025",
        period_to="31 Jul 2025",
        invoice_num="KI-31105244-9999",
        amount="123.45",
        period_charge="678.90",
        entry_type="New Bill",
        reading="Actual",
        units_kwh="1234",
        standing_charge="25.00",
        tariff="N/A",
        attachment_name="test.pdf",
        details="New invoice",
        logic_used="New Invoice Format",
        source_pdf_text="Invoice body text",
        regex_trace="inv_num via _INV_NUMBER_RE; period_from via _BILLING_PERIOD_RE",
        cancel_rebill_admitted=True,
    )


def test_full_shape_record_emits_canonical_engine_dict():
    record = _full_shape_record()
    assert record.to_dict() == {
        "Source": "Local PDF Folder",
        "Sender": "edf.co.uk",
        "Date": "01 Aug 2025",
        "Period From": "01 Jul 2025",
        "Period To": "31 Jul 2025",
        "Invoice #": "KI-31105244-9999",
        "Amount (£)": "123.45",
        "Period Charge (£)": "678.90",
        "Entry Type": "New Bill",
        "Reading": "Actual",
        "Units (kWh)": "1234",
        "Standing Chg (p/day)": "25.00",
        "Tariff": "N/A",
        "Attachment Name": "test.pdf",
        "Details": "New invoice",
        "Logic Used": "New Invoice Format",
        "Source PDF Text": "Invoice body text",
        "_regex_trace": "inv_num via _INV_NUMBER_RE; period_from via _BILLING_PERIOD_RE",
        "Cancel/Rebill Admitted": True,
    }
    assert list(record.to_dict()) == CANONICAL_KEYS


def test_minimal_record_emits_canonical_key_set():
    record = BillingRecord(source="HTM Account History", entry_type="Charge", logic_used="HTM Regex")
    assert set(record.to_dict()) == set(CANONICAL_KEYS)


def test_none_fields_map_to_sentinels():
    record = BillingRecord(source="src", entry_type="Charge", logic_used="HTM Regex")
    result = record.to_dict()
    assert result["Source"] == "src"
    assert result["Entry Type"] == "Charge"
    assert result["Logic Used"] == "HTM Regex"
    for key in CANONICAL_KEYS[:-1]:
        if key not in ("Source", "Entry Type", "Logic Used"):
            assert result[key] == "N/A"
    assert result["Cancel/Rebill Admitted"] is False


def test_explicit_na_passes_through_unchanged():
    record = BillingRecord(
        source="src",
        entry_type="Charge",
        logic_used="HTM Regex",
        date="N/A",
        amount="N/A",
        period_charge="N/A",
    )
    result = record.to_dict()
    assert result["Date"] == "N/A"
    assert result["Amount (£)"] == "N/A"
    assert result["Period Charge (£)"] == "N/A"


def test_field_passthrough_to_display_keys():
    record = BillingRecord(
        source="x",
        entry_type="y",
        logic_used="z",
        amount="1.0",
        period_charge="N/A",
    )
    assert record.to_dict()["Amount (£)"] == "1.0"
    assert record.to_dict()["Period Charge (£)"] == "N/A"
