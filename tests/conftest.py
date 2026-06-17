"""Pytest configuration and shared fixtures."""

import sys

sys.path.insert(0, "C:/Users/matthew/edf-bill-fetcher")

import pytest


@pytest.fixture
def sample_new_invoice_text():
    """Sample text from a new-style KI invoice."""
    return """
    Invoice number: KI-12345678
    Account number: A-12345678
    Date issued: 15 Jan 2024
    Your charges: 01 Jan 2024 - 31 Jan 2024
    Current balance £1,234.56 debit
    Total charges for this period £500.00 debit
    Electricity used 1,234 kWh
    Standing charge 31 days @ 45.5p/day
    Tariff name Standard Variable Payment type
    """


@pytest.fixture
def sample_new_credit_text():
    """Sample text from a new-style KCR credit note."""
    return """
    Credit note number: KCR-12345678
    Account number: A-12345678
    Date issued: 15 Jan 2024
    Total credits for this bill £250.00
    """


@pytest.fixture
def sample_htm_text():
    """Sample HTM account history text."""
    return """
    28 Feb 2026 We charged your account £1,070.48 For 2354 kWh of electricity used between 01 Feb 2026 and 28 Feb 2026 Balance £46,182.13 in debit
    27 Feb 2026 You paid us £850.00 Bank Transfer Balance £45,111.65 in debit
    26 Feb 2026 Reversed account charge £100.00 Refund Balance £44,011.65 in debit
    """


@pytest.fixture
def sample_config():
    """Default test configuration."""
    return {
        "use_anchors": True,
        "use_large": True,
        "use_reading_classification": True,
        "use_pdf_fields": True,
        "use_acc_filter": False,
        "acc_num": "",
        "min_amount": 500.0,
        "analysis_min": 500.0,
        "filter_below": True,
        "save_filtered": True,
        "use_dedup": True,
        "save_dups": True,
        "use_domain_filter": True,
        "domain_filter": "edfenergy.com",
    }
