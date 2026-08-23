"""Pytest configuration and shared fixtures.

The project root is added to ``sys.path`` automatically via
``pythonpath = ["."]`` in ``pyproject.toml``'s
``[tool.pytest.ini_options]`` block — no manual path manipulation
needed here.
"""

from email import policy
from email.message import EmailMessage
from pathlib import Path

import pytest


def _save_eml(tmp_path: Path, name: str, msg: EmailMessage) -> Path:
    """Serialize a synthetic EmailMessage to an .eml file under tmp_path."""
    path = tmp_path / name
    path.write_bytes(msg.as_bytes())
    return path


@pytest.fixture
def eml_html_path(tmp_path):
    """Path to a synthetic .eml with a single text/html body part."""
    msg = EmailMessage()
    msg["From"] = "EDF Billing <billing@edfenergy.com>"
    msg["To"] = "customer@example.com"
    msg["Subject"] = "Your EDF bill is ready"
    msg["Date"] = "Tue, 15 Jan 2024 10:30:00 +0000"
    msg["Content-Type"] = "text/html"
    msg.set_payload("<html><body><h1>Your EDF bill</h1></body></html>")
    return _save_eml(tmp_path, "bill_html.eml", msg)


@pytest.fixture
def eml_plain_path(tmp_path):
    """Path to a synthetic .eml with a single text/plain body part."""
    msg = EmailMessage()
    msg["From"] = "EDF Billing <billing@edfenergy.com>"
    msg["To"] = "customer@example.com"
    msg["Subject"] = "Your EDF bill"
    msg["Date"] = "Mon, 12 Feb 2024 09:00:00 +0000"
    msg.set_content("Your EDF bill for January is ready.\nTotal: £120.00\n")
    return _save_eml(tmp_path, "bill_plain.eml", msg)


@pytest.fixture
def eml_multipart_path(tmp_path):
    """Path to a synthetic .eml with a multipart/alternative body."""
    msg = EmailMessage()
    msg["From"] = "EDF Billing <billing@edfenergy.com>"
    msg["To"] = "customer@example.com"
    msg["Subject"] = "Your EDF bill"
    msg["Date"] = "Tue, 15 Jan 2024 10:30:00 +0000"
    msg.set_content("Plain fallback text")
    msg.add_alternative("<p>Rich <b>HTML</b></p>", subtype="html")
    return _save_eml(tmp_path, "bill_multipart.eml", msg)


@pytest.fixture
def eml_empty_path(tmp_path):
    """Path to a synthetic .eml with headers but no body content."""
    msg = EmailMessage()
    msg["From"] = "no-reply@edfenergy.com"
    msg["Subject"] = ""
    msg["Date"] = "Tue, 15 Jan 2024 10:30:00 +0000"
    return _save_eml(tmp_path, "empty.eml", msg)


@pytest.fixture
def eml_encoded_subject_path(tmp_path):
    """Path to a synthetic .eml whose Subject header is RFC2047-encoded.

    Built with the compat32 policy so the encoded subject survives
    serialization verbatim instead of being normalized by the default
    policy's generator.
    """
    msg = EmailMessage(policy=policy.compat32)
    msg["From"] = "EDF Billing <billing@edfenergy.com>"
    msg["Subject"] = "=?utf-8?b?Vml2ZSBsJ8OpbmVyZ2ll?="
    msg["Date"] = "Tue, 15 Jan 2024 10:30:00 +0000"
    msg["Content-Type"] = 'text/plain; charset="utf-8"'
    msg.set_payload("Plain body")
    return _save_eml(tmp_path, "encoded_subject.eml", msg)


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
