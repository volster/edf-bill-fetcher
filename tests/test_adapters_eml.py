"""Tests for the EML adapter — ``edf_bill_fetcher.io.adapters.eml``.

Pins the 5-key dict contract consumed by the engine's folder-ingestion
surface: ``sender`` / ``subject`` / ``date_str`` / ``body_html`` /
``body_text``, with a missing body yielding empty strings rather than
``None`` or exceptions.
"""

from pathlib import Path

from edf_bill_fetcher.io.adapters.eml import _decode_header_value, parse_eml_message

HTML_BODY = "<html><body><h1>Your EDF bill</h1></body></html>"


def test_eml_html_body_parsed(eml_html_path: Path) -> None:
    """Parse a single-part HTML message into the 5-key record dict."""
    result = parse_eml_message(eml_html_path)
    assert result == {
        "sender": "EDF Billing <billing@edfenergy.com>",
        "subject": "Your EDF bill is ready",
        "date_str": "15/01/2024",
        "body_html": HTML_BODY,
        "body_text": "",
    }


def test_eml_plain_body_parsed(eml_plain_path: Path) -> None:
    """Parse a single-part plain-text message, leaving body_html empty."""
    result = parse_eml_message(eml_plain_path)
    assert result == {
        "sender": "EDF Billing <billing@edfenergy.com>",
        "subject": "Your EDF bill",
        "date_str": "12/02/2024",
        "body_html": "",
        "body_text": "Your EDF bill for January is ready.\nTotal: £120.00\n",
    }


def test_eml_multipart_alternative_keeps_html_and_plain(eml_multipart_path: Path) -> None:
    """Collect the text/html and text/plain parts into their own fields."""
    result = parse_eml_message(eml_multipart_path)
    assert result["body_html"] == "<p>Rich <b>HTML</b></p>\n"
    assert result["body_text"] == "Plain fallback text\n"


def test_eml_missing_body_returns_empty_strings(eml_empty_path: Path) -> None:
    """Return empty strings for both body fields when no body content exists."""
    result = parse_eml_message(eml_empty_path)
    assert result == {
        "sender": "no-reply@edfenergy.com",
        "subject": "",
        "date_str": "15/01/2024",
        "body_html": "",
        "body_text": "",
    }


def test_eml_rfc2047_encoded_subject_decoded(eml_encoded_subject_path: Path) -> None:
    """Decode an RFC2047 base64-encoded Subject header to plain text."""
    result = parse_eml_message(eml_encoded_subject_path)
    assert result["subject"] == "Vive l'énergie"
    assert result["body_text"] == "Plain body"


def test_eml_date_header_formatted_via_display_date(eml_plain_path: Path) -> None:
    """Format the RFC2822 Date header through parse_to_display_date."""
    assert parse_eml_message(eml_plain_path)["date_str"] == "12/02/2024"


def test_decode_header_value_handles_bytes_chunks() -> None:
    """Decode mixed plain/encoded chunks into a single plain-text string."""
    assert _decode_header_value("=?utf-8?b?Vml2ZSBsJ8OpbmVyZ2ll?=") == "Vive l'énergie"
    assert _decode_header_value("Plain part =?utf-8?q?caf=C3=A9?=") == "Plain part café"
    assert _decode_header_value("") == ""
