"""Coverage tests for the pure string-processing helpers in
``processors/extraction.py`` — fallback invoice / period / amount
extractors, PST-record helpers, and the domain-filter check.

Closes the 112-missed-line gap from 6% -> ~100%. The 5 string-processing
helpers are called directly with synthetic strings. The 2 PST-record
helpers use ``unittest.mock.MagicMock`` stubs that mimic the libpff record
attribute surface (``get_number_of_record_sets``, ``get_record_set``,
etc.) — no real .pst file is required.
"""

from __future__ import annotations

from collections.abc import Sequence
from unittest.mock import MagicMock

import pytest

from edf_bill_fetcher.processors.extraction import (
    _extract_sender_email,
    _fallback_amount,
    _fallback_inv_num,
    _fallback_period_from,
    _fallback_period_to,
    _matches_domain_filter,
    _pst_attachment_filename,
)

# ---------- _fallback_inv_num ----------


def test_fallback_inv_num_canonical_ki_invoice() -> None:
    """`Invoice number: KI-12345` matches `_INV_NUMBER_RE` first."""
    val, label = _fallback_inv_num("Invoice number: KI-12345")
    assert val == "KI-12345"
    assert label == "_INV_NUMBER_RE"


def test_fallback_inv_num_credit_note_kcr() -> None:
    """`Credit note number: KCR-67890` matches `_CREDIT_NUMBER_RE`."""
    val, label = _fallback_inv_num("Credit note number: KCR-67890")
    assert val == "KCR-67890"
    assert label == "_CREDIT_NUMBER_RE"


def test_fallback_inv_num_cover_block_invoice() -> None:
    """`Invoice number: A-12345678` matches `_COVER_BLOCK_INV_RE` (3rd in chain)."""
    val, label = _fallback_inv_num("Invoice number: A-12345678")
    assert val == "A-12345678"
    assert label == "_COVER_BLOCK_INV_RE"


def test_fallback_inv_num_no_match_returns_none() -> None:
    """Empty / no-match input returns `(None, "")`."""
    assert _fallback_inv_num("") == (None, "")
    assert _fallback_inv_num("just some random text with no invoice number") == (None, "")


def test_fallback_inv_num_with_whitespace_returns_none() -> None:
    """All-whitespace text returns no match."""
    val, label = _fallback_inv_num("     \n\t   ")
    assert val is None
    assert label == ""


# ---------- _fallback_period_from / _fallback_period_to ----------


def test_fallback_period_from_billing_period_match() -> None:
    """`Your charges: 01 Jan 2024 - 31 Jan 2024` matches `_BILLING_PERIOD_RE`."""
    val, label = _fallback_period_from("Your charges: 01 Jan 2024 - 31 Jan 2024")
    assert val == "01 Jan 2024"
    assert label == "_BILLING_PERIOD_RE"


def test_fallback_period_to_billing_period_match() -> None:
    """`Your charges: 01 Jan 2024 - 31 Jan 2024` matches `_BILLING_PERIOD_RE`."""
    val, label = _fallback_period_to("Your charges: 01 Jan 2024 - 31 Jan 2024")
    assert val == "31 Jan 2024"
    assert label == "_BILLING_PERIOD_RE"


def test_fallback_period_from_cover_block_period() -> None:
    """`for the period: 01 Jan 2024 - 31 Jan 2024` matches `_COVER_BLOCK_PERIOD_RE`."""
    val, label = _fallback_period_from("for the period: 01 Jan 2024 - 31 Jan 2024")
    assert val == "01 Jan 2024"
    assert label == "_COVER_BLOCK_PERIOD_RE"


def test_fallback_period_to_cover_block_period() -> None:
    """`for the period: 01 Jan 2024 - 31 Jan 2024` matches `_COVER_BLOCK_PERIOD_RE`."""
    val, label = _fallback_period_to("for the period: 01 Jan 2024 - 31 Jan 2024")
    assert val == "31 Jan 2024"
    assert label == "_COVER_BLOCK_PERIOD_RE"


def test_fallback_period_from_no_match_returns_none() -> None:
    """Text without the billing-period shape returns `(None, "")`."""
    assert _fallback_period_from("") == (None, "")
    assert _fallback_period_from("just some random text with no period") == (None, "")


def test_fallback_period_to_no_match_returns_none() -> None:
    """Text without the shape returns `(None, "")`."""
    assert _fallback_period_to("") == (None, "")
    assert _fallback_period_to("just some random text") == (None, "")


# ---------- _fallback_amount ----------


def test_fallback_amount_period_charge_match() -> None:
    """`Total charges for this period £123.45` matches `_PERIOD_CHARGE_RE`."""
    val, label = _fallback_amount("Total charges for this period £123.45")
    assert val == pytest.approx(123.45)
    assert label == "_PERIOD_CHARGE_RE"


def test_fallback_amount_credit_total_match() -> None:
    """`Total credits for this bill £500.00` matches `_CREDIT_TOTAL_RE`."""
    val, label = _fallback_amount("Total credits for this bill £500.00")
    assert val == pytest.approx(500.00)
    assert label == "_CREDIT_TOTAL_RE"


def test_fallback_amount_pound_amount_fallback_match() -> None:
    """Bare `£67.89` falls through to `_POUND_AMOUNT_FALLBACK_RE`."""
    val, label = _fallback_amount("Balance outstanding: £67.89")
    assert val == pytest.approx(67.89)
    assert label == "_POUND_AMOUNT_FALLBACK_RE"


def test_fallback_amount_no_match_returns_none() -> None:
    """Text without any of the three amount patterns returns `(None, "")`."""
    assert _fallback_amount("") == (None, "")
    assert _fallback_amount("just some text without a pound amount") == (None, "")


# ---------- _matches_domain_filter ----------


def test_matches_domain_filter_exact_domain_match() -> None:
    """`sender@edf.com` matches filter `edf.com`."""
    assert _matches_domain_filter("user@edf.com", "edf.com") is True


def test_matches_domain_filter_case_insensitive() -> None:
    """Filter is case-insensitive on both sender and filter."""
    assert _matches_domain_filter("user@EDF.COM", "edf.com") is True
    assert _matches_domain_filter("user@edf.com", "EDF.COM") is True


def test_matches_domain_filter_subdomain_wildcard() -> None:
    """Filter `edf.com` matches subdomains (`user@billing.edf.com`)."""
    assert _matches_domain_filter("user@billing.edf.com", "edf.com") is True


def test_matches_domain_filter_explicit_wildcard_subdomain() -> None:
    """Filter `*.edf.com` matches BOTH apex + subdomains (lstrip('*.') normalizes)."""
    assert _matches_domain_filter("user@billing.edf.com", "*.edf.com") is True
    assert _matches_domain_filter("user@edf.com", "*.edf.com") is True


def test_matches_domain_filter_full_email_address() -> None:
    """Full email address in filter matches only that sender."""
    assert _matches_domain_filter("billing@edf.com", "billing@edf.com") is True
    assert _matches_domain_filter("user@edf.com", "billing@edf.com") is False


def test_matches_domain_filter_multiple_domains_in_filter() -> None:
    """Comma-separated filter with multiple domains."""
    assert _matches_domain_filter("user@edf.com", "edf.com,other.com") is True
    assert _matches_domain_filter("user@other.com", "edf.com,other.com") is True
    assert _matches_domain_filter("user@third.com", "edf.com,other.com") is False


def test_matches_domain_filter_empty_sender_returns_false() -> None:
    """Empty sender email returns False."""
    assert _matches_domain_filter("", "edf.com") is False


def test_matches_domain_filter_empty_filter_returns_false() -> None:
    """Empty filter string returns False even for a populated sender."""
    assert _matches_domain_filter("user@edf.com", "") is False


def test_matches_domain_filter_no_at_in_sender_returns_false() -> None:
    """Sender without `@` returns False (no domain to match)."""
    assert _matches_domain_filter("just-a-string", "edf.com") is False


# ---------- _pst_attachment_filename ----------


def _make_pst_record_set(entries: Sequence[tuple[int, str | bytes | None]]) -> MagicMock:
    """Build a MagicMock mimicking a pypff record-set with the given entries."""
    rs = MagicMock()
    rs.get_number_of_entries.return_value = len(entries)
    entry_objs = []
    for entry_type, data in entries:
        e = MagicMock()
        e.entry_type = entry_type
        if data is None:
            e.get_data_as_string.side_effect = Exception("no string")
            e.get_data.side_effect = Exception("no bytes")
        elif isinstance(data, str):
            e.get_data_as_string.return_value = data
            e.get_data.return_value = data.encode("utf-16-le")
        else:
            e.get_data_as_string.side_effect = Exception("not string")
            e.get_data.return_value = data
        entry_objs.append(e)
    rs.get_entry.side_effect = entry_objs
    return rs


def _make_pst_attachment(record_sets: Sequence[MagicMock]) -> MagicMock:
    """Build a MagicMock mimicking a pypff.attachment containing record_sets."""
    att = MagicMock()
    att.get_number_of_record_sets.return_value = len(record_sets)
    att.get_record_set.side_effect = list(record_sets)
    return att


def test_pst_attachment_filename_returns_long_filename_for_matching_entry() -> None:
    """A record-set containing the PR_ATTACH_LONG_FILENAME entry returns the filename."""
    from edf_bill_fetcher.processors.patterns import _PST_PR_ATTACH_LONG_FILENAME

    rs = _make_pst_record_set([(_PST_PR_ATTACH_LONG_FILENAME, "billing_2024.pdf")])
    att = _make_pst_attachment([rs])
    result = _pst_attachment_filename(att)
    assert result == "billing_2024.pdf"


def test_pst_attachment_filename_returns_none_when_no_matching_entry() -> None:
    """A record-set with no PR_ATTACH_LONG_FILENAME entry returns None."""
    rs = _make_pst_record_set([(0x0001, "irrelevant")])
    att = _make_pst_attachment([rs])
    result = _pst_attachment_filename(att)
    assert result is None


def test_pst_attachment_filename_returns_none_for_none_input() -> None:
    """`None` attachment returns None."""
    assert _pst_attachment_filename(None) is None


def test_pst_attachment_filename_returns_none_when_count_getter_missing() -> None:
    """An attachment object with no `get_number_of_record_sets` returns None."""
    att = MagicMock(spec=[])
    result = _pst_attachment_filename(att)
    assert result is None


def test_pst_attachment_filename_returns_none_when_count_getter_raises() -> None:
    """An attachment whose `get_number_of_record_sets()` raises returns None."""
    att = MagicMock()
    att.get_number_of_record_sets.side_effect = Exception("boom")
    result = _pst_attachment_filename(att)
    assert result is None


def test_pst_attachment_filename_handles_bytes_via_utf16_decode() -> None:
    """When ``get_data_as_string`` returns empty/non-str, fall back to ``get_data`` bytes."""
    from edf_bill_fetcher.processors.patterns import _PST_PR_ATTACH_LONG_FILENAME

    rs = MagicMock()
    rs.get_number_of_entries.return_value = 1
    e = MagicMock()
    e.entry_type = _PST_PR_ATTACH_LONG_FILENAME
    # Empty string triggers the `if isinstance(val, str) and val` early-return path
    # to fall through to the bytes-decode fallback.
    e.get_data_as_string.return_value = ""
    e.get_data.return_value = "cloudexport.pdf".encode("utf-16-le")
    rs.get_entry.return_value = e
    att = MagicMock()
    att.get_number_of_record_sets.return_value = 1
    att.get_record_set.return_value = rs
    result = _pst_attachment_filename(att)
    assert result == "cloudexport.pdf"


# ---------- _extract_sender_email ----------


def test_extract_sender_email_from_transport_headers() -> None:
    """Sender email extracted from `From:` header in transport headers."""
    msg = MagicMock()
    msg.get_transport_headers.return_value = (
        "From: sender@edf.com\r\nTo: other@example.com\r\nSubject: bill"
    )
    result = _extract_sender_email(msg)
    assert result == "sender@edf.com"


def test_extract_sender_email_falls_back_to_sender_name_field() -> None:
    """When transport headers have no `From:` line, fall back to sender name field."""
    msg = MagicMock()
    msg.get_transport_headers.return_value = ""
    msg.get_sender_name.return_value = "Billing <billing@edfenergy.com>"
    result = _extract_sender_email(msg)
    assert result == "billing@edfenergy.com"


def test_extract_sender_email_returns_empty_when_no_email() -> None:
    """Returns empty string when neither headers nor sender name contains an email."""
    msg = MagicMock()
    msg.get_transport_headers.return_value = "Subject: blank"
    msg.get_sender_name.return_value = "Some Person Without Email"
    result = _extract_sender_email(msg)
    assert result == ""


def test_extract_sender_email_returns_empty_when_transport_headers_none() -> None:
    """Returns empty string when transport headers are None."""
    msg = MagicMock()
    msg.get_transport_headers.return_value = None
    msg.get_sender_name.return_value = ""
    result = _extract_sender_email(msg)
    assert result == ""


def test_extract_sender_email_handles_bytes_transport_headers() -> None:
    """Bytes transport headers are decoded and parsed."""
    transport_headers_bytes = b"From: bytes-case@edf.com\r\nTo: x@y.com"
    msg = MagicMock()
    msg.get_transport_headers.return_value = transport_headers_bytes
    msg.get_sender_name.return_value = ""
    result = _extract_sender_email(msg)
    assert result == "bytes-case@edf.com"
