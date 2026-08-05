"""Branch-coverage tests for ``edf_bill_fetcher.io.adapters.pst``.

The module under test ships three public helpers -- ``pst_attachment_filename``,
``extract_sender_email``, and ``matches_domain_filter`` -- each guarded by a
defensive ``try/except`` ladder that tolerates a missing or version-mismatched
``pypff`` library.  These tests exercise every branch of that ladder using the
synthetic ``pypff`` shapes defined in ``tests/fixtures/pst_attachment_fixture.py``.

No real ``.pst`` file is touched; every input is a synthetic fake.
"""

from __future__ import annotations

from edf_bill_fetcher.io.adapters.pst import (
    extract_sender_email,
    matches_domain_filter,
    pst_attachment_filename,
)
from tests.fixtures.pst_attachment_fixture import (
    PR_ATTACH_LONG_FILENAME,
    PT_UNICODE,
    FakeAttachment,
    FakeAttachmentNoRecordSetGetter,
    FakeAttachmentRecordSetCountRaises,
    FakeMessage,
    FakeRecordEntry,
    FakeRecordSet,
    FakeRecordSetNoEntriesGetter,
    make_att_with_bad_entry_type,
    make_att_with_empty_raw_bytes_long_filename,
    make_att_with_empty_string_get_data_raises,
    make_att_with_empty_string_non_bytes_data,
    make_att_with_get_data_raises,
    make_att_with_get_record_set_raises,
    make_att_with_long_filename,
    make_att_with_no_filename_entries_v2,
    make_att_with_only_short_filename,
    make_att_with_raw_bytes_long_filename,
    make_att_with_record_set_entries_getter_raises,
    make_att_with_record_set_get_entry_raises,
    make_att_with_record_set_no_entries_getter,
    make_multiple_attachments_with_long_filenames,
)

# ---------------------------------------------------------------------------
# pst_attachment_filename -- lines 55-104
# ---------------------------------------------------------------------------


def test_none_attachment_returns_none() -> None:
    """``att is None`` short-circuits to ``None`` (line 55-56)."""
    assert pst_attachment_filename(None) is None


def test_attachment_without_record_set_getter_returns_none() -> None:
    """Missing ``get_number_of_record_sets`` -> ``getattr`` returns ``None`` (line 57-59)."""
    att = FakeAttachmentNoRecordSetGetter()
    assert pst_attachment_filename(att) is None


def test_record_set_count_raises_returns_none() -> None:
    """``int(getter_count())`` raising -> ``None`` (line 60-63)."""
    att = FakeAttachmentRecordSetCountRaises()
    assert pst_attachment_filename(att) is None


def test_get_record_set_raises_continues_and_returns_none() -> None:
    """Every ``get_record_set(i)`` raises -> loop ``continue``s, then ``None`` (line 64-68, 104)."""
    att = make_att_with_get_record_set_raises()
    assert pst_attachment_filename(att) is None


def test_record_set_without_entries_getter_skipped() -> None:
    """Record-set missing ``get_number_of_entries`` -> ``continue`` (line 69-71).

    A second valid record-set carries the long filename, proving the broken
    set is skipped rather than aborting the whole walk.
    """
    att = make_att_with_record_set_no_entries_getter()
    assert pst_attachment_filename(att) == "good.pdf"


def test_entries_getter_raises_skips_record_set() -> None:
    """``int(entries_getter())`` raising -> ``continue`` (line 72-75)."""
    att = make_att_with_record_set_entries_getter_raises()
    assert pst_attachment_filename(att) == "good.pdf"


def test_get_entry_raises_continues_to_return_none() -> None:
    """``get_entry(j)`` raising for every index -> ``continue`` (line 76-80, 104)."""
    att = make_att_with_record_set_get_entry_raises()
    assert pst_attachment_filename(att) is None


def test_bad_entry_type_skipped() -> None:
    """``int(entry.entry_type)`` raising -> ``continue`` (line 81-84, 104)."""
    att = make_att_with_bad_entry_type()
    assert pst_attachment_filename(att) is None


def test_non_matching_entry_type_skipped() -> None:
    """``entry_type != PR_ATTACH_LONG_FILENAME`` -> ``continue`` (line 85-86, 104)."""
    att = make_att_with_no_filename_entries_v2("display only")
    assert pst_attachment_filename(att) is None


def test_only_short_filename_returns_none() -> None:
    """Short-filename tag alone does not satisfy the long-filename check.

    The helper looks exclusively for ``PR_ATTACH_LONG_FILENAME`` (0x3707);
    a short-filename-only attachment yields ``None`` and the caller is
    expected to fall back to ``Attachment_N.pdf`` (line 85-86, 104).
    """
    att = make_att_with_only_short_filename("~ingest.pdf")
    assert pst_attachment_filename(att) is None


def test_get_data_as_string_returns_filename() -> None:
    """Happy path: ``get_data_as_string`` returns a non-empty str (line 87-92)."""
    att = make_att_with_long_filename("edf-invoice-KI-31105244.pdf")
    assert pst_attachment_filename(att) == "edf-invoice-KI-31105244.pdf"


def test_get_data_as_string_raises_falls_through_to_none() -> None:
    """``get_data_as_string`` raises -> ``continue`` (line 87-90, 93-96, 104).

    Both data getters raise on this entry, so the helper exhausts the
    record-set and returns ``None``.
    """
    att = make_att_with_get_data_raises()
    assert pst_attachment_filename(att) is None


def test_empty_string_then_get_data_raises_returns_none() -> None:
    """Empty ``get_data_as_string`` -> ``get_data`` raises -> ``continue`` (line 91-96, 104).

    The str branch at line 91 is falsy (empty string), so the helper falls
    through to ``get_data()``, which raises and triggers the second
    ``except`` at lines 95-96.
    """
    att = make_att_with_empty_string_get_data_raises()
    assert pst_attachment_filename(att) is None


def test_empty_string_non_bytes_data_returns_none() -> None:
    """Empty str + non-bytes ``get_data`` -> isinstance falsy -> loop (line 97->76, 104).

    ``get_data()`` returns an ``int``; the ``isinstance(raw_data, bytes | bytearray)``
    check at line 97 is False, so the inner loop iterates and the helper
    ultimately returns ``None``.
    """
    att = make_att_with_empty_string_non_bytes_data()
    assert pst_attachment_filename(att) is None


def test_raw_bytes_utf16le_decoded_returned() -> None:
    """Empty ``get_data_as_string`` -> raw-bytes -> UTF-16LE decode -> return (line 93-103)."""
    att = make_att_with_raw_bytes_long_filename("edf-bill-2026.pdf")
    assert pst_attachment_filename(att) == "edf-bill-2026.pdf"


def test_raw_bytes_all_nuls_falls_through_to_none() -> None:
    """Decoded result is all NULs -> ``strip`` falsy -> fall through (line 102-104)."""
    att = make_att_with_empty_raw_bytes_long_filename()
    assert pst_attachment_filename(att) is None


def test_multiple_attachments_each_yield_filename() -> None:
    """Multiple attachments in a batch each resolve independently.

    Mirrors the ``EvidenceEngine.crawl_pst`` loop shape: the helper is
    called once per attachment and each returns its own filename.
    """
    names = ["a.pdf", "b.pdf", "c.pdf"]
    attachments = make_multiple_attachments_with_long_filenames(names)
    results = [pst_attachment_filename(att) for att in attachments]
    assert results == names


def test_corrupt_record_set_attribute_error_swallowed() -> None:
    """A record-set whose getters raise ``AttributeError`` is swallowed, not propagated.

    The helper's ``except Exception`` ladders catch ``AttributeError`` from a
    corrupt record-set (missing method, broken entry) and ``continue`` to the
    next set rather than crashing the crawl.
    """
    corrupt_rs = FakeRecordSetNoEntriesGetter()
    good_entry = FakeRecordEntry(PR_ATTACH_LONG_FILENAME, PT_UNICODE, "survivor.pdf")
    good_rs = FakeRecordSet([good_entry])
    att = FakeAttachment([corrupt_rs, good_rs])  # type: ignore[list-item]
    assert pst_attachment_filename(att) == "survivor.pdf"


# ---------------------------------------------------------------------------
# extract_sender_email -- lines 107-129
# ---------------------------------------------------------------------------


def test_extract_sender_email_from_str_transport_headers() -> None:
    """``From:`` header in a ``str`` transport-headers blob (line 110-118)."""
    headers = "From: Billing Dept <billing@edf.com>\r\nSubject: Invoice"
    msg = FakeMessage(transport_headers=headers)
    assert extract_sender_email(msg) == "billing@edf.com"


def test_extract_sender_email_from_bytes_transport_headers() -> None:
    """Transport headers returned as ``bytes`` are decoded then searched (line 113-118)."""
    headers = b"From: auto@edfenergy.com\r\nDate: Mon"
    msg = FakeMessage(transport_headers=headers)
    assert extract_sender_email(msg) == "auto@edfenergy.com"


def test_extract_sender_email_falls_back_to_sender_name() -> None:
    """No transport headers -> sender-name regex fallback (line 121-128)."""
    msg = FakeMessage(transport_headers=None, sender_name="noreply@edf.com")
    assert extract_sender_email(msg) == "noreply@edf.com"


def test_extract_sender_email_transport_headers_without_match_falls_back() -> None:
    """Headers present but no ``From:`` email -> sender-name fallback (line 121-128)."""
    headers = "Subject: no from header here"
    msg = FakeMessage(transport_headers=headers, sender_name="billing@edf.com")
    assert extract_sender_email(msg) == "billing@edf.com"


def test_extract_sender_email_no_headers_no_sender_name_returns_empty() -> None:
    """Both sources absent -> empty string (line 129)."""
    msg = FakeMessage(transport_headers=None, sender_name=None)
    assert extract_sender_email(msg) == ""


def test_extract_sender_email_transport_raises_falls_back_to_sender_name() -> None:
    """``get_transport_headers`` raising -> sender-name fallback (line 119-128)."""
    msg = FakeMessage(
        transport_headers=None,
        sender_name="recovered@edf.com",
        transport_raises=True,
    )
    assert extract_sender_email(msg) == "recovered@edf.com"


def test_extract_sender_email_both_raise_returns_empty() -> None:
    """Both getters raise -> empty string (line 119-120, 127-128, 129)."""
    msg = FakeMessage(transport_raises=True, sender_raises=True)
    assert extract_sender_email(msg) == ""


def test_extract_sender_email_sender_name_without_email_returns_empty() -> None:
    """Sender name carries no parseable email -> empty string (line 124-128, 129)."""
    msg = FakeMessage(transport_headers=None, sender_name="EDF Billing Department")
    assert extract_sender_email(msg) == ""


# ---------------------------------------------------------------------------
# matches_domain_filter -- lines 132-154
# ---------------------------------------------------------------------------


def test_matches_domain_filter_empty_sender_returns_false() -> None:
    """Empty sender_email -> ``False`` (line 141-142)."""
    assert matches_domain_filter("", "edf.com") is False


def test_matches_domain_filter_empty_filter_returns_false() -> None:
    """Empty filter_str -> ``False`` (line 141-142)."""
    assert matches_domain_filter("a@edf.com", "") is False


def test_matches_domain_filter_full_address_match() -> None:
    """Pattern with ``@`` matches the full address exactly (line 145-148)."""
    assert matches_domain_filter("billing@edf.com", "billing@edf.com") is True


def test_matches_domain_filter_full_address_no_match() -> None:
    """Pattern with ``@`` that does not equal the sender -> no match (line 145-148, 154)."""
    assert matches_domain_filter("billing@edf.com", "other@edf.com") is False


def test_matches_domain_filter_domain_exact_match() -> None:
    """Bare domain matches the sender's exact domain (line 149-152)."""
    assert matches_domain_filter("user@edf.com", "edf.com") is True


def test_matches_domain_filter_subdomain_match() -> None:
    """Bare domain matches a subdomain via ``endswith`` (line 149-152)."""
    assert matches_domain_filter("user@billing.edf.com", "edf.com") is True


def test_matches_domain_filter_wildcard_domain_match() -> None:
    """``*.edf.com`` pattern strips the wildcard then matches subdomains (line 149-152)."""
    assert matches_domain_filter("user@sub.edf.com", "*.edf.com") is True


def test_matches_domain_filter_no_domain_in_sender_returns_false() -> None:
    """Sender without ``@`` -> empty sender_domain -> no match (line 151, 154)."""
    assert matches_domain_filter("no-at-sign", "edf.com") is False


def test_matches_domain_filter_no_pattern_matches_returns_false() -> None:
    """No pattern in the comma-separated list matches (line 154)."""
    assert matches_domain_filter("user@edf.com", "other.com, *.example.org") is False


def test_matches_domain_filter_comma_separated_first_match_wins() -> None:
    """Comma-separated list with a match in the first entry (line 144-148)."""
    assert matches_domain_filter("billing@edf.com", "billing@edf.com, other.com") is True


def test_matches_domain_filter_comma_separated_second_match_wins() -> None:
    """Comma-separated list with a match in a later entry (line 144-153)."""
    assert matches_domain_filter("user@edf.com", "other.com, edf.com") is True


def test_matches_domain_filter_case_insensitive() -> None:
    """Both sides are lowercased before comparison (line 143-144)."""
    assert matches_domain_filter("User@EDF.COM", "EDF.com") is True


def test_matches_domain_filter_strips_whitespace_in_parts() -> None:
    """Whitespace around comma-separated parts is stripped (line 144)."""
    assert matches_domain_filter("user@edf.com", "  edf.com  ,  other.com") is True


def test_matches_domain_filter_empty_parts_ignored() -> None:
    """Empty parts from trailing/leading commas are dropped (line 144)."""
    assert matches_domain_filter("user@edf.com", ", edf.com ,") is True
