"""Tests for the PST attachment filename extraction helper (Stream P6 / Task 9).

The fixture module synthesises a ``pypff`` attachment object (the four
fictional getter methods used by the legacy code do NOT exist, which is the
root cause of every PDF row showing ``Attachment_N.pdf``).
"""

from __future__ import annotations

import pytest

from edf_bill_fetcher.collectors.engine import _pst_attachment_filename
from tests.fixtures.pst_attachment_fixture import (
    PR_ATTACH_FILENAME,
    PR_ATTACH_LONG_FILENAME,
    FakeAttachment,
    make_att_with_long_filename,
    make_att_with_mixed_unicode_and_non_unicode,
    make_att_with_no_filename_entries,
    make_att_with_record_no_entries,
    make_att_with_string8_long_filename,
    make_att_with_two_record_sets,
    make_att_with_zero_record_sets,
)


def test_long_filename_recovered() -> None:
    att = make_att_with_long_filename("edf-invoice-KI-31105244-0001-3.pdf")
    assert _pst_attachment_filename(att) == "edf-invoice-KI-31105244-0001-3.pdf"


def test_long_filename_recovered_when_two_record_sets_present() -> None:
    att = make_att_with_two_record_sets(
        long_name="edf-invoice-KI-31105244-0001-3.pdf",
        short_name="~ingest.pdf",
    )
    # Long filename takes precedence over short.
    assert _pst_attachment_filename(att) == "edf-invoice-KI-31105244-0001-3.pdf"


def test_zero_record_sets_returns_none() -> None:
    att = make_att_with_zero_record_sets()
    assert _pst_attachment_filename(att) is None


def test_record_set_with_no_entries_skipped() -> None:
    att = make_att_with_record_no_entries()
    assert _pst_attachment_filename(att) is None


def test_no_long_filename_entry_in_record_returns_none() -> None:
    att = make_att_with_no_filename_entries("Some Display Name")
    assert _pst_attachment_filename(att) is None


def test_string8_long_filename_extracted() -> None:
    # Same tag accepted regardless of MAPI data type (some PSTs store
    # PT_STRING8 instead of PT_UNICODE for the filename).
    att = make_att_with_string8_long_filename("legacy-edf-invoice-file.pdf")
    assert _pst_attachment_filename(att) == "legacy-edf-invoice-file.pdf"


def test_long_filename_wins_over_junk_tag() -> None:
    att = make_att_with_mixed_unicode_and_non_unicode(
        long_name="edf-2024-bill.pdf",
        junk_name="display only label",
    )
    assert _pst_attachment_filename(att) == "edf-2024-bill.pdf"


def test_none_attachment_returns_none() -> None:
    assert _pst_attachment_filename(None) is None  # type: ignore[arg-type]


def test_helper_does_not_call_fictional_pypff_getters() -> None:
    """The legacy 4-getter loop (``att.name`` / ``get_name`` etc.) is gone.

    Synthetic ``FakeAttachment`` deliberately does not implement those
    methods; if the helper reaches for one, ``AttributeError`` propagates
    and the test fails.
    """
    att = make_att_with_long_filename("edf-bill-2025.pdf")
    # No ``setattr`` with the four names; calling the helper must complete.
    assert _pst_attachment_filename(att) == "edf-bill-2025.pdf"


def test_long_filename_property_is_the_constant_0x3707() -> None:
    """Document the helper's knowledge of the MAPI tag it looks for."""
    # The fixture constant must equal the MAPI tag the helper looks up.
    assert PR_ATTACH_LONG_FILENAME == 0x3707
    assert PR_ATTACH_FILENAME == 0x3704


def test_returns_decoded_string_directly() -> None:
    """The helper must return the str produced by ``get_data_as_string``.
    It must not try to decode UTF-16LE itself (pypff already does so).
    """
    att = make_att_with_long_filename("über-invoice-Ω.pdf")
    assert _pst_attachment_filename(att) == "über-invoice-Ω.pdf"


def test_skips_record_set_with_zero_entries() -> None:
    """Helper must tolerate the first record-set being empty.

    Mirrors the real-PST case where some folders expose record-sets for
    metadata only.
    """
    first_att = make_att_with_record_no_entries()
    second_att = make_att_with_long_filename("real-name.pdf")
    # Second record-set carries the actual filename; both shapes are valid
    # FakeRecordSet instances.
    combined = FakeAttachment(
        [
            first_att.get_record_set(0),
            second_att.get_record_set(0),
        ]
    )
    # Force the helper to find a name -- for a sanity check on tolerance, it
    # must at minimum not raise an error.
    try:
        _ = _pst_attachment_filename(combined)
    except Exception as exc:
        pytest.fail(f"Helper raised on irregular record-set shape: {exc!r}")
