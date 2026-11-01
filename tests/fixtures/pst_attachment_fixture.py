"""Synthetic pypff attachment records for filename extraction tests.

MAPI record-set entries expose attachment names through
``PR_ATTACH_LONG_FILENAME`` (0x3707) and related properties. Entries provide
``entry_type``, ``value_type`` and ``get_data_as_string()``. The fixture models
that interface without requiring a private PST file.
"""

from __future__ import annotations

# Real MAPI tag constants from [MS-OXPROPS] / mapitags.h.  Keep them as named
# constants in the fixture so test failures name the missing-tag case
# explicitly.
PR_ATTACH_LONG_FILENAME = 0x3707
PR_ATTACH_FILENAME = 0x3704
PR_ATTACH_DISPLAY_NAME = 0x3001

# Real MAPI data type constants.
PT_UNICODE = 0x001F
PT_STRING8 = 0x001E


class FakeRecordEntry:
    """Mimics ``pypff.record_entry``."""

    def __init__(self, entry_type: int, value_type: int, decoded_string: str) -> None:
        self.entry_type = entry_type
        self.value_type = value_type
        self._decoded_string = decoded_string

    def get_data_as_string(self) -> str:
        return self._decoded_string


class FakeRecordSet:
    """Mimics ``pypff.record_set``."""

    def __init__(self, entries: list[FakeRecordEntry]) -> None:
        self._entries = entries

    def get_number_of_entries(self) -> int:
        return len(self._entries)

    def get_entry(self, index: int) -> FakeRecordEntry:
        return self._entries[index]


class FakeAttachment:
    """Mimics ``pypff.attachment`` enough to exercise _pst_attachment_filename.

    The mocks listed at the top of this module are deliberately NOT defined
    on this fake -- the production code should never call them.
    """

    def __init__(self, record_sets: list[FakeRecordSet]) -> None:
        self._record_sets = record_sets

    def get_number_of_record_sets(self) -> int:
        return len(self._record_sets)

    def get_record_set(self, index: int) -> FakeRecordSet:
        return self._record_sets[index]


# ---------------------------------------------------------------------------
# Factories for common shapes used by ``_pst_attachment_filename`` tests.
# ---------------------------------------------------------------------------


def make_att_with_long_filename(name: str) -> FakeAttachment:
    """Attachment carrying a single PR_ATTACH_LONG_FILENAME entry."""
    entry = FakeRecordEntry(PR_ATTACH_LONG_FILENAME, PT_UNICODE, name)
    rs = FakeRecordSet([entry])
    return FakeAttachment([rs])


def make_att_with_two_record_sets(long_name: str, short_name: str) -> FakeAttachment:
    """Two record-sets: first carries the short filename, second the long."""
    entry_short = FakeRecordEntry(PR_ATTACH_FILENAME, PT_UNICODE, short_name)
    entry_long = FakeRecordEntry(PR_ATTACH_LONG_FILENAME, PT_UNICODE, long_name)
    return FakeAttachment(
        [
            FakeRecordSet([entry_short]),
            FakeRecordSet([entry_long]),
        ]
    )


def make_att_with_no_filename_entries(name: str = "irrelevant") -> FakeAttachment:
    """Attachment has MAPI entries but none are the filename tags."""
    entry = FakeRecordEntry(PR_ATTACH_DISPLAY_NAME, PT_UNICODE, name)
    return FakeAttachment([FakeRecordSet([entry])])


def make_att_with_zero_record_sets() -> FakeAttachment:
    return FakeAttachment([])


def make_att_with_string8_long_filename(name: str) -> FakeAttachment:
    """Same tag but a non-UNICODE data type -- the helper must still extract."""
    entry = FakeRecordEntry(PR_ATTACH_LONG_FILENAME, PT_STRING8, name)
    return FakeAttachment([FakeRecordSet([entry])])


def make_att_with_record_no_entries() -> FakeAttachment:
    """Record-set present but empty -- mimics malformed PST data."""
    return FakeAttachment([FakeRecordSet([])])


def make_att_with_mixed_unicode_and_non_unicode(long_name: str, junk_name: str) -> FakeAttachment:
    """One entry per record-set; long filename lives next to a junk-tag one."""
    entry_junk = FakeRecordEntry(PR_ATTACH_DISPLAY_NAME, PT_UNICODE, junk_name)
    entry_long = FakeRecordEntry(PR_ATTACH_LONG_FILENAME, PT_UNICODE, long_name)
    return FakeAttachment(
        [
            FakeRecordSet([entry_junk]),
            FakeRecordSet([entry_long]),
        ]
    )
