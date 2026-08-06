"""Synthetic ``pypff`` attachment subtree for tests.

The four ``pypff.attachment`` getters currently invoked in the live code path
(``att.name``, ``att.get_name``, ``att.get_long_filename``, ``att.get_short_filename``)
do not exist on the real ``pypff`` C extension, so the production code always
raises ``AttributeError`` and silently falls back to ``Attachment_N.pdf``.

The real fix (verified against ``scratch/input/edf.pst``) walks MAPI record-set
entries looking for ``PR_ATTACH_LONG_FILENAME`` (0x3707).  Each ``record_entry``
exposes:
    * ``.entry_type``  -- the MAPI tag (e.g. 0x3707 for the long filename).
    * ``.value_type``  -- the MAPI property data type (e.g. 0x001F for PT_UNICODE).
    * ``.get_data_as_string()`` -- returning an already-decoded Python str.

This fixture provides a synthetic ``pypff`` attach object with the same
``record_sets`` shape, so ``_pst_attachment_filename`` can be exercised without
a real PST file.
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


# ---------------------------------------------------------------------------
# Shapes added for ``edf_bill_fetcher.io.adapters.pst`` branch coverage.
# These exercise the defensive ``try/except`` ladders and the raw-bytes
# fallback in ``pst_attachment_filename``, plus the ``extract_sender_email``
# transport-headers / sender-name paths.
# ---------------------------------------------------------------------------


class FakeRecordEntryRawBytes:
    """``pypff.record_entry`` whose ``get_data_as_string`` returns empty.

    Mirrors the PT_UNICODE raw-bytes edge case: ``get_data_as_string()``
    yields an empty ``str`` (so the str branch at line 91 is falsy) and
    ``get_data()`` returns the raw UTF-16LE-encoded ``bytes`` the helper
    must decode itself.
    """

    def __init__(self, entry_type: int, raw_data: bytes) -> None:
        self.entry_type = entry_type
        self._raw_data = raw_data

    def get_data_as_string(self) -> str:
        return ""

    def get_data(self) -> bytes:
        return self._raw_data


class FakeRecordEntryEmptyBytes:
    """``get_data_as_string`` empty AND ``get_data`` returns empty bytes.

    Exercises the ``decoded.strip("\\x00")`` falsy branch (line 102-103):
    the helper decodes successfully but the result is all NULs, so it
    must fall through to ``return None`` rather than returning an empty
    filename.
    """

    def __init__(self, entry_type: int) -> None:
        self.entry_type = entry_type

    def get_data_as_string(self) -> str:
        return ""

    def get_data(self) -> bytes:
        return b"\x00\x00"


class FakeRecordEntryGetDataRaises:
    """``get_data_as_string`` and ``get_data`` both raise.

    Exercises the two ``except Exception: continue`` branches at lines
    89-90 and 95-96 on the same entry.
    """

    def __init__(self, entry_type: int) -> None:
        self.entry_type = entry_type

    def get_data_as_string(self) -> str:
        raise RuntimeError("pypff version mismatch: get_data_as_string crashed")

    def get_data(self) -> bytes:
        raise RuntimeError("pypff version mismatch: get_data crashed")


class FakeRecordEntryEmptyStringGetDataRaises:
    """``get_data_as_string`` returns empty str; ``get_data`` raises.

    Exercises the ``except`` branch at lines 95-96 specifically: the str
    branch at line 91 is falsy (empty string), so the helper falls through
    to ``get_data()``, which raises and triggers the second ``continue``.
    """

    def __init__(self, entry_type: int) -> None:
        self.entry_type = entry_type

    def get_data_as_string(self) -> str:
        return ""

    def get_data(self) -> bytes:
        raise RuntimeError("pypff version mismatch: get_data crashed")


class FakeRecordEntryEmptyStringNonBytesData:
    """``get_data_as_string`` empty; ``get_data`` returns non-bytes.

    Exercises the ``isinstance(raw_data, bytes | bytearray)`` falsy branch
    at line 97 (branch 97->76): the helper gets a value back from
    ``get_data()`` but it is neither ``bytes`` nor ``bytearray``, so the
    ``if`` is False and the loop iterates.
    """

    def __init__(self, entry_type: int) -> None:
        self.entry_type = entry_type

    def get_data_as_string(self) -> str:
        return ""

    def get_data(self) -> object:
        return 12345


class FakeRecordEntryBadEntryType:
    """``entry_type`` raises on ``int()`` conversion (line 82-83)."""

    def __init__(self) -> None:
        self.entry_type = "not-an-int"  # type: ignore[assignment]

    def get_data_as_string(self) -> str:
        return "unreachable"


class FakeRecordSetNoEntriesGetter:
    """Record-set missing ``get_number_of_entries`` (line 69-70).

    ``getattr`` returns ``None`` and the helper ``continue``s to the next
    record-set.
    """

    def __init__(self) -> None:
        pass


class FakeRecordSetEntriesGetterRaises:
    """``get_number_of_entries`` raises on ``int()`` (line 73-74)."""

    def __init__(self) -> None:
        pass

    def get_number_of_entries(self) -> int:
        raise RuntimeError("corrupt record set: count unreadable")


class FakeRecordSetGetEntryRaises:
    """``get_entry(j)`` raises for every index (line 78-79)."""

    def __init__(self, n_entries: int = 1) -> None:
        self._n = n_entries

    def get_number_of_entries(self) -> int:
        return self._n

    def get_entry(self, index: int) -> FakeRecordEntry:
        raise IndexError("corrupt record set: entry missing")


class FakeAttachmentNoRecordSetGetter:
    """Attachment missing ``get_number_of_record_sets`` (line 57-58).

    ``getattr`` returns ``None`` and the helper returns ``None``.
    """


class FakeAttachmentRecordSetCountRaises:
    """``get_number_of_record_sets()`` raises on ``int()`` (line 61-62)."""

    def get_number_of_record_sets(self) -> int:
        raise RuntimeError("corrupt attachment: record-set count unreadable")


class FakeAttachmentGetRecordSetRaises:
    """``get_record_set(i)`` raises for every index (line 66-67)."""

    def __init__(self, n_record_sets: int = 1) -> None:
        self._n = n_record_sets

    def get_number_of_record_sets(self) -> int:
        return self._n

    def get_record_set(self, index: int) -> FakeRecordSet:
        raise IndexError("corrupt attachment: record set missing")


def make_att_with_only_short_filename(short_name: str) -> FakeAttachment:
    """Attachment carrying only ``PR_ATTACH_FILENAME`` (the short tag).

    The helper looks exclusively for ``PR_ATTACH_LONG_FILENAME`` (0x3707),
    so a short-filename-only attachment yields ``None`` -- the caller is
    then expected to fall back to ``Attachment_N.pdf``.
    """
    entry_short = FakeRecordEntry(PR_ATTACH_FILENAME, PT_UNICODE, short_name)
    return FakeAttachment([FakeRecordSet([entry_short])])


def make_att_with_no_filename_entries_v2(name: str = "display only") -> FakeAttachment:
    """Attachment whose single entry is neither filename tag (line 85-86)."""
    entry = FakeRecordEntry(PR_ATTACH_DISPLAY_NAME, PT_UNICODE, name)
    return FakeAttachment([FakeRecordSet([entry])])


def make_att_with_raw_bytes_long_filename(name: str) -> FakeAttachment:
    """Long-filename entry whose ``get_data_as_string`` is empty.

    Forces the helper down the ``get_data()`` -> UTF-16LE decode branch
    (lines 93-103). ``name`` is encoded UTF-16LE so the decoded result
    equals ``name``.
    """
    entry: FakeRecordEntry = FakeRecordEntryRawBytes(
        PR_ATTACH_LONG_FILENAME, name.encode("utf-16-le")
    )  # type: ignore[assignment]
    return FakeAttachment([FakeRecordSet([entry])])


def make_att_with_empty_raw_bytes_long_filename() -> FakeAttachment:
    """Long-filename entry whose raw bytes decode to all-NULs.

    Exercises the ``decoded.strip("\\x00")`` falsy branch (line 102-103):
    the helper decodes but discards the empty result and falls through to
    ``return None``.
    """
    entry: FakeRecordEntry = FakeRecordEntryEmptyBytes(PR_ATTACH_LONG_FILENAME)  # type: ignore[assignment]
    return FakeAttachment([FakeRecordSet([entry])])


def make_att_with_get_data_raises() -> FakeAttachment:
    """Long-filename entry whose both data getters raise.

    Exercises the ``except`` branches at lines 89-90 and 95-96.
    """
    entry: FakeRecordEntry = FakeRecordEntryGetDataRaises(PR_ATTACH_LONG_FILENAME)  # type: ignore[assignment]
    return FakeAttachment([FakeRecordSet([entry])])


def make_att_with_empty_string_get_data_raises() -> FakeAttachment:
    """Long-filename entry: empty str from ``get_data_as_string``, ``get_data`` raises.

    Exercises the ``except`` branch at lines 95-96 (str branch falsy, then
    ``get_data()`` raises).
    """
    entry: FakeRecordEntry = FakeRecordEntryEmptyStringGetDataRaises(PR_ATTACH_LONG_FILENAME)  # type: ignore[assignment]
    return FakeAttachment([FakeRecordSet([entry])])


def make_att_with_empty_string_non_bytes_data() -> FakeAttachment:
    """Long-filename entry: empty str, ``get_data`` returns non-bytes.

    Exercises the ``isinstance(raw_data, bytes | bytearray)`` falsy branch
    at line 97 (branch 97->76).
    """
    entry: FakeRecordEntry = FakeRecordEntryEmptyStringNonBytesData(PR_ATTACH_LONG_FILENAME)  # type: ignore[assignment]
    return FakeAttachment([FakeRecordSet([entry])])


def make_att_with_bad_entry_type() -> FakeAttachment:
    """Entry whose ``entry_type`` cannot be coerced to ``int`` (line 82-83)."""
    entry: FakeRecordEntry = FakeRecordEntryBadEntryType()  # type: ignore[assignment]
    return FakeAttachment([FakeRecordSet([entry])])


def make_att_with_record_set_no_entries_getter() -> FakeAttachment:
    """First record-set lacks ``get_number_of_entries`` (line 69-70).

    A second, valid record-set carries the long filename so the helper
    still returns it -- proving the ``continue`` skips the broken set.
    """
    broken_rs: FakeRecordSet = FakeRecordSetNoEntriesGetter()  # type: ignore[assignment]
    good_entry = FakeRecordEntry(PR_ATTACH_LONG_FILENAME, PT_UNICODE, "good.pdf")
    good_rs = FakeRecordSet([good_entry])
    return FakeAttachment([broken_rs, good_rs])  # type: ignore[list-item]


def make_att_with_record_set_entries_getter_raises() -> FakeAttachment:
    """First record-set's ``get_number_of_entries`` raises (line 73-74)."""
    broken_rs: FakeRecordSet = FakeRecordSetEntriesGetterRaises()  # type: ignore[assignment]
    good_entry = FakeRecordEntry(PR_ATTACH_LONG_FILENAME, PT_UNICODE, "good.pdf")
    good_rs = FakeRecordSet([good_entry])
    return FakeAttachment([broken_rs, good_rs])  # type: ignore[list-item]


def make_att_with_record_set_get_entry_raises() -> FakeAttachment:
    """Record-set whose ``get_entry`` raises for every index (line 78-79)."""
    broken_rs: FakeRecordSet = FakeRecordSetGetEntryRaises(n_entries=2)  # type: ignore[assignment]
    return FakeAttachment([broken_rs])  # type: ignore[list-item]


def make_att_with_get_record_set_raises() -> FakeAttachment:
    """Attachment whose ``get_record_set`` raises for every index (line 66-67)."""
    return FakeAttachmentGetRecordSetRaises(n_record_sets=2)  # type: ignore[return-value]


def make_multiple_attachments_with_long_filenames(names: list[str]) -> list[FakeAttachment]:
    """Build N attachments, each carrying one long-filename entry.

    For exercising the helper across multiple attachments in a single
    test (mirrors the ``EvidenceEngine.crawl_pst`` loop shape).
    """
    attachments: list[FakeAttachment] = []
    for name in names:
        entry = FakeRecordEntry(PR_ATTACH_LONG_FILENAME, PT_UNICODE, name)
        attachments.append(FakeAttachment([FakeRecordSet([entry])]))
    return attachments


class FakeMessage:
    """Mimics ``pypff.message`` enough to exercise ``extract_sender_email``.

    ``get_transport_headers`` and ``get_sender_name`` are the two getters
    the helper probes; either may return ``str``, ``bytes``, ``None``, or
    raise, depending on the PST shape.
    """

    def __init__(
        self,
        transport_headers: str | bytes | None = None,
        sender_name: str | None = None,
        transport_raises: bool = False,
        sender_raises: bool = False,
    ) -> None:
        self._transport_headers = transport_headers
        self._sender_name = sender_name
        self._transport_raises = transport_raises
        self._sender_raises = sender_raises

    def get_transport_headers(self) -> str | bytes | None:
        if self._transport_raises:
            raise RuntimeError("pypff: transport headers unreadable")
        return self._transport_headers

    def get_sender_name(self) -> str | None:
        if self._sender_raises:
            raise RuntimeError("pypff: sender name unreadable")
        return self._sender_name
