"""Shared PST/OST file-reading primitives — attachment-filename walking and sender-email extraction.

Single source of truth for ``pst_attachment_filename`` and
``extract_sender_email`` (plus the ``pypff`` MAPI constants and regexes
they depend on), shared by ``io.adapters.pst``, ``collectors.engine`` and
``processors.extraction``.  Extracted during the modularization refactor so
the three sites cannot drift apart.

The helpers are the *file-reading primitives* the ``EvidenceEngine.crawl_pst``
loop depends on.  ``pypff`` / ``libpff-python`` is the only mandatory dep; the
helpers tolerate a missing or version-mismatched library by returning safe
defaults (``None`` / empty string) rather than propagating ``AttributeError``.
"""

from __future__ import annotations

import re

__all__ = [
    "EMAIL_ADDR_RE",
    "FROM_HEADER_RE",
    "PST_PR_ATTACH_LONG_FILENAME",
    "extract_sender_email",
    "pst_attachment_filename",
]


# MAPI tag constants from [MS-OXPROPS].
PST_PR_ATTACH_LONG_FILENAME = 0x3707

# `extract_sender_email` pulls an email out of either the transport
# headers (multi-line From:) or the sender name.  Compile both.
FROM_HEADER_RE = re.compile(
    r"^From:\s*.*?([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})",
    re.MULTILINE | re.IGNORECASE,
)
EMAIL_ADDR_RE = re.compile(r"([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})")


def pst_attachment_filename(att: object) -> str | None:
    """Walk the MAPI record-sets of a ``pypff.attachment`` and return its filename.

    Returns the filename string (``str``) when the ``PR_ATTACH_LONG_FILENAME``
    entry is found, else ``None``.  The caller is expected to fall back to
    ``Attachment_N.pdf`` (or whatever synthetic name) when this returns
    ``None``.

    Designed to tolerate malformed record-sets: a missing record entry,
    broken record collection, or zero-record attachment produce a clean
    ``None`` rather than propagating ``AttributeError`` / ``IndexError``
    out to the caller.
    """
    if att is None:
        return None
    getter_count = getattr(att, "get_number_of_record_sets", None)
    if getter_count is None:
        return None
    try:
        n = int(getter_count())
    except Exception:
        return None
    for i in range(n):
        try:
            rs = att.get_record_set(i)  # type: ignore[attr-defined]
        except Exception:
            continue
        entries_getter = getattr(rs, "get_number_of_entries", None)
        if entries_getter is None:
            continue
        try:
            m = int(entries_getter())
        except Exception:
            continue
        for j in range(m):
            try:
                entry = rs.get_entry(j)  # type: ignore[attr-defined]
            except Exception:
                continue
            try:
                entry_type = int(entry.entry_type)  # type: ignore[attr-defined]
            except Exception:
                continue
            if entry_type != PST_PR_ATTACH_LONG_FILENAME:
                continue
            try:
                val = entry.get_data_as_string()  # type: ignore[attr-defined]
            except Exception:
                continue
            if isinstance(val, str) and val:
                return val
            try:
                raw_data = entry.get_data()  # type: ignore[attr-defined]
            except Exception:
                continue
            if isinstance(raw_data, bytes | bytearray) and raw_data:
                try:
                    decoded = bytes(raw_data).decode("utf-16-le", errors="replace")
                except Exception:
                    continue
                if decoded.strip("\x00"):
                    return decoded.strip("\x00")
    return None


def extract_sender_email(msg: object) -> str:
    """Extract sender email address from a pypff message, trying multiple methods."""
    sender: str | None = None
    try:
        headers = msg.get_transport_headers()  # type: ignore[attr-defined]
        if headers:
            headers_str = (
                headers if isinstance(headers, str) else headers.decode("utf-8", errors="replace")
            )
            m = FROM_HEADER_RE.search(headers_str)
            if m:
                sender = m.group(1).lower()
    except Exception:
        pass
    if not sender:
        try:
            name = msg.get_sender_name() or ""  # type: ignore[attr-defined]
            m = EMAIL_ADDR_RE.search(name)
            if m:
                sender = m.group(1).lower()
        except Exception:
            pass
    return sender or ""
