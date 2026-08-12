"""PST / OST reading adapters — attachment-filename walker and sender-email extraction used by the PST / OST archive crawler.

These helpers are the *file-reading primitives* the
``EvidenceEngine.crawl_pst`` loop depends on.  ``pypff`` /
``libpff-python`` is the only mandatory dep; the helpers tolerate a
missing or version-mismatched library by returning safe defaults
(``None`` / empty string) rather than propagating ``AttributeError``.
"""

from __future__ import annotations

from edf_bill_fetcher.helpers.domain_filter import matches_domain_filter
from edf_bill_fetcher.helpers.pst_resources import (
    EMAIL_ADDR_RE,
    FROM_HEADER_RE,
    PST_PR_ATTACH_LONG_FILENAME,
    extract_sender_email,
    pst_attachment_filename,
)

__all__ = [
    "EMAIL_ADDR_RE",
    "FROM_HEADER_RE",
    "PST_PR_ATTACH_FILENAME",
    "PST_PR_ATTACH_LONG_FILENAME",
    "extract_sender_email",
    "matches_domain_filter",
    "pst_attachment_filename",
]


# MAPI tag constants from [MS-OXPROPS].
PST_PR_ATTACH_FILENAME = 0x3704
