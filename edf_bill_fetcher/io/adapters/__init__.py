"""File-reading adapters for the evidence pipeline.

Owns the framework-library wrappers used by the EvidenceEngine:

* ``pdf``  — pdfplumber-backed PDF text extraction, page slicing, and
  the admit-phrase detector used by the Back-billing analysis sheet.
* ``pst``  — pypff-backed PST / OST attachment-filename walker and
  sender-email extractor used by ``EvidenceEngine.crawl_pst``.
* ``html`` — beautifulsoup4-backed HTM account-history parser used
  by ``EvidenceEngine.process_htm_file``.

During the modularization refactor window (Tasks 5 / 7) the
extracted helpers are also re-exported from ``edf_collector.py`` so
existing ``from edf_collector import slice_pdf_pages`` call sites
keep working.  Task 7 strips those compat re-exports.
"""

from __future__ import annotations

from edf_bill_fetcher.io.adapters import html, pdf, pst
from edf_bill_fetcher.io.adapters.html import (
    htm_excerpt,
    parse_htm_account_history,
)
from edf_bill_fetcher.io.adapters.pdf import (
    ADMIT_RE,
    INV_BOUNDARY_RE,
    LEGAL_CONTEXT,
    PAGE1_BOUNDARY_RE,
    extract_admit_phrase,
    legal_context,
    slice_pdf_pages,
)
from edf_bill_fetcher.io.adapters.pst import (
    EMAIL_ADDR_RE,
    FROM_HEADER_RE,
    PST_PR_ATTACH_FILENAME,
    PST_PR_ATTACH_LONG_FILENAME,
    extract_sender_email,
    matches_domain_filter,
    pst_attachment_filename,
)

__all__ = [
    "ADMIT_RE",
    "EMAIL_ADDR_RE",
    "FROM_HEADER_RE",
    "INV_BOUNDARY_RE",
    "LEGAL_CONTEXT",
    "PAGE1_BOUNDARY_RE",
    "PST_PR_ATTACH_FILENAME",
    "PST_PR_ATTACH_LONG_FILENAME",
    "extract_admit_phrase",
    "extract_sender_email",
    "html",
    "htm_excerpt",
    "legal_context",
    "matches_domain_filter",
    "parse_htm_account_history",
    "pdf",
    "pst",
    "pst_attachment_filename",
    "slice_pdf_pages",
]
