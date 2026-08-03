"""PDF reading adapters — text extraction, page slicing, and the.

admit-phrase detector used by the Back-billing analysis sheet.

The functions in this module are the *file-reading primitives* of
the evidence engine. They are framework-agnostic from the caller's
perspective (pdfplumber is the only mandatory dep) and exposed via
``edf_bill_fetcher.io.adapters.pdf`` while the modularization
refactor window is open.

Compat re-exports live in ``edf_collector.py`` so that callers using
``from edf_collector import slice_pdf_pages`` continue to work.
"""

from __future__ import annotations

import re

__all__ = [
    "ADMIT_RE",
    "INV_BOUNDARY_RE",
    "LEGAL_CONTEXT",
    "PAGE1_BOUNDARY_RE",
    "extract_admit_phrase",
    "legal_context",
    "slice_pdf_pages",
]


INV_BOUNDARY_RE = re.compile(r"Invoice number:\s*[A-Z0-9-]+", re.I)
PAGE1_BOUNDARY_RE = re.compile(
    r"(?ix)\b"
    r"(?:page\s+)?"
    r"(?:1|one)"
    r"\s*(?:of|/)\s+"
    r"(?:\d+|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve)"
    r"\b"
)
ADMIT_RE = re.compile(
    r"(?ix)\b"
    r"(?:we['\u2019]?ve|we\ have|we\ are|we['\u2019]?re)?\s*"
    r"(?:recently\s+|previously\s+)?"
    r"(?:cancel(?:l?ed|ing)|cancell?ed|cancel(?:l?ing)"
    r"|revers(?:ed|ing)"
    r"|credit(?:ed|ing))"
    r".{0,40}?"
    r"(?:charges?|some\ charges|charges\ for\ you|your\ account)"
    r"\b"
)

_LEGAL_CONTEXT: str = (
    "Back-billing protections (Ofgem / Electricity Act 1989 s.84B):\n"
    "Suppliers may not charge a domestic customer for energy supplied\n"
    "more than 12 months before the date of the bill that first raised\n"
    "the charge, unless one of the statutory exceptions applies\n"
    "(customer has been obstructive, has unreasonably refused access,\n"
    "or has not cooperated with the supplier's reasonable requests). A\n"
    "supervisor's admission of an earlier billing error -- typically\n"
    "worded as 'we've recently cancelled some charges for you' on the\n"
    "cover page of a corrective bill -- is direct evidence that the\n"
    "cancellation is a back-billing remedy rather than a goodwill\n"
    "adjustment, and so preserves the 12-month back-billing bar for any\n"
    "superseded invoices on the same period. This workbook flags any\n"
    "invoice admitting such a cancellation as evidence of back-billing."
)


def slice_pdf_pages(page_texts: list[str]) -> list[list[str]]:
    """Slice a PDF's per-page text into one chunk per invoice.

    A page is a slice-start if it contains ``Invoice number:`` OR a
    ``Page 1 of N`` marker (variants ``1 of 4``, ``one of four``,
    ``1/4`` all match). The final page of an invoice is inclusive --
    it stays with its slice. Single-invoice PDFs return
    ``[list(page_texts)]`` (no behaviour change vs. the legacy
    whole-document concat).
    """
    boundaries: list[int] = []
    for i, text in enumerate(page_texts):
        if not text:
            continue
        if INV_BOUNDARY_RE.search(text) or PAGE1_BOUNDARY_RE.search(text):
            boundaries.append(i)

    if len(boundaries) <= 1:
        return [list(page_texts)]

    slices: list[list[str]] = []
    for j, start in enumerate(boundaries):
        end = boundaries[j + 1] if j + 1 < len(boundaries) else len(page_texts)
        slices.append(page_texts[start:end])
    return slices


def extract_admit_phrase(text: str) -> str | None:
    """Return the first admit-phrase match in *text*, or ``None``.

    An admit phrase is EDF's cover-page wording acknowledging that they
    have cancelled / reversed / credited charges (the "we've recently
    cancelled some charges for you" family). The returned string is the
    matched substring (trimmed) -- callers can store it verbatim as
    evidence.
    """
    if not text:
        return None
    m = ADMIT_RE.search(text)
    if m is None:
        return None
    return m.group(0).strip()


def legal_context() -> str:
    """Return the static legal-context blurb placed on the Back-billing.

    sheet. Kept as a function (not a bare module constant) so the text
    can be regenerated / internationalised later without changing
    call-sites.
    """
    return _LEGAL_CONTEXT
LEGAL_CONTEXT = _LEGAL_CONTEXT
