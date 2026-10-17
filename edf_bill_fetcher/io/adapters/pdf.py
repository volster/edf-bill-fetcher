"""PDF reading adapters — text extraction, page slicing, and the admit-phrase detector used by the Back-billing analysis sheet.

The functions in this module are the *file-reading primitives* of
the evidence engine. They are framework-agnostic from the caller's
perspective (pdfplumber is the only mandatory dep) and exposed via
``edf_bill_fetcher.io.adapters.pdf`` (the top-level ``edf_collector.py``
compat module was removed in the modularization).
"""

from __future__ import annotations

import logging
import re

logger = logging.getLogger(__name__)

__all__ = [
    "ADMIT_RE",
    "INV_BOUNDARY_RE",
    "PAGE1_BOUNDARY_RE",
    "extract_admit_phrase",
    "legal_context",
    "slice_pdf_pages",
]


INV_BOUNDARY_RE = re.compile(r"(?:Invoice|Credit note) number:\s*([A-Z0-9-]+)", re.I)
PAGE1_BOUNDARY_RE = re.compile(
    r"(?ix)\b"
    r"(?:page\s+)?"
    r"(?:1|one)"
    r"\s*(?:of|/)\s*"
    r"(?:\d+|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve)"
    r"\b(?!\s*/\s*\d)"
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
    "Back-billing review (Ofgem protections):\n"
    "Domestic back-billing protections may limit recovery of charges for energy supplied\n"
    "more than 12 months before the charge was first raised. Applicability requires\n"
    "review of the original bills, supplier contact, customer conduct and relevant exceptions.\n"
    "This workbook identifies consumption outside a nominal 365-day window and calculates\n"
    "indicative amounts for review. Cancellation wording records an admission that charges\n"
    "were changed; the reason for that change and eligibility for redress require assessment."
)


def slice_pdf_pages(page_texts: list[str]) -> list[list[str]]:
    """Slice a PDF's per-page text into one chunk per invoice.

    A page starts a slice when its invoice / credit-note number changes or
    it contains a ``Page 1 of N`` marker (variants ``1 of 4``, ``one of four``,
    ``1/4`` all match). The final page of an invoice is inclusive --
    it stays with its slice. Single-invoice PDFs return
    ``[list(page_texts)]`` (no behaviour change vs. the legacy
    whole-document concat).
    """
    boundaries: list[int] = []
    document_number: str | None = None
    for i, text in enumerate(page_texts):
        if not text:
            continue
        number_match = INV_BOUNDARY_RE.search(text)
        number = number_match.group(1).upper() if number_match else None
        page_one = bool(PAGE1_BOUNDARY_RE.search(text))
        # Invoice headers may repeat on every page. Only a changed document
        # identity or an explicit first-page marker starts another slice.
        if page_one or (
            number is not None
            and number != document_number
            and (document_number is not None or not boundaries)
        ):
            boundaries.append(i)
        if number is not None or page_one:
            # A first-page cover may omit the number until its continuation.
            # Learning that identity does not split the existing document.
            document_number = number

    if len(boundaries) <= 1:
        return [list(page_texts)]

    slices: list[list[str]] = []
    for j, start in enumerate(boundaries):
        end = boundaries[j + 1] if j + 1 < len(boundaries) else len(page_texts)
        slices.append(page_texts[start:end])

    first = boundaries[0]
    if first > 0:
        preamble = page_texts[:first]
        preamble_size = sum(len(p.strip()) for p in preamble)
        if preamble_size > 200:
            logger.warning(
                "slice_pdf_pages: %d page(s) before the first invoice boundary "
                "folded into the first slice (%d chars of content)",
                first,
                preamble_size,
            )
        slices[0] = preamble + slices[0]
    return slices


def join_sap_pdf_pages(page_texts: list[str]) -> tuple[str, list[int]]:
    """Remove only verified repeated partial CSV tails at PDF page boundaries.

    Some exports print the tail of the last row again at the top of the
    next page. Never remove a complete quoted row, even if it repeats.
    """
    parts: list[str] = []
    repaired: list[int] = []
    for page_number, text in enumerate(page_texts, 1):
        first, separator, rest = text.partition("\n")
        if (
            parts
            and separator
            and first
            and not first.lstrip().startswith('"')
            and first.count(",") >= 2
            and parts[-1].rstrip().endswith(first)
            and rest.lstrip().startswith('"')
        ):
            text = rest
            repaired.append(page_number)
        parts.append(text)
    return "\n".join(parts), repaired


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
    """Return the static legal-context blurb placed on the Back-billing sheet.

    Kept as a function (not a bare module constant) so the text
    can be regenerated / internationalised later without changing
    call-sites.
    """
    return _LEGAL_CONTEXT
