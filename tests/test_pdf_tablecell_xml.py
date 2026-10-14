"""PDF cells preserve literal evidence text and safely wrap inside the page.

ReportLab Table strings are plain text; printing escaped entities as those
strings changes the evidence. Paragraph cells require escaping once and
consume that escaping when rendered, retaining the original visible text.
"""

from __future__ import annotations

import io

import pandas as pd
import pdfplumber
import pytest
from pdfplumber.pdf import PDF
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate

from edf_bill_fetcher.io.reporters.pdf_report import (
    create_appendix_full_evidence,
    create_evidence_index,
)


def _make_df() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Date": ["01/03/2024", "15/04/2024"],
            "Period From": ["01/02/2024", "01/03/2024"],
            "Period To": ["01/03/2024", "01/04/2024"],
            "Invoice #": ["INV-<safe>", "INV&here"],
            "Amount (£)": [100.0, 80.0],
            "Period Charge (£)": [80.0, 60.0],
            "Units (kWh)": ["100", "80"],
            "Reading": ["Estimated<bad>", "Actual&worse"],
            "Entry Type": ["New Bill<inject>", "New Bill"],
            "Source": ["<HTM>", "PST"],
            "Tariff": ["Standard", "Standard Tariff"],
        }
    )


def _render(elements: list) -> io.BytesIO:
    buffer = io.BytesIO()
    SimpleDocTemplate(buffer, pagesize=A4).build(elements)
    buffer.seek(0)
    return buffer


def _table_text(document: PDF) -> str:
    """Read each cell separately so wrapped rows do not interleave columns."""
    return "".join(
        "".join((cell or "").split())
        for page in document.pages
        for table in page.extract_tables()
        for row in table
        for cell in row
    )


@pytest.mark.parametrize("section", ["index", "appendix"])
def test_source_markup_is_visible_literal_text_and_cells_stay_on_page(section: str) -> None:
    data = _make_df()
    elements = (
        create_evidence_index(data, None)
        if section == "index"
        else create_appendix_full_evidence(data)
    )
    with pdfplumber.open(_render(elements)) as document:
        # Wrapping can put a reference on several lines; compare characters
        # without whitespace, while checking the actual PDF's page bounds.
        text = _table_text(document)
        for original in ("<HTM>", "INV-<safe>", "INV&here", "Estimated<bad>", "Actual&worse"):
            assert original in text
        assert "&lt;" not in text
        assert "&amp;" not in text
        for page in document.pages:
            assert all(0 <= char["x0"] <= char["x1"] <= page.width for char in page.chars)


def test_source_link_markup_cannot_create_a_pdf_annotation() -> None:
    data = _make_df().iloc[:1].copy()
    data["Invoice #"] = '<link href="https://example.org">original</link>'
    data["Reading"] = "Literal &lt;token&gt;"
    with pdfplumber.open(_render(create_evidence_index(data, None))) as document:
        text = _table_text(document)
        assert '<linkhref="https://example.org">original</link>' in text
        assert "Literal&lt;token&gt;" in text
        assert all(not page.annots for page in document.pages)
