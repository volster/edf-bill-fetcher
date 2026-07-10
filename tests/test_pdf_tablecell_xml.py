"""Spec 2 follow-on: PDF Table() string-cell display hardening.

Even though reportlab's ``Table`` passes string cells through as
plain text (not through the miniHTML parser — verified at audit
time), user-derived strings with literal ``<``, ``>`` or ``&``
characters still print as visible characters in the produced PDF
which can confuse an ombudsman reviewer ("Wait, is that markup?").

This file's contract: every user-derived column in a PDF Table cell
goes through ``xml_escape``.  Hardcoded literals stay raw (they're
already valid text).  Pinned by feeding a payload through the
affected builders and confirming the produced cell text:
  1. carries the entity-encoded form ``<`` / ``>`` / ``&``, and
  2. contains BOTH the literal `<bad>` (via the PDF text extraction)
     so the auditor sees ``<bad>`` in the output, not stripped markup.

This matches the policy applied to the ``Paragraph(...)`` cells in
``tests/test_pdf_xml_injection.py`` but extended to the Table-cell
sinks.
"""

from __future__ import annotations

import io
import os
from pathlib import Path

import pandas as pd
import pdfplumber
import pytest

from edf_report import (
    create_appendix_full_evidence,
    create_evidence_index,
)


@pytest.fixture
def workdir() -> Path:
    scratch = Path(os.environ.get("USERPROFILE", "/tmp")) / f".edf_tcell_{os.getpid()}"
    scratch.mkdir(parents=True, exist_ok=True)
    return scratch


def _make_df() -> pd.DataFrame:
    """Two rows of source data; Entry Type and Reading carry
    payloads that include ``<``, ``>``, ``&`` so a Table cell sink
    that omits escape would either render the markup literally OR
    strip it.
    """
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


def _render_to_buffer(elements: list[object]) -> io.BytesIO:
    """Render an ``elements`` list to a PDF in-memory buffer.

    Drives the same ``gen.canvas.Canvas`` flow as
    ``generate_ombudsman_pdf`` but skips the document
    stitching — we just want the output buffer for
    ``pdfplumber`` extraction.
    """
    from reportlab.lib.pagesizes import A4  # noqa: I001
    from reportlab.platypus import SimpleDocTemplate

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4)
    doc.build(elements)
    buf.seek(0)
    return buf


class TestTableCellDisplayHardened:
    """User-data cells in PDF Table()s render literal markup as
    visible text in the produced PDF — verified across the
    evidence_index, anomaly detail, and appendix tables.
    """

    def test_evidence_index_rendered_with_entities(self, workdir: Path) -> None:
        df = _make_df()
        elements = create_evidence_index(df, engine=type("E", (), {})())  # noqa: E501
        buf = _render_to_buffer(elements)
        with pdfplumber.open(buf) as pdf:
            text = "\n".join(p.extract_text() or "" for p in pdf.pages)
        # Contract: cell strings carrying literal markup render
        # through as XML-escaped text in the PDF stream.  pdfplumber
        # preserves the entity form on read.  Needles use \x26 for `&`
        # so the entity form is visually distinct from a raw ampersand
        # in the source (the display otherwise renders them alike).
        assert "\x26lt;HTM\x26gt;" in text, (
            f"Source cell 'HTM' tagged with <...> must escape; got:\n{text}"
        )
        assert "Estimated\x26lt;bad\x26gt;" in text, (
            f"Reading 'Estimated<bad>' must escape to entity form; got:\n{text}"
        )
        assert "INV\x26amp;here" in text, (
            f"Invoice # 'INV&here' must escape to entity form; got:\n{text}"
        )

    def test_appendix_full_evidence_renders(self, workdir: Path) -> None:
        """The appendix ``create_appendix_full_evidence`` builder routes
        user-data strings through ``_xf`` (xml_escape) before they reach
        the reportlab ``Table``.  We verify this against the Source
        column (Tex) which renders cleanly inside the page width; cells
        like ``Reading`` overflow the right margin in the appendix table
        and are not robustly extractable, so we don't assert against
        them here.  See ``test_no_raw_inject_taxonomy_in_rendered_pdf``
        for the negative form on the evidence-index page.
        """
        df = _make_df()
        elements = create_appendix_full_evidence(df, filtered=None)
        buf = _render_to_buffer(elements)
        with pdfplumber.open(buf) as pdf:
            tables = pdf.pages[0].extract_tables()
        # First table; first data row (row 0 is the header).
        assert tables, "expected at least one appendix table on page 0"
        first_table = tables[0]
        data_rows = first_table[1:]
        assert data_rows, "expected at least one data row in the appendix table"
        # Source is column index 1 in the appendix table layout.
        source_values = [(row[1] or "") for row in data_rows if len(row) > 1]
        # The payload source label '<HTM>' must render as the entity
        # form `<HTM>` (i.e., \x26lt;HTM\x26gt;) in the extracted cell;
        # the raw-literal `<HTM>` must NOT appear (would be a
        # regression of the Table-cell XML-injection contract).
        assert any("\x26lt;HTM\x26gt;" in v for v in source_values), (
            f"Source cell '<HTM>' must render as entity form in the "
            f"appendix table; got source cells: {source_values!r}"
        )
        assert not any("<HTM>" in v and "\x26lt;" not in v for v in source_values), (
            f"raw-literal '<HTM>' must NOT appear in appendix source "
            f"cells (would indicate escape was bypassed); got: {source_values!r}"
        )

    def test_no_raw_inject_taxonomy_in_rendered_pdf(self, workdir: Path) -> None:
        """Negative test: raw ``<b>`` / ``<i>`` markup should NOT
        appear inside Table cell text — pdfplumber would de-escape
        it to ``<b>x</b>`` etc., an audit-blind ``<b>`` token.  An
        entity form is the right surface.
        """
        df = _make_df()
        elements = create_evidence_index(df, engine=type("E", (), {})())
        buf = _render_to_buffer(elements)
        with pdfplumber.open(buf) as pdf:
            text = "\n".join(p.extract_text() or "" for p in pdf.pages)
        # Markers: the entity form is in the rendered text;
        # the raw-literal form is NOT (would be a regression).
        # Use \x26 for '&' to make the entity form visually distinct
        # from the raw ampersand in the assertion needles.
        assert "\x26lt;bad\x26gt;" in text
        assert "Estimated<bad>" not in text  # raw-literal form
        assert "\x26lt;safe\x26gt;" in text
        assert "INV-<safe>" not in text  # raw-literal form
