"""Smoke/pin test for the DOCX payment analysis section.

Pins the rendered strings of ``create_payment_analysis`` so the migration
to the shared ``compute_payment_analysis`` (models/report_models.py)
cannot silently change what appears in the Word document.
"""

from __future__ import annotations

import pandas as pd
import pytest
from docx import Document
from docx.document import Document as DocumentType

from edf_bill_fetcher.io.reporters.docx_report import (
    _get_or_create_styles,
    create_payment_analysis,
)

RECORD_KEYS = ["Date", "Entry Type", "Period Charge (£)", "Amount (£)", "Details"]


@pytest.fixture
def doc_styles() -> tuple[DocumentType, object]:
    d = Document()
    return d, _get_or_create_styles(d)


def _doc_text(doc: DocumentType) -> str:
    return "\n".join(p.text for p in doc.paragraphs)


def test_two_payments_render_summary_and_interval(doc_styles) -> None:
    doc, styles = doc_styles
    df = pd.DataFrame(
        [
            {
                "Date": "01/01/2023",
                "Entry Type": "Payment",
                "Period Charge (£)": 100,
                "Amount (£)": 100,
                "Details": "First payment",
            },
            {
                "Date": "31/01/2023",
                "Entry Type": "Payment",
                "Period Charge (£)": 50,
                "Amount (£)": 50,
                "Details": "Second payment",
            },
        ]
    )

    create_payment_analysis(doc, styles, df)

    text = _doc_text(doc)
    assert "Number of payments: 2" in text
    assert "Total paid: £150.00" in text
    assert "Average payment: £75.00" in text
    assert "Average days between payments: 30.0" in text


def test_empty_frame_renders_only_page_break(doc_styles) -> None:
    doc, styles = doc_styles
    create_payment_analysis(doc, styles, pd.DataFrame(columns=RECORD_KEYS))

    text = _doc_text(doc)
    assert "Number of payments" not in text
    assert "Total paid" not in text
    assert "Average payment" not in text
