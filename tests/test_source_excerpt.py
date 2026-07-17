"""Tests for the Source Excerpt helper + column rendering on analyser tabs."""

from __future__ import annotations

import pandas as pd
from openpyxl import Workbook

from edf_collector import (
    _format_source_excerpt,
    _source_excerpt_for_invoice,
    write_back_billing_sheet,
    write_rebilling_sheet,
)


def test_format_source_excerpt_includes_regex_trace_and_truncated_text() -> None:
    text = "Invoice number: KI-31105244-0001-3\nLong body text " * 100
    trace = "inv_num via _INV_NUMBER_RE; period_from via _BILLING_PERIOD_RE"
    failed = ["amount"]
    excerpt = _format_source_excerpt(text, trace, failed)
    assert "inv_num via _INV_NUMBER_RE" in excerpt
    assert "FAILED:" in excerpt
    assert "amount" in excerpt
    # Truncation cap present (text body is far longer than the cap).
    assert len(excerpt) < len(text)


def test_format_source_excerpt_with_no_failed_fields() -> None:
    text = "small body"
    trace = "inv_num via _INV_NUMBER_RE"
    excerpt = _format_source_excerpt(text, trace, [])
    assert "FAILED:" not in excerpt
    assert "inv_num via _INV_NUMBER_RE" in excerpt


def test_format_source_excerpt_handles_none_text() -> None:
    excerpt = _format_source_excerpt(None, "", ["inv_num"])  # type: ignore[arg-type]
    assert "FAILED:" in excerpt
    assert "inv_num" in excerpt


def _sample_evidence_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Invoice #": "KI-31105244-0001-3",
                "Source PDF Text": "Invoice number: KI-31105244-0001-3\nLong body.",
                "_regex_trace": "inv_num via _INV_NUMBER_RE",
            },
            {
                "Invoice #": "KCR-31105244-0010-3",
                "Source PDF Text": "Credit note number: KCR-31105244-0010-3\nLong body.",
                "_regex_trace": "inv_num via _CREDIT_NUMBER_RE; period_from via _COVER_BLOCK_PERIOD_RE",
            },
        ]
    )


def test_source_excerpt_for_invoice_looks_up_by_invoice_number() -> None:
    df = _sample_evidence_df()
    excerpt = _source_excerpt_for_invoice(df, "KI-31105244-0001-3")
    assert excerpt is not None
    assert "inv_num via _INV_NUMBER_RE" in excerpt
    assert "Invoice number: KI-31105244-0001-3" in excerpt


def test_source_excerpt_for_invoice_returns_none_when_not_found() -> None:
    df = _sample_evidence_df()
    assert _source_excerpt_for_invoice(df, "UNKNOWN-96") is None


def test_source_excerpt_for_invoice_tolerates_missing_columns() -> None:
    df = pd.DataFrame([{"Invoice #": "X", "Source PDF Text": "abc"}])
    excerpt = _source_excerpt_for_invoice(df, "X")
    assert excerpt is not None
    assert "abc" in excerpt


def test_back_billing_sheet_emits_source_excerpt_column_header() -> None:
    bb = pd.DataFrame(
        [
            {
                "Invoice #": "KI-31105244-0001-3",
                "Bill Date": "01/01/2024",
                "Period From": pd.Timestamp("2022-01-01"),
                "Period To": pd.Timestamp("2024-01-01"),
                "Days Billed": 730,
                "Net Charge (£)": 1347.96,
                "12-Month Limit (days)": 365,
                "Excess Days": 365,
                "Cancel/Rebill Admitted": False,
                "Reason Assessment": "back-billing",
            }
        ]
    )
    ev = _sample_evidence_df()
    wb = Workbook()
    ws = wb.active
    write_back_billing_sheet(ws, bb, account="A-31105244", evidence_df=ev)
    # Header row is row 7. New Source Excerpt column = col 11.
    hdr = ws.cell(row=7, column=11).value
    assert hdr == "Source Excerpt"
    # First body row = row 8. Col 11 should carry the regex-trace excerpt.
    body_excerpt = ws.cell(row=8, column=11).value
    assert isinstance(body_excerpt, str)
    assert "inv_num via _INV_NUMBER_RE" in body_excerpt


def test_back_billing_sheet_handles_missing_evidence_df() -> None:
    bb = pd.DataFrame(
        [
            {
                "Invoice #": "KI-31105244-0001-3",
                "Bill Date": "01/01/2024",
                "Period From": pd.Timestamp("2022-01-01"),
                "Period To": pd.Timestamp("2024-01-01"),
                "Days Billed": 730,
                "Net Charge (£)": 1347.96,
                "12-Month Limit (days)": 365,
                "Excess Days": 365,
                "Cancel/Rebill Admitted": False,
                "Reason Assessment": "back-billing",
            }
        ]
    )
    wb = Workbook()
    ws = wb.active
    # No evidence_df -- should not crash, col 11 is empty / placeholder.
    write_back_billing_sheet(ws, bb, account="A-31105244")
    hdr = ws.cell(row=7, column=11).value
    assert hdr == "Source Excerpt"
    body_excerpt = ws.cell(row=8, column=11).value
    assert body_excerpt in ("", None, "Source text unavailable")


def test_rebilling_sheet_emits_source_excerpt_column_header() -> None:
    reb = pd.DataFrame(
        [
            {
                "Killed Invoice #": "T12345",
                "Killer Invoice #": "T12346",
                "Killed Period From": pd.Timestamp("2024-01-01"),
                "Killed Period To": pd.Timestamp("2024-02-01"),
                "Killer Period From": pd.Timestamp("2024-01-01"),
                "Killer Period To": pd.Timestamp("2024-02-01"),
                "Killed Amount (£)": 100.00,
                "Killer Amount (£)": -100.00,
                "Reason Assessment": "reverse-and-rebill",
                "Days Billed (Killer)": 31,
            }
        ]
    )
    ev = _sample_evidence_df()
    wb = Workbook()
    ws = wb.active
    write_rebilling_sheet(ws, reb, account="A-31105244", evidence_df=ev)
    # Find the row carrying the header text 'Source Excerpt'.
    found = False
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == "Source Excerpt":
                found = True
                break
        if found:
            break
    assert found, "Expected a 'Source Excerpt' header somewhere on the rebilling sheet"
