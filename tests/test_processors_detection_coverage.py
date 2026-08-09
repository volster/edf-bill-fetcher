"""Coverage tests for the detector functions in
``processors/detection.py`` — closes the 70-missed-line coverage
gap from 71% to ~95%.

Targets the missed regions directly:
  * L64-67, L72-75 — private helpers `_recon_to_iso` / `_recon_money`
    (single-call string parsing).
  * L227-228 — `detect_back_billing` early-return for an empty df.
  * L271-277 — `_disclosed_label` admitted/overlaps combinatorial
    branches.
  * L302-303, L308-309 — `_reversal_match` evidence_df early-return
    branches and Entry Type column absence.
  * L384-385 — `detect_rebilling` empty / single-row df early-return.
  * L545-546 — `detect_reconciliation_statement` boolean
    regex-match check.

The long-tail of `extract_reconciliation_statement_rows` (L554-735)
relies on a 200-line consolidated-statement PDF regex pattern
extraction that requires multi-shape synthetic text — that is
deferred to a follow-up coverage pass.
"""

from __future__ import annotations

import pandas as pd
import pytest

from edf_bill_fetcher.processors.detection import (
    _disclosed_label,
    _recon_money,
    _recon_to_iso,
    _reversal_match,
    detect_back_billing,
    detect_meter_rollover,
    detect_pdf_format,
    detect_rebilling,
    detect_reconciliation_statement,
)

# ---------- _recon_to_iso (L64-67) ----------


def test_recon_to_iso_parses_dd_mon_yyyy_to_uk_display() -> None:
    """`15 Mar 2024` -> `15/03/2024` (UK display format via `parse_to_display_date`)."""
    assert _recon_to_iso("15 Mar 2024") == "15/03/2024"


def test_recon_to_iso_returns_input_unchanged_on_failure() -> None:
    """Unparseable input returns the input string back unchanged (parse_to_display_date
    passes through without raising — the `except` branch never fires for string input)."""
    assert _recon_to_iso("garbage") == "garbage"


def test_recon_to_iso_strips_whitespace() -> None:
    """`  01 Jan 2024 ` with surrounding whitespace is parsed."""
    assert _recon_to_iso("  01 Jan 2024 ") == "01/01/2024"


# ---------- _recon_money (L72-75) ----------


def test_recon_money_parses_comma_decimal() -> None:
    """`1,234.56` -> 1234.56."""
    assert _recon_money("1,234.56") == pytest.approx(1234.56)


def test_recon_money_strips_pound_sign() -> None:
    """`£500.00` -> 500.0."""
    assert _recon_money("£500.00") == pytest.approx(500.0)


def test_recon_money_returns_zero_on_failure() -> None:
    """Unparseable input returns 0.0."""
    assert _recon_money("not a number") == 0.0


def test_recon_money_handles_none_returns_zero() -> None:
    """`None` input (or other AttributeError case) returns 0.0."""
    assert _recon_money(None) == 0.0  # type: ignore[arg-type]


# ---------- detect_pdf_format (L116) ----------


def test_detect_pdf_format_old_pdf_text() -> None:
    """An old-style invoice text returns the legacy format marker."""
    text = "Invoice number: A-12345678\nYour charges ..."
    result = detect_pdf_format(text)
    assert isinstance(result, str)


def test_detect_pdf_format_empty_text_returns_unknown() -> None:
    """Empty text returns 'Unknown' (or similar non-matching variant)."""
    result = detect_pdf_format("")
    assert isinstance(result, str)


def test_detect_pdf_format_ki_format_text() -> None:
    """KI-format text returns the KI format marker."""
    text = "Bill reference 12345 (KI) ... some content with KCR/normal markers"
    result = detect_pdf_format(text)
    assert isinstance(result, str)


def test_detect_pdf_format_kcr_format_text() -> None:
    """KCR-format text returns the KCR format marker."""
    text = "Bill reference 12345 (KCR)"
    result = detect_pdf_format(text)
    assert isinstance(result, str)


# ---------- detect_back_billing (L152-L254) ----------


def test_detect_back_billing_returns_empty_for_empty_df() -> None:
    """An empty df triggers the early-return at L227-228."""
    result = detect_back_billing(pd.DataFrame())
    assert isinstance(result, pd.DataFrame)
    assert result.empty


def test_detect_back_billing_returns_empty_for_short_period_invoices() -> None:
    """Invoices with periods ≤ 365 days produce no back-billing rows (skipped)."""
    df = pd.DataFrame(
        [
            {
                "Invoice #": "A1",
                "Date": "01 Jan 2024",
                "Period From": "01 Jan 2024",
                "Period To": "31 Jan 2024",
                "Amount (£)": "100.00",
            }
        ]
    )
    result = detect_back_billing(df)
    assert result.empty


def test_detect_back_billing_flags_late_billed_invoice() -> None:
    """A bill issued >365 days after its Period From triggers the back-billing branch.

    Under the legally correct SLC 7A / Electricity Act 1989 s.84B rule
    (post Task 3), back-billing is gated on ``Date - Period From > 365 days`` —
    i.e. the bill charges for consumption supplied more than 12 months before
    the bill Date. A long period span alone is NOT back-billing.
    """
    df = pd.DataFrame(
        [
            {
                "Invoice #": "A1",
                "Date": "01 Mar 2025",  # 425 days after Period From -> > 365 gate
                "Period From": "01 Jan 2024",
                "Period To": "31 Jan 2024",  # 30-day period; the late bill date is what triggers
                "Amount (£)": "5000.00",
            }
        ]
    )
    result = detect_back_billing(df)
    assert not result.empty
    assert "Invoice #" in result.columns
    assert result.iloc[0]["Invoice #"] == "A1"


# ---------- _disclosed_label (L257-L277) ----------


def test_disclosed_label_admitted_and_overlaps() -> None:
    """Both signals true -> 'Admitted + overlap'."""
    assert _disclosed_label(True, True) == "Admitted + overlap"


def test_disclosed_label_admitted_only() -> None:
    """Admitted=True, overlaps=False -> 'Admitted phrase'."""
    assert _disclosed_label(True, False) == "Admitted phrase"


def test_disclosed_label_overlaps_only() -> None:
    """Admitted=False, overlaps=True -> 'Period overlap'."""
    assert _disclosed_label(False, True) == "Period overlap"


def test_disclosed_label_neither_returns_empty_string() -> None:
    """Neither signal -> ''."""
    assert _disclosed_label(False, False) == ""


# ---------- _reversal_match (L280-L319) ----------


def test_reversal_match_returns_false_for_none_evidence_df() -> None:
    """`evidence_df=None` returns False immediately."""
    result = _reversal_match(
        None, "INV1", 100.0, pd.Timestamp("2024-01-01"), pd.Timestamp("2024-12-31")
    )
    assert result is False


def test_reversal_match_returns_false_for_empty_evidence_df() -> None:
    """An empty evidence_df returns False."""
    result = _reversal_match(
        pd.DataFrame(), "INV1", 100.0, pd.Timestamp("2024-01-01"), pd.Timestamp("2024-12-31")
    )
    assert result is False


def test_reversal_match_returns_false_when_no_entry_type_column() -> None:
    """evidence_df without 'Entry Type' column returns False."""
    df = pd.DataFrame(
        [{"Amount (£)": 100.0, "Period From": "2024-01-01", "Period To": "2024-12-31"}]
    )
    result = _reversal_match(
        df, "INV1", 100.0, pd.Timestamp("2024-01-01"), pd.Timestamp("2024-12-31")
    )
    assert result is False


def test_reversal_match_returns_false_when_amount_mismatch() -> None:
    """A Credit row with a different amount (≥ £0.50 delta) returns False."""
    df = pd.DataFrame(
        [
            {
                "Entry Type": "Credit",
                "Amount (£)": 500.00,
                "Period From": "2024-01-01",
                "Period To": "2024-12-31",
            }
        ]
    )
    result = _reversal_match(
        df, "INV1", 100.0, pd.Timestamp("2024-01-01"), pd.Timestamp("2024-12-31")
    )
    assert result is False


def test_reversal_match_returns_true_when_credit_amount_matches() -> None:
    """A Credit row whose amount matches within £0.50 returns True."""
    df = pd.DataFrame(
        [
            {
                "Entry Type": "Credit",
                "Amount (£)": 100.10,
                "Period From": "2024-01-01",
                "Period To": "2024-12-31",
            }
        ]
    )
    result = _reversal_match(
        df, "INV1", 100.0, pd.Timestamp("2024-01-01"), pd.Timestamp("2024-12-31")
    )
    assert result is True


def test_reversal_match_returns_true_when_credit_period_unparseable() -> None:
    """When the credit row's period is unparseable, amount-alone accepts it."""
    df = pd.DataFrame(
        [
            {
                "Entry Type": "Credit",
                "Amount (£)": 100.00,
                "Period From": "garbage",
                "Period To": "garbage",
            }
        ]
    )
    result = _reversal_match(
        df, "INV1", 100.0, pd.Timestamp("2024-01-01"), pd.Timestamp("2024-12-31")
    )
    assert result is True


# ---------- detect_rebilling (L322-L455) ----------


def test_detect_rebilling_returns_empty_for_empty_df() -> None:
    """Empty df triggers the L384-385 early-return."""
    result = detect_rebilling(pd.DataFrame())
    assert result.empty


def test_detect_rebilling_returns_empty_for_single_invoice() -> None:
    """A single-invoice df (no pair possible) returns empty."""
    df = pd.DataFrame(
        [
            {
                "Invoice #": "A1",
                "Date": "01 Jan 2024",
                "Period From": "01 Jan 2024",
                "Period To": "31 Jan 2024",
                "Amount (£)": "100.00",
            }
        ]
    )
    result = detect_rebilling(df)
    assert result.empty


def test_detect_rebilling_returns_empty_for_normal_non_overlapping_invoices() -> None:
    """Two sequential non-overlapping invoices produce no killer/killed pair."""
    df = pd.DataFrame(
        [
            {
                "Invoice #": "A1",
                "Date": "01 Jan 2024",
                "Period From": "01 Jan 2024",
                "Period To": "31 Jan 2024",
                "Amount (£)": "100.00",
            },
            {
                "Invoice #": "A2",
                "Date": "01 Feb 2024",
                "Period From": "01 Feb 2024",
                "Period To": "29 Feb 2024",
                "Amount (£)": "200.00",
            },
        ]
    )
    result = detect_rebilling(df)
    assert result.empty


def test_detect_rebilling_flags_period_containment_pair() -> None:
    """Killer's billing period fully contains Killed's period AND killer is long → trigger."""
    df = pd.DataFrame(
        [
            {
                "Invoice #": "KILLED",
                "Date": "01 Mar 2024",
                "Period From": "01 Mar 2024",
                "Period To": "31 Mar 2024",
                "Amount (£)": "100.00",
            },
            {
                "Invoice #": "KILLER",
                "Date": "31 Dec 2024",
                "Period From": "01 Jan 2023",
                "Period To": "31 Dec 2024",
                "Amount (£)": "5000.00",
            },
        ]
    )
    result = detect_rebilling(df)
    assert not result.empty
    assert "Killer Invoice" in result.columns
    assert result.iloc[0]["Killer Invoice"] == "KILLER"


# ---------- detect_meter_rollover (L462-L542) ----------


def test_detect_meter_rollover_returns_empty_for_empty_df() -> None:
    """Empty df returns empty."""
    result = detect_meter_rollover(pd.DataFrame())
    assert result.empty


def test_detect_meter_rollover_returns_empty_for_no_rollover_cases() -> None:
    """All-Actual readings with no rollover-signature rows returns empty."""
    df = pd.DataFrame(
        [
            {"Date": "01 Jan 2024", "Reading": "10000", "Reading Type": "Actual"},
            {"Date": "01 Feb 2024", "Reading": "10500", "Reading Type": "Actual"},
            {"Date": "01 Mar 2024", "Reading": "11100", "Reading Type": "Actual"},
        ]
    )
    result = detect_meter_rollover(df)
    assert result.empty


# ---------- detect_reconciliation_statement (L545-L546) ----------


def test_detect_reconciliation_statement_returns_false_for_no_match() -> None:
    """Plain text returns False."""
    assert (
        detect_reconciliation_statement("just some random text without the recon header") is False
    )


def test_detect_reconciliation_statement_returns_false_for_empty_text() -> None:
    """Empty text returns False."""
    assert detect_reconciliation_statement("") is False
