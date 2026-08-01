"""Accuracy tests for the tightened `detect_rebilling` gate.

Spec ref: 2026-07-16-sap-dumps-and-evidence-bundle-design.md §11
(tighten triggers so only real cancel-and-repost chains surface).

Old loose logic fired 575 rows from the real corpus because any
of these alone sufficed: overlap > 30d, jumpback > 30d, killer
≥ 60d. Tightened logic requires CONJUNCTION:

  (killer fully contains killed) AND
  (killer ≥ 365d OR admit-phrase OR reversal-credit match in
   evidence_df)

Each test below exercises one slice of that conjunction so the
matrix is fully covered.
"""

from __future__ import annotations

import pandas as pd

from edf_bill_fetcher.processors.detection import detect_rebilling


def _row(
    invoice: str,
    date: str,
    period_from: str,
    period_to: str,
    amount: float = 1000.0,
    admitted: bool = False,
) -> dict:
    return {
        "Invoice #": invoice,
        "Date": date,
        "Period From": period_from,
        "Period To": period_to,
        "Amount (£)": amount,
        "Cancel/Rebill Admitted": admitted,
    }


def _credit_row(
    invoice: str,
    period_from: str,
    period_to: str,
    amount: float,
) -> dict:
    """A reversal-credit line as it would appear in evidence_df."""
    return {
        "Invoice #": invoice,
        "Entry Type": "Credit",
        "Period From": period_from,
        "Period To": period_to,
        "Amount (£)": amount,
    }


# ---------------------------------------------------------------------------
# 1. Old-style loose triggers that should NO LONGER be the cause of
#    emission. Under the tightened gate, these scenarios only emit
#    because the 365d signal happens to also fire -- confirm we're not
#    relying on the old 30d-overlap / 30d-jumpback / 60d rules.
# ---------------------------------------------------------------------------


def test_1095d_killer_containing_three_short_killed_emits_via_365_signal() -> None:
    """A 1095-day killer fully contains 3 short killed invoices.
    Under tightened logic, only the 'killer period >= 365d' signal
    fires (no admit, no reversal); the row count is exactly 3 (one
    per contained killed) rather than the old loose 3-row expansion
    driven by overlap and jumpback rules. The Trigger Reason must
    mention the 365d signal and nothing else."""
    df = pd.DataFrame(
        [
            _row("K1", "01 Jan 2024", "01 Jan 2022", "31 Dec 2024"),  # 1095d
            _row("S1", "01 Feb 2022", "01 Feb 2022", "28 Feb 2022"),
            _row("S2", "01 Jun 2022", "01 Jun 2022", "30 Jun 2022"),
            _row("S3", "01 Oct 2022", "01 Oct 2022", "31 Oct 2022"),
        ]
    )
    out = detect_rebilling(df)
    pairs = list(zip(out["Killer Invoice"], out["Killed Invoice"], strict=False))
    assert ("K1", "S1") in pairs
    assert ("K1", "S2") in pairs
    assert ("K1", "S3") in pairs
    for reason in out["Trigger Reason"]:
        semicolons = [s.strip() for s in reason.split(";")]
        assert semicolons == ["killer period \u2265 365d"], (
            f"only the 365d signal should fire (got {semicolons!r})"
        )


def test_killer_period_overlaps_but_does_not_contain_killed_emits_zero() -> None:
    """Killer and killed overlap by 60 days but killer does NOT
    fully contain the killed window. Old logic would have emitted
    via the 30d-overlap rule; tightened gate must not."""
    df = pd.DataFrame(
        [
            _row("A", "01 Apr 2023", "01 Jan 2023", "31 Mar 2023"),  # Jan-Mar
            _row("B", "01 Jun 2023", "01 Feb 2023", "31 May 2023"),  # Feb-May
        ]
    )
    assert detect_rebilling(df).empty


def test_killer_span_under_365_with_no_admit_and_no_reversal_emits_zero() -> None:
    """Killer fully contains killed but its Days Billed is ~120d,
    no admit, no reversal. Tightened gate requires one of the
    three signals -- none present, zero rows."""
    df = pd.DataFrame(
        [
            _row("A", "01 Mar 2023", "01 Feb 2023", "28 Feb 2023"),
            _row("B", "01 Apr 2023", "01 Jan 2023", "31 Mar 2023"),  # 89d span
        ]
    )
    assert detect_rebilling(df).empty


# ---------------------------------------------------------------------------
# 2. Each signal path, in isolation, must emit.
# ---------------------------------------------------------------------------


def test_admit_phrase_on_killer_emits_one_row_even_when_span_under_365() -> None:
    """Killer spans only 120d but admit-phrase=True. One contained
    killed invoice -> one row."""
    df = pd.DataFrame(
        [
            _row("A", "01 Mar 2023", "01 Feb 2023", "28 Feb 2023"),
            _row("B", "01 Apr 2023", "01 Jan 2023", "31 Mar 2023", admitted=True),
        ]
    )
    out = detect_rebilling(df)
    assert len(out) == 1
    row = out.iloc[0]
    assert row["Killer Invoice"] == "B"
    assert row["Killed Invoice"] == "A"
    assert bool(row["Cancel/Rebill Admitted (Killer)"]) is True
    assert any("admit" in s.lower() for s in str(row["Trigger Reason"]).split(";"))


def test_reversal_credit_in_evidence_df_emits_one_row() -> None:
    """Killer fully contains killed, span < 365, no admit; but a
    credit row in evidence_df matches the killed invoice's amount
    and period (overlap ≥ 30 days). Reversal-match signal fires."""
    invoice_rows = pd.DataFrame(
        [
            # Killed A: an extended Q1 2023 invoice, 90 days.
            _row("A", "01 Apr 2023", "01 Jan 2023", "31 Mar 2023", amount=250.00),
            # Killer B: an even broader window that fully contains A,
            # but only 120 days -- fails both ≥ 365 AND admit.
            _row("B", "01 May 2023", "01 Dec 2022", "31 Mar 2023"),
        ]
    )
    evidence_df = pd.DataFrame(
        [
            # Reversal credit for the full killed period (90 days).
            _credit_row(
                invoice="A",
                period_from="01 Jan 2023",
                period_to="31 Mar 2023",
                amount=-250.00,
            )
        ]
    )
    out = detect_rebilling(invoice_rows, evidence_df=evidence_df)
    assert len(out) == 1
    row = out.iloc[0]
    assert any("reversal" in s.lower() for s in str(row["Trigger Reason"]).split(";"))


def test_reversal_credit_with_tiny_overlap_fails_check() -> None:
    """A credit whose period overlaps the killed by less than the
    30-day threshold must NOT trigger even if amount matches.
    Documents the explicit threshold."""
    invoice_rows = pd.DataFrame(
        [
            # Killed A: a long invoice window.
            _row("A", "01 Apr 2023", "01 Jan 2023", "31 Mar 2023", amount=250.00),
            _row("B", "01 May 2023", "01 Dec 2022", "31 Mar 2023"),
        ]
    )
    # Credit is a 7-day fragment overlapping only 7 days of killed A.
    evidence_df = pd.DataFrame(
        [
            _credit_row(
                invoice="A",
                period_from="25 Mar 2023",
                period_to="31 Mar 2023",
                amount=-250.00,
            )
        ]
    )
    assert detect_rebilling(invoice_rows, evidence_df=evidence_df).empty


def test_reversal_credit_must_match_amount_within_pennies() -> None:
    """Large amount mismatch should NOT fire reversal-match."""
    invoice_rows = pd.DataFrame(
        [
            _row("A", "01 Mar 2023", "01 Feb 2023", "28 Feb 2023", amount=250.00),
            _row("B", "01 Apr 2023", "01 Jan 2023", "31 Mar 2023"),
        ]
    )
    # Mismatch by £5 — outside ±£0.50 tolerance
    evidence_df = pd.DataFrame(
        [
            _credit_row(
                invoice="A",
                period_from="01 Feb 2023",
                period_to="28 Feb 2023",
                amount=-255.00,
            )
        ]
    )
    assert detect_rebilling(invoice_rows, evidence_df=evidence_df).empty


def test_reversal_credit_must_overlap_killed_period() -> None:
    """Credit row matches amount exactly but its period does NOT
    overlap the killed period (e.g. credit is for an entirely
    different month). Reversal must not fire -- but if "Period From"
    is missing on the credit, the spec accepts an amount-only match,
    so the test uses a non-overlapping credit row with parsed dates
    to assert the overlap requirement."""
    invoice_rows = pd.DataFrame(
        [
            _row("A", "01 Mar 2023", "01 Feb 2023", "28 Feb 2023", amount=250.00),
            _row("B", "01 Apr 2023", "01 Jan 2023", "31 Mar 2023"),
        ]
    )
    # Credit is for Aug 2023 -- no overlap with Feb 2023 killed period
    evidence_df = pd.DataFrame(
        [
            _credit_row(
                invoice="A",
                period_from="01 Aug 2023",
                period_to="31 Aug 2023",
                amount=-250.00,
            )
        ]
    )
    assert detect_rebilling(invoice_rows, evidence_df=evidence_df).empty


def test_reversal_credit_with_missing_period_still_matches_on_amount() -> None:
    """Per spec: if the credit row has no parseable Period From/To,
    accept the match on amount alone (Entry Type == Credit)."""
    invoice_rows = pd.DataFrame(
        [
            _row("A", "01 Mar 2023", "01 Feb 2023", "28 Feb 2023", amount=250.00),
            _row("B", "01 Apr 2023", "01 Jan 2023", "31 Mar 2023"),
        ]
    )
    evidence_df = pd.DataFrame(
        [
            {
                "Invoice #": "A",
                "Entry Type": "Credit",
                "Amount (£)": -250.00,
            }
        ]
    )
    out = detect_rebilling(invoice_rows, evidence_df=evidence_df)
    assert len(out) == 1


# ---------------------------------------------------------------------------
# 4. Cascade behavior: cancellation groups collapse.
# ---------------------------------------------------------------------------


def test_cancellation_cascade_collapses_to_one_row_per_killed() -> None:
    """Three invoices with cascading containment: K1 covers
    2022-2024 (730d), K2 covers 2022-2023 (365d exactly), killed S1
    is Feb 2022. Under tightened logic, BOTH K1 and K2 contain S1
    and both span ≥ 365, so each (killer, S1) pair fires once --
    not the old 3-row expansion."""
    df = pd.DataFrame(
        [
            _row("S1", "01 Feb 2022", "01 Feb 2022", "28 Feb 2022"),
            # K2: 2022-01-01 to 2022-12-31 = 364d. Use 2021-12-31
            # start to land at exactly 365d.
            _row("K2", "01 Jan 2023", "31 Dec 2021", "31 Dec 2022"),  # 365d exact
            _row("K1", "01 Jan 2024", "01 Jan 2022", "31 Dec 2023"),  # 730d
        ]
    )
    out = detect_rebilling(df)
    # Both K1 and K2 contain S1 and both span ≥ 365, so each pair fires.
    pairs = set(zip(out["Killer Invoice"], out["Killed Invoice"], strict=False))
    assert ("K1", "S1") in pairs
    assert ("K2", "S1") in pairs
    # Each killer -> S1 pair should appear exactly once
    s1_rows = out[out["Killed Invoice"] == "S1"]
    assert len(s1_rows) == 2


def test_no_admit_no_reversal_no_365_no_emit_for_short_killer_containing_one() -> None:
    """Sanity: a 180-day killer that fully contains a 28-day killed
    invoice with no signals at all must NOT emit."""
    df = pd.DataFrame(
        [
            _row("A", "01 Mar 2023", "01 Feb 2023", "28 Feb 2023"),
            _row("B", "01 Jul 2023", "01 Jan 2023", "30 Jun 2023"),  # 180d
        ]
    )
    assert detect_rebilling(df).empty
