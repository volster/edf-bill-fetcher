from __future__ import annotations

import pandas as pd

from edf_collector import detect_rebilling


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


def test_empty_df_returns_empty_df() -> None:
    out = detect_rebilling(pd.DataFrame())
    assert out.empty
    expected_cols = {
        "Killer Invoice",
        "Killed Invoice",
        "Killer Date",
        "Killed Date",
        "Period Overlap (days)",
        "Jump-back (days)",
        "Trigger Reason",
        "Cancel/Rebill Admitted (Killer)",
    }
    assert set(out.columns) == expected_cols


def test_non_overlapping_consecutive_invoices_not_flagged() -> None:
    df = pd.DataFrame(
        [
            _row("A", "01 Feb 2023", "01 Jan 2023", "31 Jan 2023"),
            _row("B", "01 Mar 2023", "01 Feb 2023", "28 Feb 2023"),
        ]
    )
    assert detect_rebilling(df).empty


def test_one_day_overlap_is_not_flagged_boundary() -> None:
    df = pd.DataFrame(
        [
            _row("A", "01 Feb 2023", "01 Jan 2023", "01 Feb 2023"),
            _row("B", "02 Feb 2023", "01 Feb 2023", "28 Feb 2023"),
        ]
    )
    assert detect_rebilling(df).empty


def test_60_day_overlap_emits_one_row() -> None:
    # Two 90-day invoices whose Period From ranges overlap by ~60 days
    # (Killer reaches back into Killed's window).
    df = pd.DataFrame(
        [
            _row("A", "01 Apr 2023", "01 Jan 2023", "31 Mar 2023"),  # Jan-Mar
            _row("B", "01 Jun 2023", "01 Feb 2023", "31 May 2023"),  # Feb-May
        ]
    )
    out = detect_rebilling(df)
    assert len(out) == 1
    row = out.iloc[0]
    assert row["Killer Invoice"] == "B"
    assert row["Killed Invoice"] == "A"
    # Jan-Mar vs Feb-May overlap is Feb-Mar, i.e. ~59 days.
    assert int(row["Period Overlap (days)"]) >= 30
    assert row["Trigger Reason"].startswith("period overlap")


def test_jumpback_90_days_emits_one_row() -> None:
    # Killed covers Feb. Killer covers Dec'22 -> Mar'23, so Killer's
    # Period From is ~90 days earlier than Killed's.
    df = pd.DataFrame(
        [
            _row("A", "01 Mar 2023", "01 Feb 2023", "28 Feb 2023"),
            _row("B", "01 Apr 2023", "01 Dec 2022", "31 Mar 2023"),
        ]
    )
    out = detect_rebilling(df)
    assert len(out) == 1
    row = out.iloc[0]
    assert int(row["Jump-back (days)"]) > 30
    assert row["Trigger Reason"].startswith("jump-back")


def test_long_period_killer_emits_one_row() -> None:
    # Killed covers Apr-Jun 2023. Killer covers Mar-May 2023 (90-day
    # killer, longer than 60) and Killer Period From is earlier than
    # Killed From. No overlap exceeds 30d, so the long-period rule must
    # fire.
    df = pd.DataFrame(
        [
            # Killed: issues 01 Jul, Apr-Jun.
            _row("A", "01 Jul 2023", "01 Apr 2023", "30 Jun 2023"),
            # Killer: issued later, LONG period (90 days), From is March 30
            # (just 2 days before killed's Apr 1, so jumpback=2 < 30, no
            # overlap rule, no jumpback rule -- only the long-period rule
            # fires).
            _row("B", "01 Aug 2023", "30 Mar 2023", "27 Jun 2023"),
        ]
    )
    out = detect_rebilling(df)
    assert len(out) == 1
    row = out.iloc[0]
    assert row["Killer Invoice"] == "B"
    assert "long period" in str(row["Trigger Reason"]).lower()


def test_cascade_emits_two_rows_for_three_invoices() -> None:
    df = pd.DataFrame(
        [
            _row("T65", "01 Apr 2023", "01 Feb 2023", "31 Mar 2023"),
            _row("T66", "01 Jun 2023", "01 Mar 2023", "31 May 2023"),
            _row("T67", "01 Aug 2023", "01 Apr 2023", "31 Jul 2023"),
            _row("T68", "01 Oct 2023", "01 May 2023", "30 Sep 2023"),
        ]
    )
    out = detect_rebilling(df)
    # Pairs (later, earlier): at minimum T68 vs T67, T68 vs T66.
    pairs = list(zip(out["Killer Invoice"], out["Killed Invoice"], strict=False))
    assert ("T68", "T67") in pairs
    # At least 2 rows triggered by overlap/jump-back
    assert len(out) >= 2


def test_unparseable_period_silently_skipped() -> None:
    df = pd.DataFrame(
        [
            _row("A", "01 Mar 2023", "garbage", "28 Feb 2023"),
            _row("B", "01 Apr 2023", "01 Feb 2023", "31 Mar 2023"),
        ]
    )
    out = detect_rebilling(df)
    # Only B has a parseable Period From/To; no other invoice to
    # pair against, so output is empty.
    assert out.empty


def test_output_sorted_by_killer_date() -> None:
    df = pd.DataFrame(
        [
            _row("LATE_K", "01 Dec 2023", "01 Jan 2022", "30 Nov 2023"),
            _row("EARLY_K", "01 Feb 2023", "01 Jan 2022", "31 Jan 2023"),
            _row("V0", "01 Oct 2022", "01 Sep 2022", "30 Sep 2022"),
        ]
    )
    out = detect_rebilling(df)
    # Earliest killer (by Date) first
    assert list(out["Killer Invoice"])[0] == "EARLY_K"
    assert list(out["Killer Invoice"])[-1] == "LATE_K"


def test_missing_admitted_column_treated_as_false() -> None:
    # Two invoices whose periods clearly overlap (Feb-Feb, 27 days),
    # no admit column provided.
    df = pd.DataFrame(
        [
            {
                "Invoice #": "A",
                "Date": "01 Mar 2023",
                "Period From": "01 Jan 2023",
                "Period To": "28 Feb 2023",
                "Amount (£)": 100.0,
            },
            {
                "Invoice #": "B",
                "Date": "01 Apr 2023",
                "Period From": "01 Dec 2022",
                "Period To": "31 Mar 2023",  # overlaps with A by Dec-Feb
                "Amount (£)": 100.0,
            },
        ]
    )
    out = detect_rebilling(df)
    assert len(out) == 1
    # Admit-tag should be False
    assert bool(out.iloc[0]["Cancel/Rebill Admitted (Killer)"]) is False
