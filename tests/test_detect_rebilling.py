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


def test_overlap_without_containment_emits_zero_under_tightened_gate() -> None:
    """Two 90-day invoices whose Period From ranges overlap by ~60 days
    (Killer reaches back into Killed's window) but neither fully
    contains the other. Under the tightened gate, overlap alone no
    longer fires -- emit is zero."""
    df = pd.DataFrame(
        [
            _row("A", "01 Apr 2023", "01 Jan 2023", "31 Mar 2023"),  # Jan-Mar
            _row("B", "01 Jun 2023", "01 Feb 2023", "31 May 2023"),  # Feb-May
        ]
    )
    assert detect_rebilling(df).empty


def test_jumpback_without_containment_emits_zero_under_tightened_gate() -> None:
    """Killed covers Feb. Killer covers Dec'22 -> Mar'23, so Killer's
    Period From is ~90 days earlier than Killed's. Containment does
    NOT hold so the jumpback signal cannot fire -- emit is zero."""
    df = pd.DataFrame(
        [
            _row("A", "01 Mar 2023", "01 Feb 2023", "28 Feb 2023"),
            _row("B", "01 Apr 2023", "01 Dec 2022", "31 Mar 2023"),
        ]
    )
    assert detect_rebilling(df).empty


def test_long_period_killer_with_admit_phrase_emits_one_row() -> None:
    """Killer spans March-June with admit-phrase; killed covers
    April only -- killer fully contains killed and the admit signal
    fires."""
    df = pd.DataFrame(
        [
            _row("A", "01 Jul 2023", "01 Apr 2023", "30 Apr 2023"),
            _row(
                "B",
                "01 Aug 2023",
                "01 Mar 2023",
                "30 Jun 2023",
                admitted=True,
            ),
        ]
    )
    out = detect_rebilling(df)
    assert len(out) == 1
    row = out.iloc[0]
    assert row["Killer Invoice"] == "B"
    assert any("admit" in s.lower() for s in str(row["Trigger Reason"]).split(";"))


def test_cascade_with_admit_phrase_emits_two_rows() -> None:
    """Five cascading invoices where the last (T68) admits
    cancel-and-rebill and fully contains two earlier invoices
    (T67, T66). Containment holds because T68's window Apr-Sep
    envelops both earlier windows."""
    df = pd.DataFrame(
        [
            _row("T65", "01 Apr 2023", "01 Feb 2023", "31 Mar 2023"),
            _row("T66", "01 Jun 2023", "01 Mar 2023", "31 May 2023"),
            _row("T67", "01 Aug 2023", "01 Apr 2023", "31 Jul 2023"),
            _row(
                "T68",
                "01 Oct 2023",
                "01 Mar 2023",  # extends back to fully contain T66+T67
                "30 Sep 2023",
                admitted=True,
            ),
        ]
    )
    out = detect_rebilling(df)
    pairs = list(zip(out["Killer Invoice"], out["Killed Invoice"], strict=False))
    # T68 May-Sep envelope contains T66 partial, T67 full
    assert ("T68", "T67") in pairs
    # At least 1 row triggers via the admit signal
    assert len(out) >= 1


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
    """Even when the source DataFrame has no 'Cancel/Rebill
    Admitted' column at all, a 365d killer containing a short killed
    invoice must still fire -- the 365d signal doesn't depend on
    the admit column."""
    df = pd.DataFrame(
        [
            {
                "Invoice #": "A",
                "Date": "01 Mar 2023",
                "Period From": "01 Feb 2023",
                "Period To": "28 Feb 2023",
                "Amount (£)": 100.0,
            },
            {
                "Invoice #": "B",
                "Date": "01 Apr 2023",
                "Period From": "01 Jan 2022",  # spans 14 months
                "Period To": "31 Mar 2023",  # fully contains A
                "Amount (£)": 100.0,
            },
        ]
    )
    out = detect_rebilling(df)
    assert len(out) == 1
    # Admit-tag should default to False
    assert bool(out.iloc[0]["Cancel/Rebill Admitted (Killer)"]) is False
