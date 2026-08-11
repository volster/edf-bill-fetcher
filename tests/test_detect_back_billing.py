from __future__ import annotations

import pandas as pd

from edf_bill_fetcher.processors.detection import detect_back_billing


def _row(
    invoice: str = "T-001",
    date: str = "01 Jan 2024",
    period_from: str = "01 Jan 2023",
    period_to: str = "31 Dec 2023",
    amount: float = 1000.0,
    admitted: bool | None = None,
    attachment: str = "T-001.pdf",
    period_charge: object = None,
) -> dict:
    out = {
        "Invoice #": invoice,
        "Date": date,
        "Period From": period_from,
        "Period To": period_to,
        "Amount (£)": amount,
        "Attachment Name": attachment,
    }
    if period_charge is not None:
        out["Period Charge (£)"] = period_charge
    if admitted is not None:
        out["Cancel/Rebill Admitted"] = admitted
    return out


def test_empty_df_returns_empty_df() -> None:
    out = detect_back_billing(pd.DataFrame())
    assert out.empty


def test_short_period_invoice_not_flagged() -> None:
    # Short period billed within 12 months of its Period To -> not back-billing.
    df = pd.DataFrame(
        [_row(date="01 Jan 2024", period_from="01 Dec 2023", period_to="28 Dec 2023")]
    )
    assert detect_back_billing(df).empty


def test_inverted_period_from_after_to_is_skipped() -> None:
    # Period From (2023-12-31) is AFTER Period To (2023-01-01) — an inverted
    # period. The Date - Period From gate alone would pass (bill >365 days
    # after Period From), but the negative day span must be skipped silently.
    df = pd.DataFrame(
        [
            _row(
                invoice="INVERTED",
                date="2025-01-01",
                period_from="2023-12-31",
                period_to="2023-01-01",
                amount=1000.0,
            )
        ]
    )
    assert detect_back_billing(df).empty


def test_zero_day_period_is_skipped() -> None:
    # Period From == Period To (2024-01-01) — a zero-day period. The row must
    # be skipped silently rather than emitted with a zero Days Billed span.
    df = pd.DataFrame(
        [
            _row(
                invoice="ZERO-DAY",
                date="2025-01-01",
                period_from="2024-01-01",
                period_to="2024-01-01",
                amount=1000.0,
            )
        ]
    )
    assert detect_back_billing(df).empty


def test_exactly_365_days_is_not_flagged_boundary() -> None:
    # Bill date exactly 365 days after Period From is NOT back-billing (boundary).
    # Period 02 Jan 2024 to 31 Dec 2024; bill date 01 Jan 2025 -> 365 days gap
    # from Period From.
    df = pd.DataFrame(
        [_row(date="01 Jan 2025", period_from="02 Jan 2024", period_to="31 Dec 2024")]
    )
    assert detect_back_billing(df).empty


def test_366_days_is_flagged_boundary() -> None:
    # Bill date 366+ days after Period From IS back-billing (just over the boundary).
    # Period 01 Jan 2023 to 02 Jan 2024; bill date 03 Jan 2025 -> 733 days gap
    # from Period From.
    df = pd.DataFrame(
        [_row(date="03 Jan 2025", period_from="01 Jan 2023", period_to="02 Jan 2024")]
    )
    out = detect_back_billing(df)
    assert len(out) == 1
    assert int(out.loc[0, "Excess Days"]) > 0


def test_long_period_non_admitted_row() -> None:
    # 478-day period billed >365 days after its Period To -> back-billing.
    # Period 04 Apr 2022 to 26 Jul 2023; bill date 09 Aug 2024 (379 days after Period To).
    df = pd.DataFrame(
        [
            _row(
                invoice="T-6715690",
                date="09 Aug 2024",
                period_from="04 Apr 2022",
                period_to="26 Jul 2023",  # 478 days span
                amount=4401.07,
                admitted=False,
            )
        ]
    )
    out = detect_back_billing(df)
    assert len(out) == 1
    row = out.iloc[0]
    assert int(row["Days Billed"]) == 478
    assert int(row["Excess Days"]) > 0
    assert bool(row["Cancel/Rebill Admitted"]) is False
    assert float(row["Period Charge (£)"]) == 4401.07
    assert row["Invoice #"] == "T-6715690"


def test_long_period_admitted_row() -> None:
    # ~1015-day period billed >365 days after its Period To -> back-billing.
    # Period 01 Oct 2020 to 07 Jul 2023; bill date 09 Aug 2024 (398 days after Period To).
    df = pd.DataFrame(
        [
            _row(
                invoice="KI-0001",
                date="09 Aug 2024",
                period_from="01 Oct 2020",
                period_to="07 Jul 2023",  # 1015-ish days
                amount=1525.13,
                admitted=True,
            )
        ]
    )
    out = detect_back_billing(df)
    assert len(out) == 1
    row = out.iloc[0]
    assert bool(row["Cancel/Rebill Admitted"]) is True
    assert int(row["Excess Days"]) > 0
    assert isinstance(row["Reason Assessment"], str)
    assert len(row["Reason Assessment"]) > 20


def test_mix_of_normal_and_backbilled_only_backbilled_surfaces() -> None:
    # Normal rows: short periods billed within 12 months of Period To -> not flagged.
    # X row: long period billed >365 days after Period To -> flagged.
    rows = [
        _row(invoice="A", date="01 Feb 2023", period_from="01 Jan 2023", period_to="31 Jan 2023"),
        _row(invoice="B", date="01 Mar 2023", period_from="01 Feb 2023", period_to="28 Feb 2023"),
        _row(invoice="C", date="01 Apr 2023", period_from="01 Mar 2023", period_to="31 Mar 2023"),
        _row(invoice="D", date="01 May 2023", period_from="01 Apr 2023", period_to="30 Apr 2023"),
        _row(
            invoice="X",
            date="01 Jan 2025",  # 367 days after Period To -> flagged
            period_from="01 Jan 2022",
            period_to="31 Dec 2023",  # ~730 days span
            amount=5000.0,
        ),
    ]
    df = pd.DataFrame(rows)
    out = detect_back_billing(df)
    assert set(out["Invoice #"]) == {"X"}
    assert len(out) == 1


def test_unparseable_period_from_silently_skipped() -> None:
    df = pd.DataFrame(
        [
            _row(date="01 Jan 2025", period_from="N/A", period_to="31 Dec 2023"),
            _row(date="01 Jan 2025", period_from="garbage", period_to="31 Dec 2023"),
            _row(
                invoice="FLAG",
                date="01 Jan 2025",
                period_from="01 Jan 2022",
                period_to="31 Dec 2023",
            ),
        ]
    )
    out = detect_back_billing(df)
    assert set(out["Invoice #"]) == {"FLAG"}


def test_output_sorted_by_bill_date() -> None:
    rows = [
        _row(
            invoice="LATE",
            date="01 Dec 2024",
            period_from="01 Jan 2021",
            period_to="30 Nov 2023",
        ),
        _row(
            invoice="EARLY",
            date="01 Jan 2024",
            period_from="01 Jan 2021",
            period_to="31 Dec 2022",
        ),
    ]
    df = pd.DataFrame(rows)
    out = detect_back_billing(df)
    assert list(out["Invoice #"]) == ["EARLY", "LATE"]


def test_missing_admitted_column_treated_as_false() -> None:
    df = pd.DataFrame(
        [
            _row(
                invoice="NO_ADMIT_COL",
                date="01 Jan 2025",
                period_from="01 Jan 2022",
                period_to="31 Dec 2023",
            )
        ]
    )
    out = detect_back_billing(df)
    assert len(out) == 1
    assert bool(out.iloc[0]["Cancel/Rebill Admitted"]) is False


def test_output_columns_match_spec() -> None:
    df = pd.DataFrame(
        [_row(date="01 Jan 2025", period_from="01 Jan 2022", period_to="31 Dec 2023")]
    )
    out = detect_back_billing(df)
    expected = {
        "Invoice #",
        "Bill Date",
        "Period From",
        "Period To",
        "Days Billed",
        "Period Charge (£)",
        "Value Source",
        "12-Month Limit (days)",
        "Excess Days",
        "Unlawful Charge (£)",
        "Cancel/Rebill Admitted",
        "Reason Assessment",
    }
    assert set(out.columns) == expected


def test_unlawful_charge_is_prorated_share() -> None:
    # Period 01 Jan 2022 to 31 Dec 2023 = 729 days; bill date 01 Jan 2025.
    # Excess Days = (01 Jan 2025 - 365 days - 01 Jan 2022).days = 731 days.
    # Unlawful Charge = round(charge * (min(excess, days) / days), 2).
    # Note: excess (731) > days (729) here because the bill date is far
    # enough out that the entire period is unlawful — the ratio is capped
    # at 1.0 so the unlawful charge equals the full period charge.
    df = pd.DataFrame(
        [
            _row(
                invoice="UC-001",
                date="01 Jan 2025",
                period_from="01 Jan 2022",
                period_to="31 Dec 2023",
                amount=1000.0,
            )
        ]
    )
    out = detect_back_billing(df)
    assert len(out) == 1
    row = out.iloc[0]
    days = int(row["Days Billed"])
    excess = int(row["Excess Days"])
    charge = float(row["Period Charge (£)"])
    expected_unlawful = round(charge * (min(excess, days) / days), 2)
    assert float(row["Unlawful Charge (£)"]) == expected_unlawful
    # Sanity: unlawful charge is prorated by the excess/days ratio.
    assert excess > 0
    assert days > 0


def test_unlawful_charge_capped_at_full_charge_when_excess_exceeds_days() -> None:
    # Regression: excess (731) > days (729) previously inflated the
    # unlawful charge above 100% of the Period Charge (1002.74). The
    # proration ratio is capped at 1.0, so the unlawful charge must
    # never exceed the full Period Charge.
    df = pd.DataFrame(
        [
            _row(
                invoice="UC-CAP",
                date="01 Jan 2025",
                period_from="01 Jan 2022",
                period_to="31 Dec 2023",
                amount=1000.0,
            )
        ]
    )
    out = detect_back_billing(df)
    row = out.iloc[0]
    assert int(row["Excess Days"]) > int(row["Days Billed"])
    assert float(row["Unlawful Charge (£)"]) == 1000.0


# ---------------------------------------------------------------------------
# New tests: Period Charge (£) pull + Amount (£) fallback (from brief step 1)
# ---------------------------------------------------------------------------


def test_detect_back_billing_pulls_period_charge() -> None:
    df = pd.DataFrame(
        [
            {
                "Invoice #": "T99",
                "Date": "2022-06-15",  # >365 days after Period To
                "Period From": "2020-01-01",
                "Period To": "2021-06-01",
                "Period Charge (£)": 500.0,
                "Amount (£)": 100.0,  # running balance differs
            }
        ]
    )
    result = detect_back_billing(df)
    assert len(result) == 1
    assert "Period Charge (£)" in result.columns
    assert "Value Source" in result.columns
    assert result.iloc[0]["Period Charge (£)"] == 500.0
    assert result.iloc[0]["Value Source"] == "Period Charge"
    # Old column name gone
    assert "Net Charge (£)" not in result.columns


def test_detect_back_billing_fallback_to_amount() -> None:
    df = pd.DataFrame(
        [
            {
                "Invoice #": "T99",
                "Date": "2022-06-15",  # >365 days after Period To
                "Period From": "2020-01-01",
                "Period To": "2021-06-01",
                "Period Charge (£)": "N/A",  # not parseable
                "Amount (£)": 100.0,
            }
        ]
    )
    result = detect_back_billing(df)
    assert len(result) == 1
    assert result.iloc[0]["Period Charge (£)"] == 100.0
    assert "fallback" in result.iloc[0]["Value Source"].lower()


# ---------------------------------------------------------------------------
# New tests: legal Date-vs-Period-To rule (from prompt)
# ---------------------------------------------------------------------------


def test_short_period_billed_years_late_is_flagged() -> None:
    # 1-day period billed 5 years after Period To -> flagged, Excess Days = 1.
    df = pd.DataFrame(
        [
            _row(
                invoice="T-001",
                date="01 Jan 2025",  # ~5 years after Period To
                period_from="01 Jan 2020",
                period_to="02 Jan 2020",  # 1-day period
                amount=50.0,
            )
        ]
    )
    result = detect_back_billing(df)
    assert len(result) == 1
    assert int(result.iloc[0]["Excess Days"]) >= 1


def test_long_period_billed_within_year_of_period_to_not_flagged() -> None:
    # Long-ish period billed within 365 days of its Period From -> NOT back-billing.
    # Under the per-unit SLC 21BA test the gate is Date - Period From > 365.
    # Period 01 Jan 2023 to 30 Nov 2023 (333 days); bill date 31 Dec 2023
    # -> Date - Period From = 364 days -> not flagged.
    df = pd.DataFrame(
        [
            _row(
                invoice="T-002",
                date="31 Dec 2023",
                period_from="01 Jan 2023",
                period_to="30 Nov 2023",  # 333 days span
                amount=5000.0,
            )
        ]
    )
    result = detect_back_billing(df)
    assert result.empty
