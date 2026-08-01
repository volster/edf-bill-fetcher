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
) -> dict:
    out = {
        "Invoice #": invoice,
        "Date": date,
        "Period From": period_from,
        "Period To": period_to,
        "Amount (£)": amount,
        "Attachment Name": attachment,
    }
    if admitted is not None:
        out["Cancel/Rebill Admitted"] = admitted
    return out


def test_empty_df_returns_empty_df() -> None:
    out = detect_back_billing(pd.DataFrame())
    assert out.empty


def test_short_period_invoice_not_flagged() -> None:
    df = pd.DataFrame([_row(period_from="01 Dec 2023", period_to="28 Dec 2023")])
    assert detect_back_billing(df).empty


def test_exactly_365_days_is_not_flagged_boundary() -> None:
    df = pd.DataFrame([_row(period_from="01 Jan 2023", period_to="31 Dec 2023")])
    assert detect_back_billing(df).empty


def test_366_days_is_flagged_boundary() -> None:
    df = pd.DataFrame(
        [_row(period_from="01 Jan 2023", period_to="02 Jan 2024")]
    )  # 366 days (leap year 2024)
    out = detect_back_billing(df)
    assert len(out) == 1
    assert int(out.loc[0, "Excess Days"]) == 1


def test_long_period_non_admitted_row() -> None:
    df = pd.DataFrame(
        [
            _row(
                invoice="T-6715690",
                date="09 Aug 2023",
                period_from="04 Apr 2022",
                period_to="26 Jul 2023",  # 478 days
                amount=4401.07,
                admitted=False,
            )
        ]
    )
    out = detect_back_billing(df)
    assert len(out) == 1
    row = out.iloc[0]
    assert int(row["Days Billed"]) == 478
    assert int(row["Excess Days"]) == 113
    assert bool(row["Cancel/Rebill Admitted"]) is False
    assert float(row["Net Charge (£)"]) == 4401.07
    assert row["Invoice #"] == "T-6715690"


def test_long_period_admitted_row() -> None:
    df = pd.DataFrame(
        [
            _row(
                invoice="KI-0001",
                date="09 Aug 2023",
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
    assert int(row["Excess Days"]) > 600
    assert isinstance(row["Reason Assessment"], str)
    assert len(row["Reason Assessment"]) > 20


def test_mix_of_normal_and_backbilled_only_backbilled_surfaces() -> None:
    rows = [
        _row(invoice="A", period_from="01 Jan 2023", period_to="31 Jan 2023"),
        _row(invoice="B", period_from="01 Feb 2023", period_to="28 Feb 2023"),
        _row(invoice="C", period_from="01 Mar 2023", period_to="31 Mar 2023"),
        _row(invoice="D", period_from="01 Apr 2023", period_to="30 Apr 2023"),
        _row(
            invoice="X",
            period_from="01 Jan 2022",
            period_to="31 Dec 2023",  # ~730 days
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
            _row(period_from="N/A", period_to="31 Dec 2023"),
            _row(period_from="garbage", period_to="31 Dec 2023"),
            _row(
                invoice="FLAG",
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
            date="01 Dec 2023",
            period_from="01 Jan 2021",
            period_to="30 Nov 2023",
        ),
        _row(
            invoice="EARLY",
            date="01 Jan 2023",
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
                period_from="01 Jan 2022",
                period_to="31 Dec 2023",
            )
        ]
    )
    out = detect_back_billing(df)
    assert len(out) == 1
    assert bool(out.iloc[0]["Cancel/Rebill Admitted"]) is False


def test_output_columns_match_spec() -> None:
    df = pd.DataFrame([_row(period_from="01 Jan 2022", period_to="31 Dec 2023")])
    out = detect_back_billing(df)
    expected = {
        "Invoice #",
        "Bill Date",
        "Period From",
        "Period To",
        "Days Billed",
        "Net Charge (£)",
        "12-Month Limit (days)",
        "Excess Days",
        "Cancel/Rebill Admitted",
        "Reason Assessment",
    }
    assert set(out.columns) == expected
