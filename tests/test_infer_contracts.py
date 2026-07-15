from __future__ import annotations

import pandas as pd

from edf_collector import infer_contracts


def _row(date: str, tariff: str = "Standard") -> dict:
    return {
        "Date": date,
        "Tariff": tariff,
        "Invoice #": f"INV-{date[-4:]}",
    }


def test_empty_df_returns_empty_df() -> None:
    out = infer_contracts(pd.DataFrame())
    assert out.empty
    expected_cols = {"Contract From", "Contract To", "Tariff", "Days", "# Invoices"}
    assert set(out.columns) == expected_cols


def test_constant_tariff_emits_single_contract_spanning_full_period() -> None:
    df = pd.DataFrame(
        [
            _row("01 Jan 2022"),
            _row("01 Feb 2022"),
            _row("01 Mar 2022"),
        ]
    )
    out = infer_contracts(df)
    assert len(out) == 1
    row = out.iloc[0]
    assert row["Tariff"] == "Standard"
    assert int(row["# Invoices"]) == 3
    assert int(row["Days"]) >= 59  # Jan 1 -> Mar 1 = 59 days


def test_tariff_change_produces_two_contract_rows() -> None:
    df = pd.DataFrame(
        [
            _row("01 Jan 2022", tariff="Standard"),
            _row("01 Feb 2022", tariff="Standard"),
            _row("01 Mar 2022", tariff="Fixed"),
            _row("01 Apr 2022", tariff="Fixed"),
        ]
    )
    out = infer_contracts(df)
    assert len(out) == 2
    tariffs = list(out["Tariff"])
    assert tariffs == ["Standard", "Fixed"]


def test_short_gap_does_not_merge_when_intervening_tariff_differs() -> None:
    # Spec: adjacent groups with gap < 30 days merge ONLY when they're
    # the same tariff AND there's no intervening different-tariff run.
    # Here Standard <- 37-day gap (Feb 1 -> Mar 10) with a Fixed blip
    # on Mar 1, so the two Standard runs must NOT merge -- 3 contracts.
    df = pd.DataFrame(
        [
            _row("01 Jan 2022", tariff="Standard"),
            _row("01 Feb 2022", tariff="Standard"),
            _row("01 Mar 2022", tariff="Fixed"),
            _row("10 Mar 2022", tariff="Standard"),
            _row("01 Apr 2022", tariff="Standard"),
        ]
    )
    out = infer_contracts(df)
    assert len(out) == 3
    tariffs = list(out["Tariff"])
    assert tariffs == ["Standard", "Fixed", "Standard"]


def test_short_gap_merges_same_tariff_adjacent_runs() -> None:
    # Two Standard runs separated by a < 30 day "gap" (no intervening
    # different-tariff run, just a temporal gap in the dataset).
    df = pd.DataFrame(
        [
            _row("01 Jan 2022", tariff="Standard"),
            _row("01 Feb 2022", tariff="Standard"),
            # Gap of 25 days between Feb 1 and Feb 26 -- still a Standard row
            _row("26 Feb 2022", tariff="Standard"),
            _row("01 Apr 2022", tariff="Standard"),
        ]
    )
    out = infer_contracts(df)
    # Same tariff throughout -> just one contract.
    assert len(out) == 1
    assert out.iloc[0]["Tariff"] == "Standard"
    assert int(out.iloc[0]["# Invoices"]) == 4


def test_na_tariff_skipped() -> None:
    df = pd.DataFrame(
        [
            _row("01 Jan 2022", tariff="N/A"),
            _row("01 Feb 2022", tariff="Standard"),
            _row("01 Mar 2022", tariff="Standard"),
        ]
    )
    out = infer_contracts(df)
    # N/A row is skipped, so only the two Standard rows form a contract.
    assert len(out) == 1
    assert out.iloc[0]["Tariff"] == "Standard"
    assert int(out.iloc[0]["# Invoices"]) == 2


def test_three_tariffs_produce_three_contracts() -> None:
    df = pd.DataFrame(
        [
            _row("01 Jan 2022", tariff="Old Variable"),
            _row("01 Feb 2022", tariff="Old Variable"),
            _row("01 Apr 2022", tariff="Old Variable"),
            _row("01 May 2022", tariff="Fixed 1Y"),
            _row("01 Jun 2022", tariff="Fixed 1Y"),
            _row("01 Aug 2022", tariff="Fixed 1Y"),
            _row("01 Sep 2022", tariff="New Variable"),
            _row("01 Oct 2022", tariff="New Variable"),
        ]
    )
    out = infer_contracts(df)
    assert len(out) == 3
    assert list(out["Tariff"]) == ["Old Variable", "Fixed 1Y", "New Variable"]


def test_output_sorted_by_contract_from() -> None:
    df = pd.DataFrame(
        [
            _row("01 Mar 2022", tariff="Mid_T"),
            _row("01 Jan 2022", tariff="Early_T"),
            _row("01 Feb 2022", tariff="Early_T"),
        ]
    )
    out = infer_contracts(df)
    # Order: earliest contract start (Jan 1) first.
    from_dts = pd.to_datetime(out["Contract From"], errors="coerce")
    assert from_dts.is_monotonic_increasing
    assert from_dts.iloc[0] == pd.Timestamp("2022-01-01")
