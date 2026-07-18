from __future__ import annotations

import pandas as pd

from edf_collector import run_analysers


def _df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Invoice #": "T-X1",
                "Date": "01 Aug 2023",
                "Period From": "01 Jan 2022",
                "Period To": "31 Jul 2023",  # long period
                "Reading": "Actual",
                "Units (kWh)": 300.0,
                "Amount (£)": 1000.0,
                "Tariff": "Standard",
                "Cancel/Rebill Admitted": True,
            },
            {
                "Invoice #": "T-X2",
                "Date": "01 Sep 2023",
                "Period From": "01 Jan 2022",  # fully contains T-X1 from 2022-01-01
                "Period To": "31 Aug 2023",
                "Reading": "Actual",
                "Units (kWh)": 400.0,
                "Amount (£)": 1500.0,
                "Tariff": "Standard",
                "Cancel/Rebill Admitted": False,
            },
        ]
    )


def test_run_analysers_returns_dict_with_four_keys() -> None:
    out = run_analysers(_df())
    # Four analyser frames plus the new ``evidence_index`` map (Stream P4).
    assert set(out.keys()) == {
        "back_billing",
        "rebilling",
        "meter_rollover",
        "contracts",
        "evidence_index",
    }


def test_run_analysers_back_billing_value_is_dataframe() -> None:
    out = run_analysers(_df())
    assert isinstance(out["back_billing"], pd.DataFrame)
    assert isinstance(out["rebilling"], pd.DataFrame)
    assert isinstance(out["meter_rollover"], pd.DataFrame)
    assert isinstance(out["contracts"], pd.DataFrame)
    # Stream P4: ``evidence_index`` is a dict[str, int] (possibly empty).
    assert isinstance(out["evidence_index"], dict)


def test_run_analysers_blank_input_returns_four_empty_dataframes() -> None:
    out = run_analysers(pd.DataFrame())
    for k in ("back_billing", "rebilling", "meter_rollover", "contracts"):
        assert out[k].empty, f"{k} should be empty"


def test_run_analysers_with_real_rows_finds_expected_events() -> None:
    # Two long-period invoices overlap > 30 days -> both back-billed,
    # one rebilling (Killer=T-X2 Killed=T-X1).
    out = run_analysers(_df())
    assert len(out["back_billing"]) == 2
    assert len(out["rebilling"]) >= 1
    assert out["meter_rollover"].empty
    assert len(out["contracts"]) >= 1  # at least one Standard contract


def test_run_analysers_preserves_input_df_unchanged() -> None:
    df = _df()
    df_copy = df.copy(deep=True)
    _ = run_analysers(df)
    pd.testing.assert_frame_equal(df, df_copy)
