from __future__ import annotations

import pandas as pd

from edf_bill_fetcher.processors.detection import detect_back_billing

T68_SUB_PERIODS = (
    "02/10/2020|24/03/2021|19743.0|16.42|3241.8; "
    "25/03/2021|06/04/2021|1454.0|16.42|238.75; "
    "07/04/2021|31/03/2022|37184.0|16.42|6105.61; "
    "01/04/2022|12/05/2022|3736.0|52.00|1942.72; "
    "13/05/2022|31/03/2023|30675.0|52.00|15951.0; "
    "01/04/2023|09/08/2023|10607.0|45.92|4870.73"
)


def _t68_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Invoice #": "T78701920068",
                "Date": "09 Aug 2023",
                "Period From": "02 Oct 2020",
                "Period To": "09 Aug 2023",
                "Amount (£)": 32876.86,
                "Period Charge (£)": 1525.13,
                "Cancel/Rebill Admitted": True,
                "Sub Periods": T68_SUB_PERIODS,
            }
        ]
    )


def test_t68_unlawful_from_sub_periods() -> None:
    out = detect_back_billing(_t68_df())
    row = out.iloc[0]
    # fully-unlawful 02 Oct 20 -> 12 May 22 sub-periods + straddling
    # 13 May 22 -> 31 Mar 23 slice prorated at the 09/08/2022 cutoff.
    # 3241.80 + 238.75 + 6105.61 + 1942.72 + 15951.00 * 88/322
    expected = round(3241.80 + 238.75 + 6105.61 + 1942.72 + 15951.00 * (88 / 322), 2)
    assert abs(row["Unlawful Charge (£)"] - expected) < 0.01
    assert row["Sub-Period Basis"] == "Sub-period × rate"


def test_no_sub_periods_uses_day_ratio_fallback() -> None:
    df = _t68_df().drop(columns=["Sub Periods"])
    out = detect_back_billing(df)
    row = out.iloc[0]
    assert row["Sub-Period Basis"] == "Day-ratio fallback"
    assert row["Unlawful Charge (£)"] == round(1525.13 * (676 / 1041), 2)
