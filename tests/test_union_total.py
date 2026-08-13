from __future__ import annotations

import pandas as pd

from edf_bill_fetcher.processors.detection import (
    compute_unlawful_union_total,
    detect_back_billing,
)

# T67 (bill 13 Jul 2023) recovers 15 Apr 22 - 03 Jul 23; T68 (bill 09 Aug 2023)
# recovers 02 Oct 20 - 09 Aug 23.  Their unlawful windows overlap on the days
# both invoices first recovered before each invoice's own 365-day cutoff.
T67_SUB = (
    "15/04/2022|12/05/2022|2468.0|52.00|1283.36; "
    "13/05/2022|31/03/2023|30675.0|52.00|15951.0; "
    "01/04/2023|03/07/2023|7547.0|45.92|3465.58"
)


def _df(invoice: str, date: str, pf: str, pt: str, sub_periods: str) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Invoice #": invoice,
                "Date": date,
                "Period From": pf,
                "Period To": pt,
                "Amount (£)": 1000.0,
                "Period Charge (£)": 100.0,
                "Cancel/Rebill Admitted": True,
                "Sub Periods": sub_periods,
            }
        ]
    )


def test_union_total_equals_sum_when_no_overlap() -> None:
    a = detect_back_billing(_df("A", "01 Mar 2023", "01 Jan 2020", "31 Dec 2020", ""))
    b = detect_back_billing(_df("B", "01 Mar 2024", "01 Jan 2023", "31 Dec 2023", ""))
    bb = pd.concat([a, b], ignore_index=True)
    total = compute_unlawful_union_total(bb)
    assert total == round(bb["Unlawful Charge (£)"].sum(), 2)


def test_union_total_does_not_double_count_overlap() -> None:
    # T67 and T68 overlap; the union must be <= the naive per-row sum and
    # strictly less when overlap exists.
    bb = pd.concat(
        [
            detect_back_billing(_df("T67", "13 Jul 2023", "15 Apr 2022", "03 Jul 2023", T67_SUB)),
            detect_back_billing(
                _df(
                    "T68",
                    "09 Aug 2023",
                    "02 Oct 2020",
                    "09 Aug 2023",
                    (
                        "02/10/2020|24/03/2021|19743.0|16.42|3241.8; "
                        "25/03/2021|06/04/2021|1454.0|16.42|238.75; "
                        "07/04/2021|31/03/2022|37184.0|16.42|6105.61; "
                        "01/04/2022|12/05/2022|3736.0|52.00|1942.72; "
                        "13/05/2022|31/03/2023|30675.0|52.00|15951.0; "
                        "01/04/2023|09/08/2023|10607.0|45.92|4870.73"
                    ),
                )
            ),
        ],
        ignore_index=True,
    )
    naive = round(bb["Unlawful Charge (£)"].sum(), 2)
    union = compute_unlawful_union_total(bb)
    assert union <= naive
    assert union < naive  # the overlapping days are counted once
