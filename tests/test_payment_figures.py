import pandas as pd

from edf_bill_fetcher.helpers.payment_figures import payment_amount, payment_amounts


def test_period_charge_is_preferred_over_running_balance() -> None:
    row = pd.Series({"Period Charge (£)": "12.50", "Amount (£)": "900.00"})

    assert payment_amount(row) == (12.5, "Period Charge (£)")


def test_amount_is_used_when_period_charge_is_unavailable() -> None:
    rows = pd.DataFrame(
        [{"Period Charge (£)": "N/A", "Amount (£)": "900.00"}],
    )

    assert payment_amounts(rows).tolist() == [900.0]
