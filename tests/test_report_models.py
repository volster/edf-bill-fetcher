import pandas as pd

from edf_bill_fetcher.models.report_models import compute_payment_analysis

RECORD_KEYS = ["Date", "Entry Type", "Period Charge (£)", "Amount (£)", "Details"]


def test_empty_frame_yields_zeroed_analysis() -> None:
    analysis = compute_payment_analysis(pd.DataFrame(columns=RECORD_KEYS))

    assert analysis.count == 0
    assert analysis.total_paid == 0.0
    assert analysis.avg_payment == 0.0
    assert analysis.median_payment == 0.0
    assert analysis.largest_payment == 0.0
    assert analysis.smallest_payment == 0.0
    assert analysis.avg_interval_days is None
    assert analysis.median_interval_days is None
    assert analysis.last_payment_date is None
    assert analysis.last_payment_amount is None
    assert analysis.chronology.empty


def test_single_payment_uses_period_charge_and_no_intervals() -> None:
    df = pd.DataFrame(
        [
            {
                "Date": "01/01/2023",
                "Entry Type": "Payment",
                "Period Charge (£)": 900,
                "Amount (£)": 500,
                "Details": "customer payment",
            }
        ]
    )

    analysis = compute_payment_analysis(df)

    assert analysis.count == 1
    assert analysis.total_paid == 900.0
    assert analysis.largest_payment == 900.0
    assert analysis.smallest_payment == 900.0
    assert analysis.chronology["_amount"].iloc[0] == 900.0
    assert analysis.avg_interval_days is None
    assert analysis.median_interval_days is None


def test_two_payments_thirty_days_apart() -> None:
    df = pd.DataFrame(
        [
            {
                "Date": "31/01/2023",
                "Entry Type": "Payment",
                "Period Charge (£)": 100,
                "Amount (£)": 100,
                "Details": "",
            },
            {
                "Date": "01/01/2023",
                "Entry Type": "Payment",
                "Period Charge (£)": 100,
                "Amount (£)": 100,
                "Details": "",
            },
        ]
    )

    analysis = compute_payment_analysis(df)

    assert analysis.avg_interval_days == 30.0
    assert analysis.median_interval_days == 30.0


def test_amount_fallback_when_period_charge_is_na() -> None:
    df = pd.DataFrame(
        [
            {
                "Date": "01/01/2023",
                "Entry Type": "Payment",
                "Period Charge (£)": "N/A",
                "Amount (£)": 100,
                "Details": "",
            }
        ]
    )

    analysis = compute_payment_analysis(df)

    assert analysis.chronology["_amount"].iloc[0] == 100.0
    assert analysis.total_paid == 100.0


def test_credit_included_and_new_bill_excluded() -> None:
    df = pd.DataFrame(
        [
            {
                "Date": "01/01/2023",
                "Entry Type": "Payment",
                "Period Charge (£)": 200,
                "Amount (£)": 200,
                "Details": "",
            },
            {
                "Date": "02/01/2023",
                "Entry Type": "Credit",
                "Period Charge (£)": 50,
                "Amount (£)": 50,
                "Details": "",
            },
            {
                "Date": "03/01/2023",
                "Entry Type": "New Bill",
                "Period Charge (£)": 999,
                "Amount (£)": 999,
                "Details": "",
            },
        ]
    )

    analysis = compute_payment_analysis(df)

    assert analysis.count == 2
    assert analysis.total_paid == 250.0


def test_negative_amount_is_absoled_at_stat_level() -> None:
    df = pd.DataFrame(
        [
            {
                "Date": "01/01/2023",
                "Entry Type": "Payment",
                "Period Charge (£)": -500,
                "Amount (£)": -500,
                "Details": "",
            }
        ]
    )

    analysis = compute_payment_analysis(df)

    assert analysis.total_paid == 500.0
    assert analysis.chronology["_amount"].iloc[0] == -500.0


def test_last_payment_is_chronologically_last_row() -> None:
    df = pd.DataFrame(
        [
            {
                "Date": "05/01/2023",
                "Entry Type": "Payment",
                "Period Charge (£)": 100,
                "Amount (£)": 100,
                "Details": "",
            },
            {
                "Date": "20/01/2023",
                "Entry Type": "Payment",
                "Period Charge (£)": 250,
                "Amount (£)": 250,
                "Details": "",
            },
        ]
    )

    analysis = compute_payment_analysis(df)

    assert analysis.last_payment_date == "20/01/2023"
    assert analysis.last_payment_amount == 250.0
