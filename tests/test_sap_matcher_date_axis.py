import pandas as pd

from edf_bill_fetcher.models.events import SapBackBillingEvent
from edf_bill_fetcher.processors.matching import match_sap_events_to_edf


def _build_event(
    clearing_doc: str = "023002707231",
    clearing_date: pd.Timestamp = pd.Timestamp("2022-01-01"),
    net_amount: float = 100.0,
    rows: list[dict] | None = None,
) -> SapBackBillingEvent:
    """Build a minimal SapBackBillingEvent for testing."""
    if rows is None:
        rows = [
            {
                "Posting Date": "2021-03-15",
                "Clearing Date": "2022-01-01",
                "Amount": "100.00",
            }
        ]
    return SapBackBillingEvent(
        clearing_doc=clearing_doc,
        clearing_date=clearing_date,
        clearing_reason="Statistical Item Reset",
        rows=rows,
        net_amount=net_amount,
        has_credit_for_consum_billing=False,
        has_account_maintenance=False,
        largest_single_posting=net_amount,
        posting_date_range=("2021-03-15", "2021-03-15"),
        evidence_trail="",
    )


def test_sap_matcher_uses_posting_date() -> None:
    # Clearing Date is outside the EDF period, but the underlying rows'
    # Posting Date is inside the EDF period.
    sap_event = _build_event(
        clearing_date=pd.Timestamp("2022-01-01"),  # outside EDF period
        rows=[
            {
                "Posting Date": "2021-03-15",
                "Clearing Date": "2022-01-01",
                "Amount": "100.00",
            }
        ],
        net_amount=100.0,
    )
    edf_records = [
        {
            "Invoice #": "T34",
            "Period From": "2020-01-01",
            "Period To": "2021-03-31",
            "Period Charge (£)": 100.0,
            "Amount (£)": 100.0,
        }
    ]
    result = match_sap_events_to_edf([sap_event], edf_records)
    assert len(result) == 1
    assert result[0].event.clearing_doc == "023002707231"
    assert result[0].edf_record["Invoice #"] == "T34"
    assert result[0].confidence_band == "High"  # score >= 75
