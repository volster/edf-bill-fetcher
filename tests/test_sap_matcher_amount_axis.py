import pandas as pd

from edf_bill_fetcher.models.events import SapBackBillingEvent
from edf_bill_fetcher.processors.matching import match_sap_events_to_edf


def _build_event(
    clearing_doc: str = "023002707231",
    clearing_date: pd.Timestamp = pd.Timestamp("2021-04-01"),
    net_amount: float = 500.0,
) -> SapBackBillingEvent:
    return SapBackBillingEvent(
        clearing_doc=clearing_doc,
        clearing_date=clearing_date,
        clearing_reason="Statistical Item Reset",
        rows=[{"Posting Date": "2021-03-15", "Clearing Date": "2021-04-01", "Amount": "500.00"}],
        net_amount=net_amount,
        has_credit_for_consum_billing=False,
        has_account_maintenance=False,
        largest_single_posting=net_amount,
        posting_date_range=("2021-03-15", "2021-03-15"),
        evidence_trail="",
    )


def test_sap_matcher_uses_period_charge() -> None:
    # Period Charge (£) matches the SAP net amount; the old Amount (£) running
    # balance disagreed, so the old code would have mis-scored this event.
    sap_event = _build_event(net_amount=500.0)
    edf_records = [
        {
            "Invoice #": "T34",
            "Period From": "2020-01-01",
            "Period To": "2021-03-31",
            "Period Charge (£)": 500.0,  # matches SAP net amount
            "Amount (£)": 100.0,  # running balance differs; old code would mis-score here
        }
    ]
    result = match_sap_events_to_edf([sap_event], edf_records)
    assert len(result) == 1
    assert result[0].event.clearing_doc == "023002707231"
    assert result[0].edf_record["Invoice #"] == "T34"
    assert result[0].confidence_band == "High"  # score >= 75


def test_sap_matcher_falls_back_to_amount_when_period_charge_na() -> None:
    # When Period Charge (£) is N/A/unparseable, the matcher falls back to the
    # Amount (£) running balance — same value here, so the match is still High.
    sap_event = _build_event(net_amount=500.0)
    edf_records = [
        {
            "Invoice #": "T34",
            "Period From": "2020-01-01",
            "Period To": "2021-03-31",
            "Period Charge (£)": "N/A",
            "Amount (£)": 500.0,
        }
    ]
    result = match_sap_events_to_edf([sap_event], edf_records)
    assert len(result) == 1
    assert result[0].edf_record["Invoice #"] == "T34"
    assert result[0].confidence_band == "High"


def test_sap_matcher_falls_back_to_amount_when_period_charge_missing() -> None:
    # Records that predate the Period Charge column (or never populate it) must
    # still match on the Amount (£) running balance.
    sap_event = _build_event(net_amount=500.0)
    edf_records = [
        {
            "Invoice #": "T34",
            "Period From": "2020-01-01",
            "Period To": "2021-03-31",
            "Amount (£)": 500.0,
        }
    ]
    result = match_sap_events_to_edf([sap_event], edf_records)
    assert len(result) == 1
    assert result[0].edf_record["Invoice #"] == "T34"
    assert result[0].confidence_band == "High"
