"""SAP reconciliation distinguishes unmatched money from an agreed zero comparison."""

import pandas as pd
import pytest

from edf_bill_fetcher.models.events import SapBackBillingEvent
from edf_bill_fetcher.processors.matching import _build_bb_clusters, analyse_sap_back_billing


@pytest.mark.parametrize(
    ("reversal", "challenged", "verdict"),
    [(0, 0, "Reconciled"), (100, 0, "SAP-only"), (0, 100, "Ours-only"), (100, 100, "Reconciled")],
)
def test_reconciliation_uses_credit_magnitude_and_preserves_the_source_period(
    reversal, challenged, verdict
):
    event = SapBackBillingEvent(
        "CLEAR-1",
        pd.Timestamp("2024-02-05"),
        "Cancel and rebill",
        rows=[
            {"Transaction Text": "Cr- Credit for Consum Billing", "Amount": -reversal},
            {"Transaction Text": "Consumption Billing", "Amount": reversal},
        ],
        net_amount=0,
        has_credit_for_consum_billing=True,
        matched_edf_invoice="INV-1",
    )
    evidence = pd.DataFrame(
        [
            {
                "Source": "Statement Reconciliation",
                "Invoice #": "INV-1",
                "Details": "Reversal (1 January 2024 - 31 January 2024)",
            },
            {
                "Source": "Statement Reconciliation",
                "Invoice #": "OTHER",
                "Details": "No period in this source line",
            },
        ]
    )
    bills = pd.DataFrame([{"Invoice #": "INV-1", "Unlawful Charge (£)": challenged}])
    result = analyse_sap_back_billing([event], evidence, bills)
    assert result["reconciliation"][0]["Verdict"] == verdict
    assert result["reconciliation"][0]["SAP Net (£)"] == reversal
    assert result["events"][0]["Net Amount (£)"] == 0
    assert result["events"][0]["Period(s)"] == "1 January 2024 - 31 January 2024"


@pytest.mark.parametrize("invalid", ["unreadable", "2024-02-30"])
def test_invalid_period_does_not_create_a_backbilling_cluster(invalid):
    rows = pd.DataFrame(
        [
            {
                "Invoice #": "BAD",
                "Period From": invalid,
                "Period To": "2024-01-31",
                "Period Charge (£)": 100,
            },
            {
                "Invoice #": "GOOD",
                "Period From": "2024-01-01",
                "Period To": "2024-01-31",
                "Period Charge (£)": 100,
            },
        ]
    )
    clusters = _build_bb_clusters(rows)
    assert len(clusters) == 1
    assert clusters[0]["name"] == "GOOD"
    assert clusters[0]["posting_date_start"] == "2024-01-01"
