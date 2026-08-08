"""Unit tests for ``handle_cluster_unmatched`` (Task 8 — Decision 4).

When a SAP back-billing event's Posting Date falls inside a known
back-billing cluster's posting-date window but no invoice in that
cluster achieves amount-band agreement, the handler tags the event as
an internal mechanism of that cluster.
"""

from __future__ import annotations

import pandas as pd

from edf_bill_fetcher.models.events import SapBackBillingEvent
from edf_bill_fetcher.writers._helpers import handle_cluster_unmatched


def _make_event(
    posting_date_range: tuple[str, str] = ("2021-03-15", "2021-03-15"),
    net_amount: float = 999.0,
    clearing_doc: str = "023002707231",
) -> SapBackBillingEvent:
    """Build a minimal SapBackBillingEvent for tests."""
    return SapBackBillingEvent(
        clearing_doc=clearing_doc,
        clearing_date=pd.Timestamp("2021-04-01"),
        clearing_reason="Statistical Item Reset",
        rows=[{"Posting Date": posting_date_range[0], "Amount": str(net_amount)}],
        net_amount=net_amount,
        has_credit_for_consum_billing=False,
        has_account_maintenance=False,
        largest_single_posting=net_amount,
        posting_date_range=posting_date_range,
        evidence_trail="",
    )


def test_cluster_unmatched_tag() -> None:
    """Posting Date inside cluster window + no amount agreement → tagged."""
    sap_event = _make_event(
        posting_date_range=("2021-03-15", "2021-03-15"),
        net_amount=999.0,
    )
    clusters = [
        {
            "name": "T33/T34",
            "posting_date_start": "2021-03-01",
            "posting_date_end": "2021-03-31",
            "invoices": [
                {"Invoice #": "T33", "Period Charge (£)": 100.0},
                {"Invoice #": "T34", "Period Charge (£)": 200.0},
            ],
        }
    ]
    result = handle_cluster_unmatched(sap_event, clusters)
    assert result is not None
    assert result["Matched EDF Invoice #"] == "T33/T34 internal mechanism"
    assert "Posting Date inside cluster window" in result["Notes"]


def test_cluster_unmatched_no_window_overlap_returns_none() -> None:
    """Posting Date outside every cluster window → None."""
    sap_event = _make_event(
        posting_date_range=("2021-06-15", "2021-06-15"),
        net_amount=999.0,
    )
    clusters = [
        {
            "name": "T33/T34",
            "posting_date_start": "2021-03-01",
            "posting_date_end": "2021-03-31",
            "invoices": [
                {"Invoice #": "T33", "Period Charge (£)": 100.0},
            ],
        }
    ]
    result = handle_cluster_unmatched(sap_event, clusters)
    assert result is None


def test_cluster_unmatched_amount_agreement_returns_none() -> None:
    """When an in-cluster invoice matches on amount band, do NOT tag."""
    # SAP net 100.0, invoice 100.0 → within 5% → amount agreement.
    sap_event = _make_event(
        posting_date_range=("2021-03-15", "2021-03-15"),
        net_amount=100.0,
    )
    clusters = [
        {
            "name": "T33/T34",
            "posting_date_start": "2021-03-01",
            "posting_date_end": "2021-03-31",
            "invoices": [
                {"Invoice #": "T33", "Period Charge (£)": 100.0},
            ],
        }
    ]
    result = handle_cluster_unmatched(sap_event, clusters)
    assert result is None


def test_cluster_unmatched_empty_posting_date_range_returns_none() -> None:
    """Empty posting_date_range → no date to compare → None."""
    sap_event = _make_event(
        posting_date_range=("", ""),
        net_amount=999.0,
    )
    clusters = [
        {
            "name": "T33/T34",
            "posting_date_start": "2021-03-01",
            "posting_date_end": "2021-03-31",
            "invoices": [
                {"Invoice #": "T33", "Period Charge (£)": 100.0},
            ],
        }
    ]
    result = handle_cluster_unmatched(sap_event, clusters)
    assert result is None


def test_cluster_unmatched_empty_clusters_returns_none() -> None:
    """No clusters → None."""
    sap_event = _make_event()
    result = handle_cluster_unmatched(sap_event, [])
    assert result is None
