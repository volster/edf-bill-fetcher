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


# ---------------------------------------------------------------------------
# Integration tests: _build_bb_clusters (matching.py) + handle_cluster_unmatched
# These exercise the production wiring path used in export.py, where clusters
# are built from detect_back_billing output and fed to the tagger.
# ---------------------------------------------------------------------------

from edf_bill_fetcher.processors.matching import _build_bb_clusters  # noqa: E402


def _make_back_billing_df() -> pd.DataFrame:
    """Build a detect_back_billing-shaped DataFrame with one cluster row."""
    return pd.DataFrame(
        [
            {
                "Invoice #": "KI-31105244-0014",
                "Bill Date": "2021-04-01",
                "Period From": pd.Timestamp("2021-03-01"),
                "Period To": pd.Timestamp("2021-03-31"),
                "Days Billed": 30,
                "Period Charge (£)": 150.0,
                "Value Source": "Period Charge",
                "12-Month Limit (days)": 365,
                "Excess Days": 0,
                "Unlawful Charge (£)": 0.0,
                "Cancel/Rebill Admitted": False,
                "Reason Assessment": "",
            }
        ]
    )


def test_build_bb_clusters_from_detect_back_billing_output() -> None:
    """_build_bb_clusters converts a back-billing row into a cluster dict
    with ISO date strings suitable for handle_cluster_unmatched's string
    comparison."""
    df = _make_back_billing_df()
    clusters = _build_bb_clusters(df)
    assert len(clusters) == 1
    c = clusters[0]
    assert c["name"] == "KI-31105244-0014"
    assert c["posting_date_start"] == "2021-03-01"
    assert c["posting_date_end"] == "2021-03-31"
    assert c["invoices"] == [{"Invoice #": "KI-31105244-0014", "Period Charge (£)": 150.0}]


def test_build_bb_clusters_empty_df() -> None:
    """Empty/None DataFrame → empty cluster list."""
    assert _build_bb_clusters(pd.DataFrame()) == []
    assert _build_bb_clusters(None) == []


def test_build_bb_clusters_skips_nat_dates() -> None:
    """Rows with NaT Period From/To are skipped."""
    df = pd.DataFrame(
        [
            {
                "Invoice #": "KI-1",
                "Period From": pd.NaT,
                "Period To": pd.Timestamp("2021-03-31"),
                "Period Charge (£)": 100.0,
            }
        ]
    )
    assert _build_bb_clusters(df) == []


def test_build_bb_clusters_skips_empty_invoice_id() -> None:
    """Rows with empty/N/A invoice id are skipped."""
    df = pd.DataFrame(
        [
            {
                "Invoice #": "",
                "Period From": pd.Timestamp("2021-03-01"),
                "Period To": pd.Timestamp("2021-03-31"),
                "Period Charge (£)": 100.0,
            },
            {
                "Invoice #": "N/A",
                "Period From": pd.Timestamp("2021-03-01"),
                "Period To": pd.Timestamp("2021-03-31"),
                "Period Charge (£)": 100.0,
            },
        ]
    )
    assert _build_bb_clusters(df) == []


def test_build_bb_clusters_unparseable_charge_falls_back_to_zero() -> None:
    """Rows with unparseable Period Charge fall back to 0.0 (not skipped)."""
    df = pd.DataFrame(
        [
            {
                "Invoice #": "KI-1",
                "Period From": pd.Timestamp("2021-03-01"),
                "Period To": pd.Timestamp("2021-03-31"),
                "Period Charge (£)": "not-a-number",
            }
        ]
    )
    clusters = _build_bb_clusters(df)
    assert len(clusters) == 1
    assert clusters[0]["invoices"][0]["Period Charge (£)"] == 0.0


def test_build_bb_clusters_nan_charge_falls_back_to_zero() -> None:
    """Rows with NaN Period Charge fall back to 0.0 (not skipped)."""
    df = pd.DataFrame(
        [
            {
                "Invoice #": "KI-1",
                "Period From": pd.Timestamp("2021-03-01"),
                "Period To": pd.Timestamp("2021-03-31"),
                "Period Charge (£)": float("nan"),
            }
        ]
    )
    clusters = _build_bb_clusters(df)
    assert len(clusters) == 1
    assert clusters[0]["invoices"][0]["Period Charge (£)"] == 0.0


def test_integration_cluster_unmatched_via_build_bb_clusters() -> None:
    """Full production path: detect_back_billing df → _build_bb_clusters →
    handle_cluster_unmatched → tagged event. An unmatched SAP event whose
    posting date falls inside the cluster window but whose amount disagrees
    with the cluster invoice gets tagged as 'internal mechanism'."""
    df = _make_back_billing_df()
    clusters = _build_bb_clusters(df)
    # SAP event: posting date inside 2021-03-01..03-31, amount 999 ≠ 150.
    sap_event = _make_event(
        posting_date_range=("2021-03-15", "2021-03-15"),
        net_amount=999.0,
    )
    tag = handle_cluster_unmatched(sap_event, clusters)
    assert tag is not None
    assert tag["Matched EDF Invoice #"] == "KI-31105244-0014 internal mechanism"
    assert tag["Confidence"] == 0


def test_integration_no_tag_when_amount_agrees() -> None:
    """When the SAP event's amount agrees with the cluster invoice, the
    tagger returns None — the event is NOT cluster-unmatched."""
    df = _make_back_billing_df()
    clusters = _build_bb_clusters(df)
    sap_event = _make_event(
        posting_date_range=("2021-03-15", "2021-03-15"),
        net_amount=150.0,
    )
    tag = handle_cluster_unmatched(sap_event, clusters)
    assert tag is None


def test_integration_no_tag_when_posting_date_outside_window() -> None:
    """When the SAP event's posting date is outside the cluster window,
    the tagger returns None."""
    df = _make_back_billing_df()
    clusters = _build_bb_clusters(df)
    sap_event = _make_event(
        posting_date_range=("2021-06-15", "2021-06-15"),
        net_amount=999.0,
    )
    tag = handle_cluster_unmatched(sap_event, clusters)
    assert tag is None
