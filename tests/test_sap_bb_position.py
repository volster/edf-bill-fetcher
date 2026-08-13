from __future__ import annotations

import pandas as pd

from edf_bill_fetcher.models.events import SapBackBillingEvent
from edf_bill_fetcher.processors.matching import analyse_sap_back_billing


def _event(cd, rows, matched_edf_invoice=None) -> SapBackBillingEvent:
    net = round(sum(float(r["Amount"]) for r in rows), 2)
    has_credit = any("Credit for Consum Billing" in r["Transaction Text"] for r in rows)
    return SapBackBillingEvent(
        clearing_doc=cd,
        clearing_date=pd.Timestamp("2023-08-01"),
        clearing_reason="Reversal",
        rows=rows,
        net_amount=net,
        has_credit_for_consum_billing=has_credit,
        has_account_maintenance=False,
        largest_single_posting=net,
        posting_date_range=("2023-07-13", "2023-07-13"),
        matched_edf_invoice=matched_edf_invoice,
    )


def _fixture() -> dict:
    events = [
        # A real back-billing cluster: reversal credit + rebill debit, matched to T-001.
        _event(
            "CLR-100",
            [
                {"Document No.": "DOC-1", "Posting Date": "2023-07-13", "Amount": -436.0, "Transaction Text": "Cr- Credit for Consum Billing"},
                {"Document No.": "DOC-2", "Posting Date": "2023-07-13", "Amount": 436.0, "Transaction Text": "Dr- Consum Billing Receivable"},
            ],
            matched_edf_invoice="T-001",
        ),
        # A 2-row non-credit cluster (installment + interest) — must be excluded
        # by the credit filter, NOT by cluster size.
        _event(
            "CLR-999",
            [
                {"Document No.": "DOC-3", "Posting Date": "2023-07-13", "Amount": 565.0, "Transaction Text": "Dr- Installment Receivable"},
                {"Document No.": "DOC-4", "Posting Date": "2023-07-13", "Amount": 12.0, "Transaction Text": "Dr- Late Payment Charge"},
            ],
        ),
        # A credit-containing event with no matched invoice — must be SAP-only.
        _event(
            "CLR-200",
            [
                {"Document No.": "DOC-5", "Posting Date": "2023-07-13", "Amount": -100.0, "Transaction Text": "Cr- Credit for Consum Billing"},
                {"Document No.": "DOC-6", "Posting Date": "2023-07-13", "Amount": 100.0, "Transaction Text": "Dr- Consum Billing Receivable"},
            ],
        ),
    ]
    bb = pd.DataFrame(
        [
            {
                "Invoice #": "T-001",
                "Bill Date": "2023-07-13",
                "Period From": "2022-04-15",
                "Period To": "2023-07-03",
                "Period Charge (£)": 436.0,
                "Unlawful Charge (£)": 200.0,
                "_unlawful_slices": [],
            }
        ]
    )
    return {"events": events, "evidence": pd.DataFrame(), "bb": bb}


def test_sap_events_restricted_to_reversal_clusters() -> None:
    fx = _fixture()
    out = analyse_sap_back_billing(fx["events"], fx["evidence"], fx["bb"])
    docs = {e["Clearing Doc #"] for e in out["events"]}
    assert docs == {"CLR-100", "CLR-200"}  # CLR-999 excluded by credit filter


def test_sap_bb_summary_totals() -> None:
    fx = _fixture()
    out = analyse_sap_back_billing(fx["events"], fx["evidence"], fx["bb"])
    assert out["summary"]["sap_events"] == 2
    assert out["summary"]["sap_net_total"] == 0.0  # CLR-100 net 0 + CLR-200 net 0


def test_sap_bb_reconciliation_verdicts() -> None:
    fx = _fixture()
    out = analyse_sap_back_billing(fx["events"], fx["evidence"], fx["bb"])
    by_event = {r["SAP Event"]: r for r in out["reconciliation"]}
    # Matched + both amounts non-zero but differing -> Δ row.
    # CLR-100 has a single -436 credit-for-consum-billing row -> reversal
    # magnitude 436.0, compared against our 200.0 -> Δ.
    assert by_event["CLR-100"]["EDF Invoice #"] == "T-001"
    assert by_event["CLR-100"]["EDF Unlawful Charge (£)"] == 200.0
    assert by_event["CLR-100"]["SAP Net (£)"] == 436.0
    assert "Δ £" in by_event["CLR-100"]["Verdict"]
    # Unmatched credit event -> SAP-only (money with no matching invoice).
    assert by_event["CLR-200"]["EDF Invoice #"] == "—"
    assert by_event["CLR-200"]["Verdict"] == "SAP-only"


def test_sap_bb_reconciliation_reversal_magnitude_reconciles() -> None:
    fx = _fixture()
    # Match our unlawful charge to the reversal magnitude exactly (436.0).
    # CLR-100 nets to ~£0 (credit + rebill), so this only reconciles
    # because the verdict compares against the credit magnitude, not net.
    fx["bb"].loc[0, "Unlawful Charge (£)"] = 436.0
    out = analyse_sap_back_billing(fx["events"], fx["evidence"], fx["bb"])
    by_event = {r["SAP Event"]: r for r in out["reconciliation"]}
    assert by_event["CLR-100"]["Verdict"] == "Reconciled"
    assert by_event["CLR-100"]["SAP Net (£)"] == 436.0
    assert out["summary"]["reconciled"] == 1
