from __future__ import annotations

import pandas as pd

from edf_bill_fetcher.processors.matching import analyse_sap_back_billing


def _sap_row(doc, cd, reason, amount, txt) -> dict:
    return {
        "Document No.": doc,
        "Posting Date": "2023-07-13",
        "Amount": amount,
        "Transaction Text": txt,
        "Clearing Document": cd,
        "Clearing Date": "2023-08-01",
        "Clearing Reason": reason,
        "Clearing Status": "Cleared Item",
        "Statistical Key Flag": "",
    }


def _fixture() -> dict:
    sap = [
        # A real back-billing cluster: reversal credit + rebill debit.
        _sap_row("DOC-1", "CLR-100", "Reversal", -436.0, "Cr- Credit for Consum Billing"),
        _sap_row("DOC-2", "CLR-100", "Reversal", 436.0, "Dr- Consum Billing Receivable"),
        # An unrelated cluster (installment) — must be excluded.
        _sap_row("DOC-3", "CLR-999", "Automatic Clearing", 565.0, "Dr- Installment Receivable"),
    ]
    ev = pd.DataFrame(
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
    return {"sap": sap, "evidence": pd.DataFrame(), "bb": ev}


def test_sap_events_restricted_to_reversal_clusters() -> None:
    fx = _fixture()
    out = analyse_sap_back_billing(fx["sap"], fx["evidence"], fx["bb"])
    docs = {e["Clearing Doc #"] for e in out["events"]}
    assert docs == {"CLR-100"}  # CLR-999 excluded


def test_sap_bb_summary_totals() -> None:
    fx = _fixture()
    out = analyse_sap_back_billing(fx["sap"], fx["evidence"], fx["bb"])
    assert out["summary"]["sap_events"] == 1
    assert out["summary"]["sap_net_total"] == 0.0  # -436 + 436
