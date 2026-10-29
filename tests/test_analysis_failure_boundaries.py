"""Reject stale input bindings and keep analytical descriptions within the available evidence."""

import copy

import pandas as pd
import pytest

from edf_bill_fetcher.application import analysis_service
from edf_bill_fetcher.application.analysis_service import _reconciliation, _timeline
from edf_bill_fetcher.application.legacy_adapter import capture_legacy
from edf_bill_fetcher.application.preparation_service import prepare_snapshot
from edf_bill_fetcher.models.report_view import RenderSelection, build_report_view
from tests.synthetic_analysis import snapshot
from tests.test_financial_identity_boundaries import bill
from tests.test_rebilling_ledger import position


@pytest.mark.parametrize(
    ("state", "message"),
    [
        ({"dispute_schedule": [position("missing")]}, "Dispute position refers"),
        (
            {
                "claim_reconciliation": {
                    "groups": [{"id": "C1", "members": ["missing"], "credits": []}]
                }
            },
            "Claim group refers",
        ),
        (
            {"human_review_statuses": {"missing": {"status": "verified", "reason": "Checked"}}},
            "Review status refers",
        ),
    ],
)
def test_unknown_review_references_are_not_transferred_to_another_record(state, message):
    original = snapshot([bill()])
    before = copy.deepcopy(state)
    captured = capture_legacy([bill()], state, original.prepared.extraction.request)
    with pytest.raises(ValueError, match=message):
        prepare_snapshot(captured)
    assert state == before


@pytest.mark.parametrize("last_reading", ["Actual", "Estimated"])
def test_timeline_estimated_run_reports_observed_readings_without_claiming_hidden_bills(
    last_reading,
):
    records = [
        bill(
            **{
                "Date": f"0{n}/02/2024",
                "Invoice #": f"INV-{n}",
                "Balance (£)": 100 + n,
                "Reading": "Estimated" if n < 4 else last_reading,
            }
        )
        for n in range(1, 5)
    ]
    timeline = build_report_view(snapshot(records), RenderSelection()).model("sap")["timeline"]
    estimates = [row for row in timeline if row[1] == "ESTIMATED READINGS"]
    assert len(estimates) == 1
    assert estimates[0][0] == "01/02/2024"
    assert ("3 consecutive" if last_reading == "Actual" else "4 consecutive") in estimates[0][2]


def test_legacy_timeline_without_reading_column_still_describes_observed_balances():
    frame = pd.DataFrame(
        [
            {"Date": "01/01/2024", "_dt": pd.Timestamp("2024-01-01"), "Amount (£)": 100},
            {"Date": "02/01/2024", "_dt": pd.Timestamp("2024-01-02"), "Amount (£)": 200},
        ]
    )
    timeline = _timeline(frame, [])
    assert timeline[0][1] == "FIRST OBSERVED BALANCE"
    assert not any(row[1] == "ESTIMATED READINGS" for row in timeline)


def test_invalid_financial_result_is_rejected_before_publication(monkeypatch):
    accepted = snapshot([bill()])
    monkeypatch.setattr(analysis_service, "financial_totals", lambda frame: (float("nan"), 0))
    with pytest.raises(ValueError, match="unavailable total"):
        analysis_service.analyse_prepared(accepted.prepared)
    assert str(accepted.summary.charges) == "100.00"


def test_reconciliation_matches_contract_codes_and_signed_payments_one_to_one():
    state = {
        "sap_contract_rows": [
            {"Contract From": "01/01/2024", "Product Code": "CODE", "Account": "A-EXAMPLE"}
        ],
        "sap_financial_rows": [
            {
                "Posting Date": "01/02/2024",
                "Amount": 25,
                "Document Type": "Payment",
                "Account": "A-EXAMPLE",
            }
        ],
    }
    frame = pd.DataFrame(
        [bill(**{"Date": "01/02/2024", "Entry Type": "Payment", "Transaction Amount (£)": -25})]
    )
    contracts = pd.DataFrame(
        [{"Contract From": "01/01/2024", "Product Code": "CODE", "Account": "A-EXAMPLE"}]
    )
    model = _reconciliation(state, frame, contracts)
    summaries = {row["values"][0]: row["values"] for row in model["summaries"]}
    assert summaries["Contract"][3] == 1
    assert summaries["Contract"][-1] == "MATCHED"
    assert summaries["Financial"][3] == 1
    assert summaries["Financial"][-1] == "MATCHED"
