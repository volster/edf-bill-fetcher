"""Explicit demands, quotes and planned/failed collections keep distinct monetary meanings."""

import pytest

from edf_bill_fetcher.collectors.engine import EvidenceEngine
from edf_bill_fetcher.processors.message_amounts import message_amount_context


@pytest.mark.parametrize(
    "text",
    [
        "We recently let you know that your bill for £1,234.56 is overdue.",
        "The total outstanding amount of £1,234.56 is now due in full.",
    ],
)
def test_explicit_demand_beats_fee_boilerplate_without_creating_another_charge(text):
    engine = EvidenceEngine({"filter_below": False}, lambda text: None)
    try:
        engine.process_text(
            text + " A late payment charge of £10 applies if your balance is unpaid.",
            "Email Body",
            "Your bill is overdue",
            "01/02/2025",
        )
        row = engine.records[0]
        assert row["Amount Kind"] == "Balance"
        assert row["Balance (£)"] == 1234.56
        assert row["Demand Amount (£)"] == 1234.56
        assert row["Transaction Amount (£)"] is None
        assert row["Period Charge (£)"] is None
        assert row["Document Role"] == "Payment reminder"
        assert row["Demand Context"]
    finally:
        engine.cleanup()


SUPPORTING = [
    (
        "Total annual cost £1,234.56. No exit fees.",
        "Your energy quote from EDF Energy",
        "Energy quote",
        "Quoted Annual Cost (£)",
    ),
    (
        "Your monthly Direct Debit payment amount is £1,234.56.",
        "Your energy bill is ready to view",
        "Direct Debit schedule",
        "Scheduled Direct Debit (£)",
    ),
    (
        "We'll collect your first electricity payment of £1,234.56 on 1 March.",
        "Your Direct Debit is now up and running",
        "Direct Debit schedule",
        "Scheduled first Direct Debit (£)",
    ),
    (
        "Your bank/building society hasn't paid us your latest Direct Debit of £1,234.56.",
        "An important message about your Direct Debit",
        "Failed Direct Debit notice",
        "Failed Direct Debit (£)",
    ),
]


@pytest.mark.parametrize(("text", "subject", "role", "field"), SUPPORTING)
def test_supporting_amount_is_preserved_without_entering_financial_totals(
    text, subject, role, field
):
    engine = EvidenceEngine({"filter_below": False}, lambda text: None)
    try:
        engine.process_text(text, "Email Body", subject, "01/02/2025")
        row = engine.records[0]
        assert row["Entry Type"] == "Supporting Document"
        assert row["Document Role"] == role
        assert row["Amount Kind"] == "Not monetary"
        assert row["Amount (£)"] == row[field] == 1234.56
        assert row["Date"] == "01/02/2025"
        assert row["Balance (£)"] is None
        assert row["Transaction Amount (£)"] is None
        assert row["Period Charge (£)"] is None
        assert row["Reading"] == "Unknown"
    finally:
        engine.cleanup()


@pytest.mark.parametrize(("text", "subject", "role", "field"), SUPPORTING)
@pytest.mark.parametrize(
    ("balance", "amount"),
    [
        ("Current balance £900 debit.", 900),
        ("Current balance £900 credit.", -900),
        ("Current balance is £900 in credit.", -900),
    ],
)
def test_a_specific_observed_balance_has_priority_over_supporting_context(
    text, subject, role, field, balance, amount
):
    engine = EvidenceEngine({"filter_below": False}, lambda text: None)
    try:
        engine.process_text(balance + " " + text, "Email Body", subject, "01/02/2025")
        assert engine.records[0]["Balance (£)"] == amount
        assert engine.records[0]["Amount Kind"] == "Balance"
    finally:
        engine.cleanup()


@pytest.mark.parametrize(
    "text",
    [
        "If your monthly Direct Debit payment amount is £1,234.56, contact us.",
        "If your bank hasn't paid us your latest Direct Debit of £1,234.56, contact us.",
        "For example, the total outstanding amount of £1,234.56 is now due in full.",
        "If your bill for £1,234.56 is overdue, contact us.",
    ],
)
def test_conditional_or_example_language_is_not_an_account_specific_fact(text):
    assert message_amount_context(text, "Advice") is None


@pytest.mark.parametrize(
    "text",
    [
        "We've collected your first electricity payment of £1,234.56.",
        "Your bank has paid us your latest Direct Debit of £1,234.56.",
        "Your next payment may be £1,234.56.",
        "Your payment is £1,234.56.",
        "Total annual cost £1,234.56.",
    ],
)
def test_nearby_unqualified_or_past_tense_language_is_left_for_other_parsers(text):
    assert message_amount_context(text, "Account information") is None


@pytest.mark.parametrize("value", ["12,34", "1234.567", "1,234.567"])
def test_malformed_amount_is_not_partially_consumed(value):
    assert (
        message_amount_context(
            "The total outstanding amount of £" + value + " is now due in full.", "Demand"
        )
        is None
    )


def test_curly_apostrophe_and_gas_schedule_are_supported():
    fact = message_amount_context("We’ll collect your first gas payment of £99.50.", "Schedule")
    assert fact is not None
    assert fact["Scheduled first Direct Debit (£)"] == 99.5


def test_account_filter_still_applies_to_supporting_messages():
    engine = EvidenceEngine(
        {"filter_below": False, "use_acc_filter": True, "acc_num": "811122233344"},
        lambda text: None,
    )
    try:
        engine.process_text(
            "Account number: 999988887777 Your monthly Direct Debit payment amount is £120.",
            "Email Body",
            "Your bill is ready",
            "01/02/2025",
        )
        assert engine.records == []
    finally:
        engine.cleanup()


def test_disabled_anchors_and_non_email_sources_do_not_gain_email_context():
    for source, anchors in [("Local PDF Folder", True), ("Email Body", False)]:
        engine = EvidenceEngine({"filter_below": False, "use_anchors": anchors}, lambda text: None)
        try:
            engine.process_text(
                "Your monthly Direct Debit payment amount is £120.",
                source,
                "Monthly notice",
                "01/02/2025",
            )
            assert engine.records[0]["Amount Kind"] == "Unknown"
        finally:
            engine.cleanup()


def test_unidentified_demand_stays_separate_and_gets_a_concrete_review_priority():
    from edf_bill_fetcher.models.output_overview import build_output_overview
    from edf_bill_fetcher.models.report_view import capture_report_view

    engine = EvidenceEngine({"filter_below": False}, lambda text: None)
    try:
        engine.process_text(
            "The total outstanding amount of £1,234.56 is now due in full.",
            "Email Body",
            "Instalment plan cancelled",
            "01/02/2025",
        )
        known = {
            "Account": "A-45678901",
            "Date": "01/03/2025",
            "Entry Type": "Ongoing Balance",
            "Amount (£)": 1500,
        }
        view = capture_report_view(
            [*engine.records, known],
            {"as_of": "2025-03-01"},
            formats=("html",),
        )
        assert view.snapshot.summary.closing_balance is None
        assert len(view.ledger_overview) == 2
        assigned = next(row for row in view.ledger_overview if row["references"])
        assert assigned["closing_balance"] == "1500.00"
        assert view.records[0]["Account"] == ""
        priorities = build_output_overview(view).priorities
        assert any(
            "Assign account identity" in item.title and item.sheet == "Review & Corrections"
            for item in priorities
        )
    finally:
        engine.cleanup()
