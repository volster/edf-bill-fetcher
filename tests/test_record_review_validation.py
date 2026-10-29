"""Shared Excel/desktop correction boundaries preserve original evidence."""

import copy
from datetime import date, datetime
from decimal import Decimal

import pytest

from edf_bill_fetcher.application.record_review import (
    apply_review,
    corrected_record,
    input_value,
    review_row_data,
    validate_review_statuses,
)


def context():
    return {
        "version": 1,
        "base_fingerprint": "a" * 64,
        "rows": [
            {
                "id": "b" * 64,
                "included": True,
                "record": {
                    "Entry Type": "New Bill",
                    "Amount Kind": "Transaction",
                    "Transaction Amount (£)": 100,
                    "Period Charge (£)": 100,
                    "Source ID": "original",
                    "provenance": {"pages": [1, 2]},
                },
            }
        ],
        "decisions": [{"row_id": "b" * 64, "include": True, "changes": {}, "reason": ""}],
    }


@pytest.mark.parametrize(
    ("field", "value", "expected"),
    [
        ("Date", date(2024, 2, 29), "29/02/2024"),
        ("Date", datetime(2024, 2, 29), "29/02/2024"),
        ("Date", "2024-02-29", "29/02/2024"),
        ("Date", "29/02/2024", "29/02/2024"),
        ("Date", "", None),
        ("Cancel/Rebill Admitted", True, True),
        ("Cancel/Rebill Admitted", False, False),
        ("Cancel/Rebill Admitted", "Yes", True),
        ("Cancel/Rebill Admitted", "No", False),
        ("Cancel/Rebill Admitted", None, False),
        ("Balance (£)", Decimal("-12.34"), -12.34),
        ("Balance (£)", 0, 0.0),
        ("Balance (£)", "1000000000000", 1e12),
        ("Balance (£)", "", None),
        ("Units (kWh)", "0.123456", 0.123456),
        ("Account", "000123", "000123"),
        ("Invoice #", None, None),
        ("Details", "=literal text", "=literal text"),
        ("Amount Kind", "Unknown", "Unknown"),
    ],
)
def test_supported_correction_values_are_lossless(field, value, expected):
    assert input_value(field, value) == expected


@pytest.mark.parametrize(
    ("field", "value", "message"),
    [
        ("Entry Type", None, "choose a value"),
        ("Amount Kind", "", "choose a value"),
        ("Cancel/Rebill Admitted", 1, "Yes or No"),
        ("Cancel/Rebill Admitted", "true", "Yes or No"),
        ("Date", datetime(2024, 2, 29, 0, 0, 1), "without a time"),
        ("Date", "29/02/2023", "real Excel date"),
        ("Date", "next Monday", "real Excel date"),
        ("Date", 45351, "real Excel date"),
        ("Balance (£)", True, "finite number"),
        ("Balance (£)", [], "finite number"),
        ("Balance (£)", "NaN", "finite number"),
        ("Balance (£)", float("inf"), "finite number"),
        ("Balance (£)", "-1000000000000.01", "finite number"),
        ("Balance (£)", "£12.34", "finite number"),
        ("Balance (£)", "1,234.56", "finite number"),
        ("Balance (£)", "12.345", "2 decimal places"),
        ("Units (kWh)", "0.1234567", "6 decimal places"),
        ("Account", 123, "format identifiers as Text"),
        ("Details", "x" * 32001, "enter text"),
        ("Amount Kind", "Payment", "Balance, Transaction or Unknown"),
        ("Entry Type", "Invented", "supported type"),
    ],
)
def test_invalid_correction_is_rejected_instead_of_guessed(field, value, message):
    with pytest.raises(ValueError, match=message):
        input_value(field, value)


@pytest.mark.parametrize(
    ("entry_type", "old_charge", "old_transaction", "new_charge", "expected"),
    [
        ("New Bill", 100, 100, 125, 125),
        ("New Bill", 100, 100, None, None),
        ("New Bill", None, None, 125, 125),
        ("New Bill", 100, 20, 125, 20),
        ("New Bill", 100, 100, -25, -25),
        ("Payment", 100, -100, 125, -125),
        ("Credit", None, None, 125, -125),
        ("Refund", 100, 100, 125, 125),
    ],
)
def test_charge_correction_updates_only_linked_transaction(
    entry_type, old_charge, old_transaction, new_charge, expected
):
    original = {
        "Entry Type": entry_type,
        "Amount Kind": "Transaction",
        "Period Charge (£)": old_charge,
        "Transaction Amount (£)": old_transaction,
        "provenance": {"pages": [1]},
    }
    before = copy.deepcopy(original)
    corrected = corrected_record(original, {"Period Charge (£)": new_charge})
    assert corrected["Transaction Amount (£)"] == expected
    assert corrected["Amount (£)"] == expected
    assert corrected["Review Explicit Amounts"] is True
    corrected["provenance"]["pages"].append(2)
    assert original == before


@pytest.mark.parametrize(("entry_type", "amount"), [("Payment", 1), ("Credit", 1), ("Refund", -1)])
def test_invalid_transaction_sign_never_changes_original(entry_type, amount):
    original = {"Entry Type": entry_type, "Transaction Amount (£)": 0}
    with pytest.raises(ValueError, match="negative transactions"):
        corrected_record(original, {"Transaction Amount (£)": amount})
    assert original["Transaction Amount (£)"] == 0


def test_explicit_transaction_and_unknown_meaning_are_preserved():
    original = context()["rows"][0]["record"]
    corrected = corrected_record(
        original,
        {"Period Charge (£)": 120, "Transaction Amount (£)": 25, "Amount Kind": "Unknown"},
    )
    assert corrected["Transaction Amount (£)"] == 25
    assert corrected["Period Charge (£)"] == 120
    assert corrected["Amount (£)"] is None
    assert original["Period Charge (£)"] == 100


@pytest.mark.parametrize("rate", [12.345678, None])
def test_rate_correction_records_basis_and_explicit_clear(rate):
    original = {"Energy Unit Rate (p/kWh)": 10, "Unit Rate (p/kWh)": 10}
    corrected = corrected_record(original, {"Energy Unit Rate (p/kWh)": rate})
    assert corrected["Unit Rate (p/kWh)"] == rate
    assert corrected["Unit Rate Basis"] == ("Human correction" if rate is not None else "Unknown")
    assert original["Unit Rate (p/kWh)"] == 10


def test_malformed_existing_period_can_be_repaired_one_endpoint_at_a_time():
    original = {"Period From": "unreadable", "Period To": "also unreadable"}
    first = corrected_record(original, {"Period From": "01/01/2024"})
    assert first["Period To"] == "also unreadable"
    second = corrected_record(first, {"Period To": "31/01/2024"})
    assert second["Period From"] == "01/01/2024"
    with pytest.raises(ValueError, match="must not be later"):
        corrected_record(second, {"Period From": "01/02/2024"})
    assert second["Period From"] == "01/01/2024"
    assert original["Period From"] == "unreadable"


@pytest.mark.parametrize(
    ("target", "field", "value", "message"),
    [
        ("context", "version", True, "human review context"),
        ("context", "version", 2, "human review context"),
        ("context", "base_fingerprint", "bad", "base fingerprint"),
        ("context", "rows", {}, "must be arrays"),
        ("context", "decisions", {}, "must be arrays"),
        ("row", "id", "bad", "review row identity"),
        ("row", "included", 1, "review row identity"),
        ("row", "record", [], "review row identity"),
        ("decision", "row_id", "c" * 64, "review row identity"),
        ("decision", "include", 1, "correction fields"),
        ("decision", "changes", [], "correction fields"),
        ("decision", "changes", {"Source ID": "forged"}, "correction fields"),
        ("decision", "reason", None, "correction fields"),
        ("decision", "reason", "x" * 4001, "correction fields"),
        ("decision", "include", False, "reason is required"),
        ("decision", "changes", {"Details": "amended"}, "reason is required"),
    ],
)
def test_invalid_review_context_is_rejected_without_mutation(target, field, value, message):
    review = context()
    destination = {"context": review, "row": review["rows"][0], "decision": review["decisions"][0]}[
        target
    ]
    destination[field] = value
    before = copy.deepcopy(review)
    with pytest.raises(ValueError, match=message):
        apply_review(review)
    assert review == before


@pytest.mark.parametrize("duplicate", ["rows", "decisions"])
def test_repeated_review_identity_is_rejected(duplicate):
    review = context()
    review[duplicate].append(copy.deepcopy(review[duplicate][0]))
    with pytest.raises(ValueError, match="duplicate review row identity"):
        apply_review(review)


def test_incomplete_decisions_do_not_partially_apply_edits():
    review = context()
    review["rows"].append({"id": "c" * 64, "included": True, "record": {}})
    review["decisions"][0].update(changes={"Period Charge (£)": 125}, reason="Verified source")
    before = copy.deepcopy(review)
    with pytest.raises(ValueError, match="use Include = No"):
        apply_review(review)
    assert review == before


def test_exclusion_preserves_corrected_record_and_original_provenance():
    review = context()
    review["decisions"][0].update(
        include=False, changes={"Period Charge (£)": "125.00"}, reason="Unrelated account"
    )
    before = copy.deepcopy(review)
    included, excluded = apply_review(review)
    assert included == []
    assert len(excluded) == 1
    assert excluded[0]["Period Charge (£)"] == excluded[0]["Transaction Amount (£)"] == 125
    assert excluded[0]["Review Row ID"] == "b" * 64
    assert excluded[0]["Source ID"] == "original"
    excluded[0]["provenance"]["pages"].append(3)
    assert review == before


def test_stable_review_identity_survives_exclusion_and_detects_duplicate_bindings():
    original = {"Invoice #": "INV-1"}
    first = review_row_data([original], [], "a" * 64)[0]
    preserved = {**original, "Review Row ID": first["id"]}
    excluded = review_row_data([], [preserved], "a" * 64)[0]
    assert excluded["id"] == first["id"]
    assert excluded["included"] is False
    assert original == {"Invoice #": "INV-1"}
    with pytest.raises(ValueError, match="identities are ambiguous"):
        review_row_data([preserved], [preserved], "a" * 64)


@pytest.mark.parametrize(
    "statuses",
    [
        [],
        {"unknown": {"status": "verified", "reason": "Checked"}},
        {"known": None},
        {"known": {"status": "verified"}},
        {"known": {"status": "invented", "reason": "Checked"}},
        {"known": {"status": "verified", "reason": 7}},
        {"known": {"status": "verified", "reason": "x" * 4001}},
        {"known": {"status": "verified", "reason": "  "}},
        {"known": {"status": "needs_information", "reason": ""}},
    ],
)
def test_invalid_verification_annotation_never_changes_state(statuses):
    before = copy.deepcopy(statuses)
    with pytest.raises(ValueError):
        validate_review_statuses({"known"}, statuses)
    assert statuses == before


@pytest.mark.parametrize("status", ["unreviewed", "verified", "needs_information"])
def test_verification_is_independent_of_financial_inclusion(status):
    statuses = {"known": {"status": status, "reason": "Compared with original"}}
    before = copy.deepcopy(statuses)
    validate_review_statuses({"known"}, statuses)
    assert statuses == before
