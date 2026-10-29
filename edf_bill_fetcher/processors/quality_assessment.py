"""Evidence-type eligibility and unresolved issues captured during analysis."""

from __future__ import annotations

import pandas as pd

from edf_bill_fetcher.helpers.date_utils import parse_to_sort_date
from edf_bill_fetcher.processors.preparation import identity, number, prepare_evidence

ASSESSMENT_FIELDS = frozenset(
    {
        "version",
        "financial_records",
        "bills",
        "dated_financial_records",
        "financial_amounts",
        "bill_charges",
        "transaction_records",
        "transaction_amounts",
        "complete_bill_periods",
        "classified_bill_readings",
        "bill_unit_rates",
        "duplicates_removed",
        "duplicates_retained",
        "unverified_amounts",
        "ambiguous_transactions",
        "unreconciled_statements",
    }
)


def build_quality_assessment(frame: pd.DataFrame, *, removed: int = 0) -> dict[str, int]:
    """Assess financial observations and bills without penalizing unrelated evidence."""
    records = frame.to_dict("records")
    financial = [
        row
        for row in records
        if row.get("Amount Kind") != "Not monetary"
        and row.get("Entry Type") != "Supporting Document"
    ]
    bills = [row for row in financial if row.get("Entry Type") == "New Bill"]
    transactions = [
        row
        for row in financial
        if row.get("Entry Type") != "New Bill"
        and (
            row.get("Entry Type") in {"Payment", "Credit", "Refund", "Late Payment", "Charge"}
            or row.get("Amount Kind") == "Transaction"
        )
    ]

    def valid_period(row: dict) -> bool:
        start = parse_to_sort_date(row.get("Period From"))
        end = parse_to_sort_date(row.get("Period To"))
        return not pd.isna(start) and not pd.isna(end) and start <= end

    return {
        "version": 1,
        "financial_records": len(financial),
        "bills": len(bills),
        "dated_financial_records": sum(
            not pd.isna(parse_to_sort_date(row.get("Date"))) for row in financial
        ),
        "financial_amounts": sum(number(row.get("Amount (£)")) is not None for row in financial),
        "bill_charges": sum(number(row.get("Period Charge (£)")) is not None for row in bills),
        "transaction_records": len(transactions),
        "transaction_amounts": sum(
            number(row.get("Transaction Amount (£)")) is not None for row in transactions
        ),
        "complete_bill_periods": sum(valid_period(row) for row in bills),
        "classified_bill_readings": sum(
            row.get("Reading")
            in {"Actual", "Smart", "Estimated", "Customer", "Deemed", "Deemed / estimated"}
            for row in bills
        ),
        "bill_unit_rates": sum((number(row.get("Unit Rate (p/kWh)")) or 0) > 0 for row in bills),
        "duplicates_removed": removed,
        "duplicates_retained": len(prepare_evidence(records, {"use_dedup": True}).duplicates),
        "unverified_amounts": sum(
            row.get("Amount Kind") == "Unknown" and number(row.get("Amount (£)")) is not None
            for row in financial
        ),
        "ambiguous_transactions": sum(
            bool(identity(row.get("Transaction Identity Review"))) for row in financial
        ),
        "unreconciled_statements": sum(
            row.get("Statement Reconciliation Status") == "Review required" for row in financial
        ),
    }


def validate_quality_assessment(value: dict[str, int], total: int, duplicates: int) -> None:
    """Validate stored scope/count relationships without examining original records."""
    if (
        not isinstance(value, dict)
        or set(value) != ASSESSMENT_FIELDS
        or any(type(count) is not int or count < 0 for count in value.values())
        or value["version"] != 1
    ):
        raise ValueError("Invalid quality assessment")
    financial, bills = value["financial_records"], value["bills"]
    if (
        not bills <= financial <= total
        or not value["transaction_amounts"] <= value["transaction_records"] <= financial
        or value["duplicates_retained"] > total
    ):
        raise ValueError("Invalid quality assessment eligibility")
    for key in (
        "dated_financial_records",
        "financial_amounts",
        "unverified_amounts",
        "ambiguous_transactions",
        "unreconciled_statements",
    ):
        if value[key] > financial:
            raise ValueError("Quality count exceeds eligible financial records")
    for key in (
        "bill_charges",
        "complete_bill_periods",
        "classified_bill_readings",
        "bill_unit_rates",
    ):
        if value[key] > bills:
            raise ValueError("Quality count exceeds eligible bills")
    if value["duplicates_removed"] + value["duplicates_retained"] != duplicates:
        raise ValueError("Quality duplicate counts disagree")
