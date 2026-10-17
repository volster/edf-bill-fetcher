"""Reader-facing facts and review priorities from an accepted analysis."""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from typing import TYPE_CHECKING

from edf_bill_fetcher.domain.json_values import object_value
from edf_bill_fetcher.helpers.date_utils import parse_to_sort_date
from edf_bill_fetcher.helpers.formatting import fmt_money
from edf_bill_fetcher.helpers.identity import account_key
from edf_bill_fetcher.processors.preparation import identity, number, scope

if TYPE_CHECKING:
    from edf_bill_fetcher.models.report_view import ReportView


@dataclass(frozen=True)
class ReviewPriority:
    """A concrete follow-up and the workbook/report location that explains it."""

    title: str
    action: str
    sheet: str
    section: str


@dataclass(frozen=True)
class OutputOverview:
    """Presentation only: no fitting, reconciliation or financial recalculation."""

    analysis_date: str
    coverage: str
    record_count: int
    source_count: int
    ledger_count: int
    metrics: tuple[tuple[str, str], ...]
    priorities: tuple[ReviewPriority, ...]
    balance_note: str
    exclusions: int
    duplicates: int
    notes: int


def report_account_label(view: ReportView) -> str:
    """Use a supplied display reference, or describe the actual included accounts."""
    config = view.config
    if config.get("report_account_ref") or config.get("acc_num"):
        return str(config.get("report_account_ref") or config["acc_num"])
    accounts = {account_key(identity(row.get("Account"))) for row in view.records}
    known = sorted(account for account in accounts if account)
    if len(known) == 1:
        return known[0] + (" + unassigned records" if "" in accounts else "")
    if len(known) > 1:
        return f"Multiple accounts ({len(known)})"
    return "Not supplied"


def build_output_overview(view: ReportView) -> OutputOverview:
    """Summarize supplied evidence without certifying accuracy or completeness."""
    import pandas as pd

    records = view.records
    state = object_value(view.snapshot.prepared.extraction.legacy_state)
    dates = [parse_to_sort_date(row.get("Date")) for row in records]
    dated = [value for value in dates if not pd.isna(value)]
    coverage = (
        f"{min(dated):%d %b %Y} – {max(dated):%d %b %Y}" if dated else "No usable record dates"
    )
    bills = [row for row in records if row.get("Entry Type") in {"New Bill", "Charge"}]
    ledger_count = len({scope(row) for row in records})
    manifest = state.get("source_manifest", {})
    source_count = len(manifest) if isinstance(manifest, dict) else 0
    exclusions = len(state.get("review_excluded_records", []))
    notes = len(state.get("error_log", []))
    priorities = []
    historical = sum(
        row.get("Document Role") == "Historical account statement"
        and row.get("Entry Type") == "Charge"
        and not identity(row.get("First Billed Date"))
        for row in records
    )
    unbalanced = sum(
        row.get("Statement Reconciliation Status") == "Review required" for row in records
    )
    if historical or unbalanced:
        priorities.append(
            ReviewPriority(
                "Reconcile historical statement evidence",
                f"{historical} statement charge observations have no established first-billed date; {unbalanced} statement equations do not balance. Historical observations may repeat other sources. Review these relationships before relying on aggregate movement; undated historical charges are excluded from back-billing date assessments.",
                "Review & Corrections",
                "data_quality",
            )
        )
    needs_info = sum(
        item.get("status") == "needs_information"
        for item in state.get("human_review_statuses", {}).values()
    )
    if needs_info:
        priorities.append(
            ReviewPriority(
                f"Resolve {needs_info} human review information request{'s' if needs_info != 1 else ''}",
                "Inspect the record review states and reasons in the app and Review history before sharing the pack.",
                "Review history",
                "detailed_findings",
            )
        )
    counts = view.model("dispute").counts
    high, medium = int(counts.get("HIGH", 0)), int(counts.get("MEDIUM", 0))
    if high or medium:
        priorities.append(
            ReviewPriority(
                f"{high} high / {medium} medium priority findings",
                "Read the dated findings, then compare each observation with its source and account context.",
                "Dispute Flags",
                "key_findings",
            )
        )
    ocr = sum(bool(row.get("OCR Review Required")) for row in records)
    if ocr or view.ocr_notice:
        priorities.append(
            ReviewPriority(
                f"Check OCR-derived evidence ({ocr} included record{'s' if ocr != 1 else ''})",
                "Verify dates, invoice references and money against the scanned pages. Confidence is not proof of accuracy.",
                "Review & Corrections",
                "data_quality",
            )
        )
    missing_refs = sum(not identity(row.get("Invoice #")) for row in bills)
    missing_periods = sum(
        pd.isna(parse_to_sort_date(row.get("Period From")))
        or pd.isna(parse_to_sort_date(row.get("Period To")))
        or parse_to_sort_date(row.get("Period From")) > parse_to_sort_date(row.get("Period To"))
        for row in bills
    )
    missing_charges = sum(number(row.get("Period Charge (£)")) is None for row in bills)
    if missing_refs or missing_periods or missing_charges:
        priorities.append(
            ReviewPriority(
                "Complete the bill references and amounts",
                f"Included bill rows missing an invoice reference: {missing_refs}; missing/invalid period: {missing_periods}; missing charge: {missing_charges}. Check and correct these fields in the app or Excel.",
                "Review & Corrections",
                "data_quality",
            )
        )
    if notes:
        priorities.append(
            ReviewPriority(
                f"Read {notes} extraction / review note{'s' if notes != 1 else ''}",
                "Check Parse Errors for unreadable messages, missing files and other limitations before relying on the report.",
                "Parse Errors",
                "evidence_index",
            )
        )
    payments = sum(row.get("Entry Type") in {"Payment", "Credit", "Refund"} for row in records)
    if not payments:
        priorities.append(
            ReviewPriority(
                "No payment, credit or refund rows supplied",
                "Locate the payment history before treating the recorded movement as a complete account reconciliation.",
                "Payment Analysis",
                "payment",
            )
        )
    if not priorities:
        priorities.append(
            ReviewPriority(
                "Compare the summary with the originals",
                "No priority gaps were found by these checks. Verify source coverage and the final complaint wording; automated checks cannot establish that every bill or payment is present.",
                "EDF Evidence Report",
                "evidence_index",
            )
        )
    summary = view.snapshot.summary

    def money(value: Decimal | None) -> str:
        return fmt_money(value) if value is not None else "Unavailable"

    balance_note = (
        "First and last balances are observations in the supplied records, not a certified opening or current balance. Recorded net movement is charges less net payments/credits; it need not explain the difference between those balances."
        if ledger_count <= 1
        else f"These records span {ledger_count} account/fuel/meter/register ledgers. First/last balances are unavailable across mixed ledgers; review each ledger separately. Recorded movement is an aggregate of supplied transactions."
    )
    return OutputOverview(
        view.snapshot.prepared.extraction.request.as_of.isoformat(),
        coverage,
        len(records),
        source_count,
        ledger_count,
        (
            ("Recorded charges", money(summary.charges)),
            ("Net payments / credits", money(summary.net_credits)),
            ("Recorded net movement", money(summary.charges - summary.net_credits)),
            ("First observed balance", money(summary.opening_balance)),
            ("Last observed balance", money(summary.closing_balance)),
        ),
        tuple(priorities),
        balance_note,
        exclusions,
        len(view.duplicates),
        notes,
    )
