"""Shared presentation of accepted data-quality facts across all output formats."""

from __future__ import annotations

from edf_bill_fetcher.models.report_models import DataQualityReport

NOTICE = (
    "Dates and numeric amounts are checked on financial observations; bill charges and transaction "
    "values are checked separately. Periods, reading basis and energy "
    "rates are checked on bills; reminders, payments, credits and supporting documents do not need "
    "their own bill period or reading. REVIEW means compare the source and correct or exclude the "
    "record if needed. Missing reading information does not prove an estimated reading. Removed "
    "duplicates are processing history, not unresolved transactions. Energy-rate coverage is informational."
)
LEGACY_NOTICE = (
    "This saved analysis predates evidence-type eligibility. Its original counts and rates are "
    "preserved using the saved all-record scope. Duplicate disposition was not recorded here. "
    "Regenerate from the review workbook or re-extract the originals for a scoped assessment."
)


def quality_notice(report: DataQualityReport) -> str:
    """Explain whether the result contains modern eligibility or preserved legacy metrics."""
    return NOTICE if report.assessment is not None else LEGACY_NOTICE


def quality_table(report: DataQualityReport) -> list[list[str]]:
    """Format stored counts and their denominators without running analysis."""
    rows = [["Check", "Count", "Eligible", "Rate", "Status"]]
    scope = report.assessment
    if scope is None:
        for label, count, rate in (
            ("Dates (all records)", report.date_parsed, report.date_parse_rate),
            ("Amounts (all records)", report.amt_complete, None),
            ("Periods (all records)", report.period_complete, report.period_completeness_rate),
            ("Readings (all records)", report.reading_classified, report.reading_classify_rate),
            ("Recorded duplicate decisions", report.duplicate_count, report.duplicate_rate),
        ):
            if label == "Amounts (all records)":
                rate = count / report.total_records if report.total_records else None
            rows.append([label, str(count), str(report.total_records), _rate(rate), "INFO"])
        return rows

    def add(label: str, count: int, total: int, *, issue: bool = False, info: bool = False) -> None:
        status = (
            "INFO"
            if info
            else "N/A"
            if not total
            else "REVIEW"
            if (count > 0 if issue else count < total)
            else "PASS"
        )
        rows.append(
            [label, str(count), str(total), _rate(count / total if total else None), status]
        )

    financial, bills = scope["financial_records"], scope["bills"]
    add("Financial record dates", scope["dated_financial_records"], financial)
    add("Financial amounts present", scope["financial_amounts"], financial)
    add("Bill charge amounts", scope["bill_charges"], bills)
    add("Transaction values", scope["transaction_amounts"], scope["transaction_records"])
    add("Bill periods complete", scope["complete_bill_periods"], bills)
    add("Bill reading basis", scope["classified_bill_readings"], bills)
    add("Bill energy rates", scope["bill_unit_rates"], bills, info=True)
    add(
        "Duplicate observations removed",
        scope["duplicates_removed"],
        report.total_records + scope["duplicates_removed"],
        info=True,
    )
    add(
        "Duplicate observations retained",
        scope["duplicates_retained"],
        report.total_records,
        issue=True,
    )
    add("Amounts with unverified meaning", scope["unverified_amounts"], financial, issue=True)
    add("Ambiguous transaction identity", scope["ambiguous_transactions"], financial, issue=True)
    add(
        "Statements needing reconciliation", scope["unreconciled_statements"], financial, issue=True
    )
    return rows


def _rate(value: float | None) -> str:
    return f"{value:.1%}" if value is not None else "—"
