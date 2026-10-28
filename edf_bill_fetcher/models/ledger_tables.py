"""Format stored ledger observations consistently without recalculation."""

from edf_bill_fetcher.helpers.formatting import fmt_money, report_date

NOTICE = (
    "Dated observations by analytical ledger; these are not current balances. "
    "Account references are linked only by explicit SAP alias evidence. Missing fuel, meter or register "
    "context is grouped only where unique within the supplied evidence; original fields remain unchanged. "
    "Unassigned or conflicting identities stay separate. Charges and credits are recorded movements, "
    "not a reconstructed closing balance."
)


def ledger_table(rows: list[dict]) -> list[list[str]]:
    """Present accepted per-ledger balances with their dates and observed references."""
    result = [
        [
            "Account references / ledger",
            "Records",
            "Charges",
            "Net credits",
            "First observed balance",
            "Last observed balance",
        ]
    ]
    for row in rows:
        account, fuel, meter, register = row["scope"]
        label = ", ".join(row["references"]) or "Unassigned account"
        label += " / " + " / ".join(part for part in (fuel, meter, register) if part)
        result.append(
            [
                label.rstrip(" /"),
                str(row["records"]),
                fmt_money(row["charges"]),
                fmt_money(row["net_credits"]),
                fmt_money(row["opening_balance"])
                + (" on " + report_date(row["opening_date"]) if row["opening_date"] else ""),
                fmt_money(row["closing_balance"])
                + (" on " + report_date(row["closing_date"]) if row["closing_date"] else ""),
            ]
        )
    return result
