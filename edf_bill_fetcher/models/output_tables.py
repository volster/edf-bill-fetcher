"""Readable evidence tables shared by document outputs; no financial inference."""

from __future__ import annotations

from typing import Any

import pandas as pd

from edf_bill_fetcher.helpers.date_utils import parse_to_sort_date
from edf_bill_fetcher.helpers.formatting import fmt_money, report_date
from edf_bill_fetcher.processors.preparation import identity

EVIDENCE_HEADERS = [
    "Date",
    "Account / ledger",
    "Entry type",
    "Balance (£)",
    "Transaction (£)",
    "Period charge (£)",
    "Billing period",
    "Invoice #",
    "Source",
    "Attachment",
    "Pages",
    "Details",
]

EVIDENCE_WIDTHS = (2.7, 3.5, 2.3, 2.3, 3, 2.7, 3, 3, 1.7, 4.7, 1.4, 4)

EVIDENCE_NOTE = (
    "Records are sorted by date. Balance is an account observation; transaction is a "
    "separate charge/payment/credit/refund movement. Blank amounts are unknown, not zero. "
    "Page numbers refer to the source PDF. Readings, consumption, rates and full extraction "
    "details remain in the Excel workbook."
)


def ledger_label(row: dict[str, Any]) -> str:
    """Retain the account/fuel/meter/register context without inventing identity."""
    parts = [identity(row.get("Account")) or "Unassigned account"]
    for field in ("Fuel", "Meter", "Register"):
        if value := identity(row.get(field)):
            parts.append(f"{field}: {value}" if field != "Fuel" else value)
    return " / ".join(parts)


def evidence_cells(row: dict[str, Any]) -> list[str]:
    """Format explicit accepted fields, retaining full invoice and source references."""
    first, last = identity(row.get("Page From")), identity(row.get("Page To"))
    pages = first + (f"–{last}" if last and last != first else "") if first else last
    period = " – ".join(
        report_date(row.get(field)) or "Unknown" for field in ("Period From", "Period To")
    )
    return [
        report_date(row.get("Date")),
        ledger_label(row),
        identity(row.get("Entry Type")),
        fmt_money(row.get("Balance (£)")),
        fmt_money(row.get("Transaction Amount (£)")),
        fmt_money(row.get("Period Charge (£)")),
        period if period != "Unknown – Unknown" else "",
        identity(row.get("Invoice #")),
        identity(row.get("Source")),
        identity(row.get("Attachment Name")),
        pages,
        identity(row.get("Details"))[:150],
    ]


def timeline_table(df: pd.DataFrame, flags: list) -> list[list[str]]:
    """Separate observed balances from transactions and labelled automated flag values."""
    events = []
    for row in df.to_dict("records"):
        events.append(
            (
                row.get("Date"),
                [
                    report_date(row.get("Date")),
                    identity(row.get("Entry Type")),
                    fmt_money(row.get("Balance (£)")),
                    fmt_money(row.get("Transaction Amount (£)")),
                    f"{ledger_label(row)} — {identity(row.get('Details'))[:150]}",
                ],
            )
        )
    for kind, day, amount, detail, severity in flags:
        events.append(
            (
                day,
                [
                    report_date(day),
                    f"FLAG ({severity}): {kind}",
                    "",
                    "",
                    f"Flag value: {fmt_money(amount)}. {detail}",
                ],
            )
        )

    def sort_key(event):
        parsed = parse_to_sort_date(event[0])
        return pd.Timestamp.max if pd.isna(parsed) else parsed

    events.sort(key=sort_key)
    return [
        ["Date", "Event", "Balance (£)", "Transaction (£)", "Context / detail"],
        *[row for _, row in events],
    ]


def revision_table(model: dict) -> list[list[str]]:
    """Present bounded comparisons from stored decisions; retain every source reference."""
    links = sorted(
        model["relationships"],
        key=lambda link: {"confirmed": 0, "rejected": 1, "suggested": 2}[link["status"]],
    )
    rows = [["Decision / ledger", "Earlier document", "Later document", "Comparison / reason"]]
    for link in links[:30]:
        a, b = link["earlier"], link["later"]
        documents = []
        for row, included in ((a, link["earlier_included"]), (b, link["later_included"])):
            documents.append(
                "\n".join(
                    (
                        identity(row.get("Invoice #")) or "Invoice unavailable",
                        report_date(row.get("Date")),
                        f"Charge: {fmt_money(row.get('Period Charge (£)')) or 'Unknown'}; transaction: {fmt_money(row.get('Transaction Amount (£)')) or 'Unknown'}",
                        "Included" if included else "Excluded",
                        identity(row.get("Attachment Name")) or "Source unavailable",
                        "Pages: "
                        + (
                            "–".join(
                                identity(row.get(k))
                                for k in ("Page From", "Page To")
                                if identity(row.get(k))
                            )
                            or "Unknown"
                        ),
                    )
                )
            )
        difference = fmt_money(link.get("charge_difference"))
        rows.append(
            [
                link["status"].title()
                + " · "
                + link["kind"].replace("_", " ")
                + "\n"
                + ledger_label(a),
                *documents,
                (f"Document charge difference: {difference}. " if difference else "")
                + "Changed: "
                + (", ".join(link["changed_fields"]) or "No compared fields")
                + ".\n"
                + (link["reason"] or link["basis"]),
            ]
        )
    return rows


def revision_note(model: dict) -> str:
    """Explain inference limits, separate inclusion decisions and the complete workbook view."""
    return (
        str(model["notice"])
        + " Document charge differences are not refunds or changes to account liability. "
        + (
            "Showing 30 relationships, prioritising explicit decisions. "
            if len(model["relationships"]) > 30
            else ""
        )
        + ("Automatic suggestions were capped at 500. " if model["truncated"] else "")
        + "The Bill revisions worksheet contains all stored relationships and full references."
    )
