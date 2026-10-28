"""Link repeated statement/history postings using unique, corroborated sequences."""

from __future__ import annotations

import hashlib
from collections import defaultdict

import pandas as pd

from edf_bill_fetcher.helpers.date_utils import parse_to_sort_date
from edf_bill_fetcher.processors.preparation import identity, scope


def link_posting_sequences(rows: list[dict]) -> None:
    """Require multiple dated events and uniqueness within each source before linking.

    This is limited to account histories and historical statements, not arbitrary
    invoices or equal payments. Repeated equal events in one document are left
    for review. Existing original-invoice relationships propagate transitively.
    """
    statements: dict[str, list[dict]] = defaultdict(list)
    histories: dict[str, list[dict]] = defaultdict(list)
    for row in rows:
        row.pop("Transaction Identity Review", None)
        if row.get("Entry Type") not in {"Payment", "Credit"}:
            continue
        if not scope(row)[0] or row.get("Transaction Amount (£)") is None:
            continue
        if pd.isna(parse_to_sort_date(row.get("Date"))):
            continue
        source = identity(row.get("Source ID"))
        if not source:
            continue
        if row.get("Source") in {"HTM Account History", "PDF Account History"}:
            histories[source].append(row)
        elif row.get("Document Role") == "Historical account statement":
            statements[source].append(row)

    def indexed(items: list[dict]) -> dict[tuple, list[dict]]:
        result: dict[tuple, list[dict]] = defaultdict(list)
        for row in items:
            key = (
                scope(row),
                parse_to_sort_date(row["Date"]).isoformat(),
                row["Entry Type"],
                row["Transaction Amount (£)"],
            )
            result[key].append(row)
        return result

    for statement in statements.values():
        left = indexed(statement)
        for history in histories.values():
            right = indexed(history)
            common = left.keys() & right.keys()
            by_ledger: dict[tuple, list[tuple]] = defaultdict(list)
            for key in common:
                by_ledger[key[0]].append(key)
            for keys in by_ledger.values():
                unique = [key for key in keys if len(left[key]) == len(right[key]) == 1]
                if len({key[1] for key in unique}) < 2:
                    for key in keys:
                        for row in left[key] + right[key]:
                            row["Transaction Identity Review"] = (
                                "A single matching date/amount is insufficient to establish a shared posting sequence"
                            )
                    continue
                for key in keys:
                    group = left[key] + right[key]
                    if len(left[key]) != 1 or len(right[key]) != 1:
                        for row in group:
                            row["Transaction Identity Review"] = (
                                "Repeated equal postings within a source; event identity is ambiguous"
                            )
                        continue
                    a, b = group
                    if (
                        a.get("Balance (£)") is not None
                        and b.get("Balance (£)") is not None
                        and a["Balance (£)"] != b["Balance (£)"]
                    ):
                        for row in group:
                            row["Transaction Identity Review"] = (
                                "Cross-source posting has conflicting running balances"
                            )
                        continue
                    old = {
                        row["Reconciled Event ID"]
                        for row in group
                        if row.get("Reconciled Event ID")
                    }
                    token = (
                        min(old)
                        if old
                        else "posting:" + hashlib.sha256(repr(key).encode()).hexdigest()
                    )
                    for row in rows:
                        if row is a or row is b or (old and row.get("Reconciled Event ID") in old):
                            row["Reconciled Event ID"] = token
                    for row in group:
                        row["Transaction Identity Basis"] = (
                            "Unique posting corroborated by multiple distinct dated events in both statement and account history"
                        )
