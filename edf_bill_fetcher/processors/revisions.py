"""Conservative bill relationships, kept separate from ledger transactions."""

from __future__ import annotations

import re

import pandas as pd

from edf_bill_fetcher.helpers.date_utils import parse_to_sort_date
from edf_bill_fetcher.helpers.identity import account_key
from edf_bill_fetcher.processors.preparation import identity, number

BILLS = {"New Bill", "Ongoing Balance", "Charge", "Late Payment"}
NOTICE = (
    "Revision links describe documents. Confirming a replacement does not remove ledger "
    "charges or credits: a cancelled charge may already be offset by a credit. Inclusion "
    "decisions remain explicit in Review & Corrections. Suggestions require human confirmation."
)
COMPARE_FIELDS = (
    "Invoice #",
    "Date",
    "Period From",
    "Period To",
    "Period Charge (£)",
    "Transaction Amount (£)",
    "Balance (£)",
    "Units (kWh)",
    "Energy Unit Rate (p/kWh)",
    "Reading",
    "Account",
    "Fuel",
    "Meter",
    "Register",
    "Attachment Name",
    "Page From",
    "Page To",
    "Source ID",
)


def compatible(left: dict, right: dict) -> bool:
    """Require a known shared account and never cross conflicting ledger identities."""
    account = account_key(identity(left.get("Account")))
    return (
        bool(account)
        and account == account_key(identity(right.get("Account")))
        and all(
            not (a := identity(left.get(field)))
            or not (b := identity(right.get(field)))
            or a.casefold() == b.casefold()
            for field in ("Fuel", "Meter", "Register")
        )
    )


def validate_decisions(items: list[dict], decisions: object) -> list[dict]:
    """Validate explicit links and reject cycles, conflicting identities and invalid dates."""
    if not isinstance(decisions, list) or len(decisions) > 10000:
        raise ValueError("Invalid bill revision decisions")
    records = {item["id"]: item["record"] for item in items}
    result, seen = [], set()
    graph: dict[str, list[str]] = {}
    for decision in decisions:
        if not isinstance(decision, dict) or set(decision) != {
            "older",
            "newer",
            "kind",
            "status",
            "reason",
        }:
            raise ValueError("Invalid bill revision relationship")
        older, newer = decision["older"], decision["newer"]
        if (
            not isinstance(older, str)
            or not isinstance(newer, str)
            or older not in records
            or newer not in records
            or older == newer
            or (older, newer) in seen
        ):
            raise ValueError("Unknown, duplicate or identical revision endpoints")
        seen.add((older, newer))
        left, right = records[older], records[newer]
        kind, status, reason = decision["kind"], decision["status"], decision["reason"]
        if (
            not isinstance(kind, str)
            or kind not in {"replacement", "cancellation_credit"}
            or not isinstance(status, str)
            or status not in {"confirmed", "rejected"}
            or not isinstance(reason, str)
            or not reason.strip()
            or len(reason) > 4000
        ):
            raise ValueError("Revision decisions need a supported type, status and reason")
        if left.get("Entry Type") not in BILLS or right.get("Entry Type") not in (
            BILLS if kind == "replacement" else {"Credit"}
        ):
            raise ValueError(
                "A replacement must link two bills; a cancellation credit must link a bill to a credit"
            )
        if status == "confirmed":
            if not compatible(left, right):
                raise ValueError(
                    "Confirmed revisions require the same known account and compatible fuel/meter/register. Clear conflicting links before changing identity fields."
                )
            before, after = (
                parse_to_sort_date(left.get("Date")),
                parse_to_sort_date(right.get("Date")),
            )
            if pd.isna(before) or pd.isna(after) or after < before:
                raise ValueError(
                    "Confirmed revisions need usable dates in earlier-to-later order. Correct dates or clear the link first."
                )
            graph.setdefault(older, []).append(newer)
        result.append(dict(decision))
    # Iterative topological traversal also handles long chains without recursion.
    degree = dict.fromkeys(records, 0)
    for children in graph.values():
        for child in children:
            degree[child] += 1
    ready = [key for key, value in degree.items() if value == 0]
    visited = 0
    while ready:
        node = ready.pop()
        visited += 1
        for child in graph.get(node, []):
            degree[child] -= 1
            if degree[child] == 0:
                ready.append(child)
    if visited != len(records):
        raise ValueError("Bill revision links cannot form a cycle")
    return result


def build_revisions(items: list[dict], decisions: object = None, *, limit: int = 500) -> dict:
    """Suggest overlapping bills or explicitly referenced credits; never infer cancellation."""
    decisions = validate_decisions(items, [] if decisions is None else decisions)
    indexed = {item["id"]: item for item in items}
    links = {(d["older"], d["newer"]): {**d, "basis": "Human decision"} for d in decisions}
    groups: dict[str, list] = {}
    for item in items:
        row = item["record"]
        if row.get("Entry Type") not in BILLS:
            continue
        start, end, day = (
            parse_to_sort_date(row.get(k)) for k in ("Period From", "Period To", "Date")
        )
        if pd.isna(start) or pd.isna(end) or pd.isna(day) or end <= start:
            continue
        account = account_key(identity(row.get("Account")))
        if account:
            groups.setdefault(account, []).append((start, end, day, item))
    truncated = False
    for group in groups.values():
        active: list[tuple] = []
        for start, end, day, item in sorted(
            group, key=lambda entry: (entry[0], entry[2], entry[3]["id"])
        ):
            active = [entry for entry in active if entry[1] > start]
            for other_start, other_end, other_day, other in active:
                if other_day == day or not compatible(other["record"], item["record"]):
                    continue
                older, newer = (other, item) if other_day < day else (item, other)
                key = (older["id"], newer["id"])
                if key not in links:
                    if len(links) >= limit + len(decisions):
                        truncated = True
                        break
                    links[key] = {
                        "older": key[0],
                        "newer": key[1],
                        "kind": "replacement",
                        "status": "suggested",
                        "reason": "",
                        "basis": "Same consumption period"
                        if start == other_start and end == other_end
                        else "Overlapping consumption periods; check partial replacement",
                    }
            if truncated:
                break
            active.append((start, end, day, item))
        if truncated:
            break
    invoices: dict[str, list[dict]] = {}
    for item in items:
        row = item["record"]
        if row.get("Entry Type") in BILLS and (invoice := identity(row.get("Invoice #"))):
            invoices.setdefault(invoice.casefold(), []).append(item)
    for item in items:
        row = item["record"]
        if row.get("Entry Type") != "Credit":
            continue
        tokens = set(
            re.findall(
                r"[\w-]+",
                " ".join(identity(row.get(k)) for k in ("Invoice #", "Details", "Source PDF Text"))[
                    :100000
                ].casefold(),
            )
        )
        for token in tokens & invoices.keys():
            for older in invoices[token]:
                key = older["id"], item["id"]
                left_day, right_day = (
                    parse_to_sort_date(r.get("Date")) for r in (older["record"], row)
                )
                if (
                    key in links
                    or not compatible(older["record"], row)
                    or pd.isna(left_day)
                    or pd.isna(right_day)
                    or right_day < left_day
                ):
                    continue
                if len(links) >= limit + len(decisions):
                    truncated = True
                    continue
                links[key] = {
                    "older": key[0],
                    "newer": key[1],
                    "kind": "cancellation_credit",
                    "status": "suggested",
                    "reason": "",
                    "basis": "Credit document references this invoice; verify cancellation and amount",
                }
    relationships = []
    for link in links.values():
        left, right = indexed[link["older"]], indexed[link["newer"]]
        a, b = left["record"], right["record"]
        old_charge, new_charge = (
            number(a.get("Period Charge (£)")),
            number(b.get("Period Charge (£)")),
        )
        relationships.append(
            {
                **link,
                "earlier": {key: a.get(key) for key in COMPARE_FIELDS},
                "later": {key: b.get(key) for key in COMPARE_FIELDS},
                "earlier_included": left["included"],
                "later_included": right["included"],
                "changed_fields": [
                    field for field in COMPARE_FIELDS if a.get(field) != b.get(field)
                ],
                "charge_difference": new_charge - old_charge
                if new_charge is not None and old_charge is not None
                else None,
            }
        )
    return {"version": 1, "notice": NOTICE, "relationships": relationships, "truncated": truncated}
