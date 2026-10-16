"""Traceable corrections to canonical records, separate from original observations."""

from __future__ import annotations

import copy
import hashlib
import re
from datetime import date, datetime, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from edf_bill_fetcher.domain.analysis import AnalysisSnapshot
from edf_bill_fetcher.domain.json_values import canonical_text, object_value, rows_value

# Shared correction contract for Excel and the in-app editor. Source facts,
# extraction text, hashes, source IDs and analytical findings are never editable.
REVIEW_FIELDS = {
    "Date": "date",
    "Invoice #": "text",
    "Entry Type": "text",
    "Balance (£)": "money",
    "Period Charge (£)": "money",
    "Transaction Amount (£)": "money",
    "Period From": "date",
    "Period To": "date",
    "Units (kWh)": "number",
    "Energy Unit Rate (p/kWh)": "number",
    "Standing Chg (p/day)": "number",
    "Reading": "text",
    "Register Reading (kWh)": "number",
    "Account": "text",
    "Meter": "text",
    "Register": "text",
    "Fuel": "text",
    "Tariff": "text",
    "Amount Kind": "text",
    "Cancel/Rebill Admitted": "bool",
    "Details": "text",
}
ENTRY_TYPES = (
    "New Bill",
    "Ongoing Balance",
    "Payment",
    "Credit",
    "Refund",
    "Late Payment",
    "Charge",
    "Statement Reconciliation",
    "Other",
    "Unknown",
)


def input_value(field: str, value: Any) -> Any:
    """Validate a user-supplied scalar without interpreting formulas or fuzzy dates."""
    kind = REVIEW_FIELDS[field]
    if value is None or value == "":
        if field in {"Entry Type", "Amount Kind"}:
            raise ValueError(f"{field}: choose a value or Unknown")
        return False if kind == "bool" else None
    if kind == "bool":
        if type(value) is bool:
            return value
        if value in ("Yes", "No"):
            return value == "Yes"
        raise ValueError(f"{field}: use Yes or No")
    if kind == "date":
        if isinstance(value, datetime):
            if value.time().isoformat() != "00:00:00":
                raise ValueError(f"{field}: enter a date without a time")
            return value.date().strftime("%d/%m/%Y")
        if type(value) is date:
            return value.strftime("%d/%m/%Y")
        if isinstance(value, str):
            for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
                try:
                    return datetime.strptime(value, fmt).strftime("%d/%m/%Y")
                except ValueError:
                    pass
        raise ValueError(f"{field}: use a real Excel date, DD/MM/YYYY or YYYY-MM-DD")
    if kind in {"money", "number"}:
        try:
            if type(value) is bool or not isinstance(value, (int, float, str, Decimal)):
                raise InvalidOperation
            number = Decimal(str(value))
            if not number.is_finite() or abs(number) > Decimal("1000000000000"):
                raise InvalidOperation
            precision = Decimal("0.01") if kind == "money" else Decimal("0.000001")
            if number != number.quantize(precision):
                raise ValueError(
                    f"{field}: use at most {2 if kind == 'money' else 6} decimal places"
                )
            return float(number)
        except InvalidOperation:
            raise ValueError(f"{field}: enter a finite number, without currency symbols") from None
    if not isinstance(value, str) or len(value) > 32000:
        raise ValueError(f"{field}: enter text (format identifiers as Text in Excel)")
    if field == "Amount Kind" and value not in {"Balance", "Transaction", "Unknown"}:
        raise ValueError("Amount Kind: use Balance, Transaction or Unknown")
    if field == "Entry Type" and value not in ENTRY_TYPES:
        raise ValueError("Entry Type: choose a supported type from the dropdown")
    return value


def review_rows(snapshot: AnalysisSnapshot) -> list[dict]:
    """Expose one decision per canonical row, including previously excluded rows."""
    state = object_value(snapshot.prepared.extraction.legacy_state)
    return review_row_data(
        rows_value(snapshot.prepared.records),
        state.get("review_excluded_records", []),
        snapshot.input_fingerprint,
    )


def review_row_data(records: list[dict], excluded: list[dict], fingerprint: str) -> list[dict]:
    """Assign the same stable review identities during analysis and human review."""
    result = []
    for included, group in (
        (True, records),
        (False, excluded),
    ):
        for ordinal, record in enumerate(group):
            identity = (
                record.get("Review Row ID")
                or hashlib.sha256(
                    canonical_text([fingerprint, included, ordinal, record]).encode()
                ).hexdigest()
            )
            result.append({"id": identity, "included": included, "record": record})
    if len({row["id"] for row in result}) != len(result):
        raise ValueError("Review row identities are ambiguous")
    return result


def corrected_record(record: dict, changes: dict) -> dict:
    """Apply explicit corrections and keep dependent legacy money fields coherent."""
    row = {**copy.deepcopy(record), **copy.deepcopy(changes)}
    if "Period Charge (£)" in changes and "Transaction Amount (£)" not in changes:
        old_charge = record.get("Period Charge (£)")
        old_transaction = record.get("Transaction Amount (£)")
        linked = old_charge is not None and old_transaction in (old_charge, -old_charge)
        missing_bill_transaction = (
            old_charge is None
            and old_transaction is None
            and row.get("Entry Type")
            in {
                "New Bill",
                "Charge",
                "Ongoing Balance",
                "Late Payment",
                "Payment",
                "Credit",
                "Refund",
            }
        )
        if linked or missing_bill_transaction:
            charge = changes["Period Charge (£)"]
            row["Transaction Amount (£)"] = (
                None
                if charge is None
                else -abs(charge)
                if row.get("Entry Type") in {"Payment", "Credit"}
                else charge
            )
    if any(
        key in changes
        for key in (
            "Balance (£)",
            "Period Charge (£)",
            "Transaction Amount (£)",
            "Amount Kind",
            "Entry Type",
        )
    ):
        row["Review Explicit Amounts"] = True
        role = row.get("Amount Kind")
        row["Amount (£)"] = (
            row.get("Balance (£)" if role == "Balance" else "Transaction Amount (£)")
            if role in {"Balance", "Transaction"}
            else None
        )
        transaction = row.get("Transaction Amount (£)")
        if transaction is not None:
            kind = row.get("Entry Type")
            if (kind in {"Payment", "Credit"} and transaction > 0) or (
                kind == "Refund" and transaction < 0
            ):
                raise ValueError(
                    "Payments/credits need negative transactions; refunds need positive transactions"
                )
    if "Energy Unit Rate (p/kWh)" in changes:
        row["Unit Rate (p/kWh)"] = changes["Energy Unit Rate (p/kWh)"]
        row["Unit Rate Basis"] = (
            "Human correction" if changes["Energy Unit Rate (p/kWh)"] is not None else "Unknown"
        )
    for key in ("Period From", "Period To"):
        # Existing malformed dates remain reviewable; only reject a newly
        # contradictory period when the user actually edits its endpoints.
        if key in changes and row.get("Period From") and row.get("Period To"):
            try:
                start = input_value("Period From", row["Period From"])
                end = input_value("Period To", row["Period To"])
            except ValueError:
                continue
            if datetime.strptime(start, "%d/%m/%Y") > datetime.strptime(end, "%d/%m/%Y"):
                raise ValueError("Period From must not be later than Period To")
    return row


def apply_review(context: dict) -> tuple[list[dict], list[dict]]:
    """Validate a complete decision set and derive inputs without changing raw evidence."""
    if (
        not isinstance(context, dict)
        or set(context) != {"version", "base_fingerprint", "rows", "decisions"}
        or type(context["version"]) is not int
        or context["version"] != 1
    ):
        raise ValueError("Invalid human review context")
    if not isinstance(context["base_fingerprint"], str) or not re.fullmatch(
        r"[0-9a-f]{64}", context["base_fingerprint"]
    ):
        raise ValueError("Invalid review base fingerprint")
    rows, decisions = context["rows"], context["decisions"]
    if not isinstance(rows, list) or not isinstance(decisions, list):
        raise ValueError("Review rows and decisions must be arrays")
    indexed = {}
    for item in rows:
        if (
            not isinstance(item, dict)
            or set(item) != {"id", "included", "record"}
            or not isinstance(item["id"], str)
            or not re.fullmatch(r"[0-9a-f]{64}", item["id"])
            or type(item["included"]) is not bool
            or not isinstance(item["record"], dict)
            or item["id"] in indexed
        ):
            raise ValueError("Invalid or duplicate review row identity")
        indexed[item["id"]] = item
    resolved: list[dict] = []
    excluded: list[dict] = []
    seen: set[str] = set()
    for decision in decisions:
        if not isinstance(decision, dict) or set(decision) != {
            "row_id",
            "include",
            "changes",
            "reason",
        }:
            raise ValueError("Invalid review decision")
        identity = decision["row_id"]
        if not isinstance(identity, str) or identity not in indexed or identity in seen:
            raise ValueError("Unknown or duplicate review row identity")
        seen.add(identity)
        original = indexed[identity]
        changes = decision["changes"]
        reason = decision["reason"]
        if (
            type(decision["include"]) is not bool
            or not isinstance(changes, dict)
            or not set(changes) <= REVIEW_FIELDS.keys()
            or not isinstance(reason, str)
            or len(reason) > 4000
        ):
            raise ValueError("Invalid correction fields or review reason")
        if (changes or decision["include"] != original["included"]) and not reason.strip():
            raise ValueError("A reason is required for every correction or inclusion change")
        validated = {field: input_value(field, value) for field, value in changes.items()}
        record = corrected_record(original["record"], validated)
        record["Review Row ID"] = identity
        (resolved if decision["include"] else excluded).append(record)
    if seen != set(indexed):
        raise ValueError("Review rows were removed; use Include = No instead of deleting rows")
    return resolved, excluded


def validate_review_statuses(row_ids: set[str], statuses: object) -> None:
    """Validate per-record verification independently of financial inclusion."""
    if not isinstance(statuses, dict) or not set(statuses) <= row_ids:
        raise ValueError("Invalid record review states")
    for annotation in statuses.values():
        if (
            not isinstance(annotation, dict)
            or set(annotation) != {"status", "reason"}
            or not isinstance(annotation["status"], str)
            or annotation["status"] not in {"unreviewed", "verified", "needs_information"}
            or not isinstance(annotation["reason"], str)
            or len(annotation["reason"]) > 4000
        ):
            raise ValueError("Invalid review status or reason")
        if annotation["status"] != "unreviewed" and not annotation["reason"].strip():
            raise ValueError("Give a reason when verifying evidence or requesting information")


def review_inputs(
    snapshot: AnalysisSnapshot,
    decisions: list[dict],
    *,
    input_path: str,
    origin: str = "reviewed_workbook",
    relative_root: str | None = None,
    statuses: dict | None = None,
    revisions: list[dict] | None = None,
) -> tuple[list[dict], dict, str]:
    """Build traceable review inputs shared by Excel and the desktop editor."""
    baseline = review_rows(snapshot)
    by_id = {item["id"]: item for item in baseline}
    decisions = [
        {
            **decision,
            "changes": {
                field: input_value(field, value) for field, value in decision["changes"].items()
            },
        }
        for decision in decisions
    ]
    context = {
        "version": 1,
        "base_fingerprint": snapshot.input_fingerprint,
        "rows": baseline,
        "decisions": decisions,
    }
    resolved, excluded = apply_review(context)
    state = object_value(snapshot.prepared.extraction.legacy_state)
    history = list(state.get("review_history", []))
    reviewed_at = datetime.now(timezone.utc).isoformat()
    for decision in decisions:
        original = by_id[decision["row_id"]]
        after = corrected_record(original["record"], decision["changes"])
        changed_fields = {
            key for key in REVIEW_FIELDS if after.get(key) != original["record"].get(key)
        }
        for field in sorted(changed_fields):
            history.append(
                {
                    "at": reviewed_at,
                    "base_fingerprint": snapshot.input_fingerprint,
                    "row_id": original["id"],
                    "invoice": original["record"].get("Invoice #"),
                    "field": field,
                    "before": original["record"].get(field),
                    "after": after.get(field),
                    "reason": decision["reason"],
                }
            )
        if decision["include"] != original["included"]:
            history.append(
                {
                    "at": reviewed_at,
                    "base_fingerprint": snapshot.input_fingerprint,
                    "row_id": original["id"],
                    "invoice": original["record"].get("Invoice #"),
                    "field": "Include",
                    "before": original["included"],
                    "after": decision["include"],
                    "reason": decision["reason"],
                }
            )
    originals = [
        object_value(observation.raw) for observation in snapshot.prepared.extraction.observations
    ]
    state.update(
        {
            "original_records": originals,
            "human_review": context,
            "review_excluded_records": excluded,
            "review_history": history,
            "imported_workbook": str(Path(input_path).resolve()),
            "provenance": origin,
            "review_config": object_value(snapshot.prepared.effective_config),
            "review_source_root_relative": relative_root,
        }
    )
    if statuses is None and "human_review_statuses" in state:
        # Excel edits have no explicit verification control. Previously verified
        # values must return to review when their fields or inclusion change.
        statuses = {key: dict(value) for key, value in state["human_review_statuses"].items()}
        for decision in decisions:
            row_id = decision["row_id"]
            if row_id in statuses and (
                decision["changes"] or decision["include"] != by_id[row_id]["included"]
            ):
                statuses[row_id] = {
                    "status": "unreviewed",
                    "reason": "Excel correction changed a reviewed record; verify it again against the source.",
                }
    if statuses is not None:
        validate_review_statuses(set(by_id), statuses)
        previous = state.get("human_review_statuses", {})
        for row_id in sorted(set(previous) | set(statuses)):
            if previous.get(row_id) != statuses.get(row_id):
                history.append(
                    {
                        "at": reviewed_at,
                        "base_fingerprint": snapshot.input_fingerprint,
                        "row_id": row_id,
                        "invoice": by_id[row_id]["record"].get("Invoice #"),
                        "field": "Review status",
                        "before": previous.get(row_id, {}).get("status"),
                        "after": statuses.get(row_id, {}).get("status"),
                        "reason": statuses.get(row_id, previous.get(row_id, {})).get("reason", ""),
                    }
                )
        state["human_review_statuses"] = statuses
    if revisions is not None:
        from edf_bill_fetcher.processors.revisions import validate_decisions

        effective = [
            {"id": row["Review Row ID"], "included": included, "record": row}
            for included, rows in ((True, resolved), (False, excluded))
            for row in rows
        ]
        revisions = validate_decisions(effective, revisions)
        previous = {(d["older"], d["newer"]): d for d in state.get("bill_revision_decisions", [])}
        current = {(d["older"], d["newer"]): d for d in revisions}
        for pair in sorted(previous.keys() | current.keys()):
            before_link, after_link = previous.get(pair), current.get(pair)
            if before_link != after_link:
                history.append(
                    {
                        "at": reviewed_at,
                        "base_fingerprint": snapshot.input_fingerprint,
                        "row_id": pair[0],
                        "invoice": by_id[pair[0]]["record"].get("Invoice #"),
                        "field": "Revision relationship",
                        "before": str(before_link) if before_link else None,
                        "after": str(after_link) if after_link else None,
                        "reason": current.get(pair, previous.get(pair, {}))["reason"],
                    }
                )
        state["bill_revision_decisions"] = revisions
    state.pop("records", None)
    return originals, state, snapshot.prepared.extraction.request.as_of.isoformat()
