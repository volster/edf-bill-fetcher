"""Human reconciliation of repeated disputes, credits and proposed balance reductions."""

from __future__ import annotations

import copy
import hashlib
import re
from datetime import date
from decimal import Decimal, InvalidOperation

from edf_bill_fetcher.domain.json_values import canonical_text
from edf_bill_fetcher.domain.money import money_text, parse_money
from edf_bill_fetcher.helpers.identity import account_key
from edf_bill_fetcher.processors.preparation import identity
from edf_bill_fetcher.processors.rebilling_ledger import schedule_digest
from edf_bill_fetcher.processors.revisions import BILLS

NOTICE = (
    "These are human proposals, not confirmed write-offs. Each group represents one consolidated "
    "challenge; successive invoice amounts are never automatically added. Amounts represent the "
    "requested reduction in the balance, including applicable VAT. Relief already received reduces "
    "the further reduction requested; a cancellation offset is history, not a second deduction. "
    "Totals require reviewed, current groups and explicit resolution of potential overlaps. "
    "Illustrative balances use dated observations and are not statements of current liability."
)
CONTEXT_DEFAULTS = {"complaint_summary": "", "requested_outcome": "", "include_pack": True}
GROUP_FIELDS = {
    "id",
    "title",
    "status",
    "members",
    "challenged",
    "reduction",
    "credits",
    "reason",
    "account_link_reason",
    "source_digest",
}


def empty_state() -> dict:
    """Return independent defaults without adding a claim or presuming entitlement."""
    return {"version": 1, "groups": [], "separations": [], "context": dict(CONTEXT_DEFAULTS)}


def has_active_claims(model: dict | None) -> bool:
    """Distinguish unresolved work from a history containing only withdrawn proposals."""
    return bool(
        model
        and (
            model["unresolved_positions"]
            or any(group["input"]["status"] != "withdrawn" for group in model["groups"])
        )
    )


def _text(value: object, limit: int = 4000, *, required: bool = False) -> str:
    if not isinstance(value, str) or len(value) > limit or (required and not value.strip()):
        raise ValueError(
            f"Enter {'nonempty ' if required else ''}text of at most {limit} characters"
        )
    return value


def _digest(value: object) -> str:
    if not isinstance(value, str) or not re.fullmatch(r"[0-9a-f]{64}", value):
        raise ValueError("Invalid reconciliation evidence digest")
    return value


def _money(value: object) -> str | None:
    if value is None or value == "":
        return None
    try:
        if isinstance(value, bool) or not isinstance(value, (int, float, str, Decimal)):
            raise InvalidOperation
        amount = Decimal(str(value))
        if (
            not amount.is_finite()
            or not 0 <= amount <= Decimal("1000000000000")
            or amount != amount.quantize(Decimal(".01"))
        ):
            raise InvalidOperation
        return money_text(amount)
    except InvalidOperation:
        raise ValueError(
            "Enter non-negative money with at most two decimal places, without currency symbols"
        ) from None


def validate_state(items: list[dict], raw: object = None) -> dict:
    """Validate human inputs, preserving unresolved evidence as review issues instead of guesses."""
    if raw is None:
        return empty_state()
    if (
        not isinstance(raw, dict)
        or set(raw) != {"version", "groups", "separations", "context"}
        or type(raw["version"]) is not int
        or raw["version"] != 1
    ):
        raise ValueError("Unsupported claim reconciliation state")
    groups, separations, context = raw["groups"], raw["separations"], raw["context"]
    if (
        not isinstance(groups, list)
        or len(groups) > 1000
        or not isinstance(separations, list)
        or len(separations) > 10000
    ):
        raise ValueError("Too many claim groups or overlap decisions")
    if (
        not isinstance(context, dict)
        or set(context) != set(CONTEXT_DEFAULTS)
        or type(context["include_pack"]) is not bool
    ):
        raise ValueError("Invalid complaint context")
    for field in ("complaint_summary", "requested_outcome"):
        _text(context[field])
    records = {item["id"]: item["record"] for item in items}
    result, ids = [], set()
    for group in groups:
        if not isinstance(group, dict) or set(group) != GROUP_FIELDS:
            raise ValueError("Invalid claim group fields")
        key = group["id"]
        if (
            not isinstance(key, str)
            or not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]{0,63}", key)
            or key.casefold() in ids
        ):
            raise ValueError(
                "Claim IDs must be unique short identifiers containing letters, numbers, hyphens or underscores"
            )
        ids.add(key.casefold())
        status = group["status"]
        if not isinstance(status, str) or status not in {"draft", "reviewed", "withdrawn"}:
            raise ValueError("Claim review status must be draft, reviewed or withdrawn")
        _text(group["title"], 200, required=True)
        _text(group["reason"], required=status == "reviewed")
        _text(group["account_link_reason"])
        _digest(group["source_digest"])
        members = group["members"]
        if (
            not isinstance(members, list)
            or len(members) > 10000
            or any(
                not isinstance(key, str)
                or key not in records
                or records[key].get("Entry Type") not in BILLS
                for key in members
            )
            or len(set(members)) != len(members)
        ):
            raise ValueError("A claim must reference unique, known bill records")
        if not isinstance(group["credits"], list) or len(group["credits"]) > 10000:
            raise ValueError("Invalid claim credit allocations")
        allocations, seen = [], set()
        for credit in group["credits"]:
            if not isinstance(credit, dict) or set(credit) != {
                "row_id",
                "amount",
                "treatment",
                "reason",
            }:
                raise ValueError("Invalid credit allocation fields")
            row_id, treatment = credit["row_id"], credit["treatment"]
            if (
                not isinstance(row_id, str)
                or row_id not in records
                or records[row_id].get("Entry Type") not in BILLS | {"Credit"}
                or not isinstance(treatment, str)
                or treatment not in {"already_applied_relief", "cancellation_offset"}
                or (row_id, treatment) in seen
            ):
                raise ValueError(
                    "Credit allocations need a known source and a unique treatment per group"
                )
            seen.add((row_id, treatment))
            amount = _money(credit["amount"])
            if amount is None or Decimal(amount) <= 0:
                raise ValueError("Credit allocations must be positive amounts")
            _text(credit["reason"], required=True)
            allocations.append({**credit, "amount": amount})
        normalized = {
            **copy.deepcopy(group),
            "challenged": _money(group["challenged"]),
            "reduction": _money(group["reduction"]),
            "credits": sorted(allocations, key=lambda item: (item["row_id"], item["treatment"])),
            "members": sorted(members),
        }
        challenged, reduction = normalized["challenged"], normalized["reduction"]
        if reduction is not None and (
            challenged is None or Decimal(reduction) > Decimal(challenged)
        ):
            raise ValueError("A proposed reduction needs a challenged amount and cannot exceed it")
        if status == "reviewed" and (not members or challenged is None or reduction is None):
            raise ValueError(
                "Reviewing a group requires bills and its consolidated challenged/reduction amounts"
            )
        result.append(normalized)
    known = {group["id"] for group in result}
    decisions, seen_pairs = [], set()
    for decision in separations:
        if not isinstance(decision, dict) or set(decision) != {"groups", "reason", "source_digest"}:
            raise ValueError("Invalid overlap review fields")
        pair = decision["groups"]
        if (
            not isinstance(pair, list)
            or len(pair) != 2
            or any(not isinstance(key, str) or key not in known for key in pair)
            or pair[0] >= pair[1]
            or tuple(pair) in seen_pairs
        ):
            raise ValueError(
                "Overlap review must reference two distinct existing groups in identifier order"
            )
        seen_pairs.add(tuple(pair))
        _text(decision["reason"], required=True)
        _digest(decision["source_digest"])
        decisions.append(copy.deepcopy(decision))
    return {
        "version": 1,
        "groups": sorted(result, key=lambda group: group["id"]),
        "separations": sorted(decisions, key=lambda decision: decision["groups"]),
        "context": copy.deepcopy(context),
    }


def group_digest(items: list[dict], ledger: dict, group: dict) -> str:
    """Bind a proposal to its scope, current positions and the entire compatible history."""
    docs = {doc["row_id"]: doc for doc in ledger["documents"]}
    ids = set(group["members"]) | {credit["row_id"] for credit in group["credits"]}
    explicit = {}
    for doc in ledger["documents"]:
        for link in doc["incoming"]:
            if link["status"] != "suggested":
                explicit[(link["older"], link["newer"])] = {
                    key: link[key] for key in ("older", "newer", "kind", "status", "reason")
                }
    revisions = list(explicit.values())
    payload = {
        "group": {key: value for key, value in group.items() if key != "source_digest"},
        "evidence": [
            {
                "row_id": key,
                "context": schedule_digest(items, revisions, key),
                "position": docs.get(key, {}).get("position"),
                "needs_review": docs.get(key, {}).get("needs_review", False),
            }
            for key in sorted(ids)
        ],
    }
    return hashlib.sha256(canonical_text(payload).encode()).hexdigest()


def separation_digest(items: list[dict], ledger: dict, left: dict, right: dict) -> str:
    """Invalidate a separate-charges decision when either group or its evidence changes."""
    pair = sorted((left, right), key=lambda group: group["id"])
    return hashlib.sha256(
        canonical_text([group_digest(items, ledger, group) for group in pair]).encode()
    ).hexdigest()


def _scope(records: list[dict]) -> dict:
    return {
        "accounts": sorted(
            {
                account_key(identity(row.get("Account")))
                for row in records
                if account_key(identity(row.get("Account")))
            }
        ),
        **{
            field: sorted(
                {identity(row.get(field)).casefold() for row in records if identity(row.get(field))}
            )
            for field in ("Fuel", "Meter", "Register")
        },
    }


def _period(row: dict) -> tuple[date, date] | None:
    try:
        start, end = date.fromisoformat(row["period_from"]), date.fromisoformat(row["period_to"])
        return (start, end) if start <= end else None
    except (KeyError, TypeError, ValueError):
        return None


def _potential_overlap(left: dict, right: dict, docs: dict) -> str | None:
    a, b = left["input"], right["input"]
    if set(a["members"]) & set(b["members"]):
        return "shared_documents"
    x, y = left["scope"], right["scope"]
    if any(
        x[field] and y[field] and set(x[field]).isdisjoint(y[field])
        for field in ("Fuel", "Meter", "Register")
    ):
        return None
    if not set(x["accounts"]) & set(y["accounts"]) and not (
        x["Meter"] and x["Meter"] == y["Meter"]
    ):
        return "unverified_account_relationship"
    if not a["members"] or not b["members"]:
        return "unknown_periods"
    periods_a = [_period(docs[key]) for key in a["members"]]
    periods_b = [_period(docs[key]) for key in b["members"]]
    if any(period is None for period in periods_a + periods_b):
        return "unknown_periods"
    if any(
        max(before[0], after[0]) <= min(before[1], after[1])
        for before in periods_a
        for after in periods_b
        if before and after
    ):
        return "overlapping_periods"
    return None


def build_reconciliation(items: list[dict], ledger: dict, raw: object = None) -> dict:
    """Total reviewed economic claims once, withholding unresolved or stale proposals."""
    state = validate_state(items, raw)
    indexed = {item["id"]: item for item in items}
    docs = {doc["row_id"]: doc for doc in ledger["documents"]}
    groups = []
    live_digests = {}
    for group in state["groups"]:
        ids = list(
            dict.fromkeys([*group["members"], *(credit["row_id"] for credit in group["credits"])])
        )
        rows = [indexed[key]["record"] for key in ids]
        scope = _scope(rows)
        issues = []
        if not group["members"]:
            issues.append("No challenged bills have been assigned to this group.")
        if any(not account_key(identity(row.get("Account"))) for row in rows):
            issues.append("An assigned record has no known account; verify its identity.")
        if any(len(scope[field]) > 1 for field in ("Fuel", "Meter", "Register")):
            issues.append("Conflicting fuel, meter or register identities cannot be consolidated.")
        if len(scope["accounts"]) > 1 and (
            not group["account_link_reason"].strip()
            or len(scope["Meter"]) != 1
            or len(scope["Fuel"]) != 1
            or any(not identity(row.get("Meter")) or not identity(row.get("Fuel")) for row in rows)
        ):
            issues.append(
                "Linking accounts needs a shared known meter/fuel and an explicit source-backed account-link reason."
            )
        live_digests[group["id"]] = group_digest(items, ledger, group)
        stale = group["source_digest"] != live_digests[group["id"]]
        if stale:
            issues.append("The group or its evidence changed; verify and save the group again.")
        for key in group["members"]:
            doc = docs[key]
            if not doc["included"]:
                issues.append(f"Excluded bill: {doc['record'].get('Invoice #') or key}.")
            if (
                doc["needs_review"]
                or not doc["position"]
                or doc["position"]["status"] != "challenged"
            ):
                issues.append(
                    f"Bill {doc['record'].get('Invoice #') or key} needs a current challenged position."
                )
        relief, offsets = Decimal("0.00"), Decimal("0.00")
        allocated: dict[str, Decimal] = {}
        for credit in group["credits"]:
            key, amount = credit["row_id"], Decimal(credit["amount"])
            allocated[key] = allocated.get(key, Decimal("0")) + amount
            actual = parse_money(indexed[key]["record"].get("Transaction Amount (£)"))
            if (
                indexed[key]["record"].get("Entry Type") not in BILLS | {"Credit"}
                or not indexed[key]["included"]
                or actual is None
                or actual >= 0
                or allocated[key] > abs(actual)
            ):
                issues.append(
                    f"Credit allocation for {indexed[key]['record'].get('Invoice #') or key} exceeds an included, known negative transaction."
                )
            if credit["treatment"] == "already_applied_relief":
                relief += amount
            else:
                offsets += amount
        reduction = parse_money(group["reduction"])
        if reduction is not None and relief > reduction:
            issues.append("Relief already received exceeds this group's proposed reduction.")
        further = reduction - relief if reduction is not None and relief <= reduction else None
        if group["status"] == "draft":
            issues.append("Group is still a draft.")
        groups.append(
            {
                "input": group,
                "scope": scope,
                "needs_review": stale,
                "issues": list(dict.fromkeys(issues)),
                "relief_received": money_text(relief),
                "cancellation_offsets": money_text(offsets),
                "further_reduction": money_text(further),
                "ready": group["status"] == "reviewed" and not issues,
            }
        )
    active = [group for group in groups if group["input"]["status"] != "withdrawn"]
    reviewed = [group for group in active if group["input"]["status"] == "reviewed"]
    allocations: dict[str, list[tuple[dict, Decimal]]] = {}
    for group in reviewed:
        for credit in group["input"]["credits"]:
            allocations.setdefault(credit["row_id"], []).append((group, Decimal(credit["amount"])))
    for key, uses in allocations.items():
        actual = parse_money(indexed[key]["record"].get("Transaction Amount (£)"))
        if (
            actual is None
            or actual >= 0
            or sum((amount for _, amount in uses), Decimal("0")) > abs(actual)
        ):
            for group, _ in uses:
                group["ready"] = False
                group["issues"].append(
                    "This credit has been allocated more than once across reviewed groups; reconcile its total use."
                )
    separations = {tuple(decision["groups"]): decision for decision in state["separations"]}
    overlaps = []
    for index, left in enumerate(active):
        for right in active[index + 1 :]:
            kind = _potential_overlap(left, right, docs)
            if kind is None:
                continue
            pair = sorted((left["input"]["id"], right["input"]["id"]))
            decision = separations.get(tuple(pair))
            current_digest = hashlib.sha256(
                canonical_text([live_digests[key] for key in pair]).encode()
            ).hexdigest()
            stale = bool(decision and decision["source_digest"] != current_digest)
            resolved = bool(decision and not stale and kind != "shared_documents")
            overlaps.append(
                {
                    "groups": pair,
                    "kind": kind,
                    "resolved": resolved,
                    "needs_review": stale,
                    "reason": decision["reason"] if decision else "",
                }
            )
            if not resolved:
                message = (
                    f"Groups {pair[0]} and {pair[1]} share a document; merge or reassign it."
                    if kind == "shared_documents"
                    else f"Potential overlap between {pair[0]} and {pair[1]}; merge them or substantiate separate charges."
                )
                for group in (left, right):
                    group["ready"] = False
                    group["issues"].append(message)
    covered = {key for group in active if group["ready"] for key in group["input"]["members"]}
    unresolved = [
        doc["row_id"]
        for doc in ledger["documents"]
        if doc["included"]
        and doc["position"]
        and doc["position"]["status"] == "challenged"
        and doc["row_id"] not in covered
    ]
    ready_groups = [group for group in active if group["ready"]]
    subtotal = sum((Decimal(group["further_reduction"]) for group in ready_groups), Decimal("0.00"))
    ready = bool(active) and all(group["ready"] for group in active) and not unresolved
    ledgers: dict[str, dict] = {}
    for group in ready_groups:
        scope = group["scope"]
        key = hashlib.sha256(canonical_text(scope).encode()).hexdigest()
        entry = ledgers.setdefault(
            key,
            {
                "scope": scope,
                "groups": [],
                "reduction": Decimal("0"),
                "observed_balance": None,
                "balance_date": None,
                "illustrative_balance": None,
            },
        )
        entry["groups"].append(group["input"]["id"])
        entry["reduction"] += Decimal(group["further_reduction"])
    for entry in ledgers.values():
        scope = entry["scope"]
        candidates = []
        for item in items:
            row = item["record"]
            if (
                not item["included"]
                or account_key(identity(row.get("Account"))) not in scope["accounts"]
            ):
                continue
            if any(
                scope[field]
                != ([identity(row.get(field)).casefold()] if identity(row.get(field)) else [])
                for field in ("Fuel", "Meter", "Register")
            ):
                continue
            balance = parse_money(row.get("Balance (£)"))
            try:
                parts = identity(row.get("Date")).split("/")
                when = (
                    date(int(parts[2]), int(parts[1]), int(parts[0]))
                    if len(parts) == 3
                    else date.fromisoformat(identity(row.get("Date")))
                )
            except (ValueError, IndexError):
                continue
            if balance is not None:
                candidates.append((when, balance))
        if candidates:
            last = max(when for when, _ in candidates)
            values = {balance for when, balance in candidates if when == last}
            if len(values) == 1:
                balance = values.pop()
                entry.update(
                    observed_balance=money_text(balance),
                    balance_date=last.isoformat(),
                    illustrative_balance=money_text(balance - entry["reduction"]),
                )
        entry["reduction"] = money_text(entry["reduction"])
    return {
        "version": 1,
        "notice": NOTICE,
        "input": state,
        "groups": groups,
        "overlaps": overlaps,
        "unresolved_positions": unresolved,
        "reviewed_subtotal": money_text(subtotal),
        "proposed_reduction": money_text(subtotal) if ready else None,
        "ready": ready,
        "ledgers": list(ledgers.values()),
    }
