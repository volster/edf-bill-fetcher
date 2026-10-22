"""Traceable rebilling histories and explicit, non-additive human dispute positions."""

from __future__ import annotations

import hashlib
import re
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation

from edf_bill_fetcher.domain.json_values import canonical_text
from edf_bill_fetcher.processors.preparation import identity, number
from edf_bill_fetcher.processors.revisions import BILLS, COMPARE_FIELDS, compatible

NOTICE = (
    "This ledger preserves original charges, cancellations, credits and replacements. "
    "Document differences are not write-offs. Proposed reductions are human dispute positions, "
    "not confirmed entitlements, and do not change the recorded account balance. "
    "Schedule amounts are not summed: successive bills may demand the same charges. "
    "Credits shown in the history must not be claimed a second time. A net credit bill can "
    "still contain positive replacement charges; its net movement is not a cap on a challenge."
)
STATUSES = ("unreviewed", "challenged", "not_challenged", "withdrawn")
SCHEDULE_FIELDS = (
    "status",
    "challenged_amount",
    "proposed_reduction",
    "first_demand_date",
    "basis",
    "edf_request",
)
DOCUMENT_FIELDS = (
    *COMPARE_FIELDS,
    "Entry Type",
    "Reading Date",
    "Closing Reading Type",
    "Document Date",
    "Posting Date",
    "First Billed Date",
    "Billing Date Status",
    "Register Reading (kWh)",
    "Standing Chg (p/day)",
    "Gross Electricity Charges (£)",
    "Gross Electricity Charges Basis",
    "Sub Periods",
    "Rebilling Components",
    "Content Hash",
    "OCR Review Required",
)


def money(value: object) -> str | None:
    """Represent observed amounts as decimal money without silently filling gaps."""
    parsed = number(value)
    if parsed is None:
        return None
    try:
        rounded = Decimal(str(value if isinstance(value, Decimal) else parsed)).quantize(
            Decimal(".01"), rounding=ROUND_HALF_UP
        )
    except InvalidOperation:
        return None
    return "0.00" if rounded == 0 else format(rounded, ".2f")


def _date(value: object) -> str | None:
    if not isinstance(value, str) or not value:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt).date().isoformat()
        except ValueError:
            pass
    return None


def _period(start: object, end: object) -> tuple[str, str] | None:
    a, b = _date(start), _date(end)
    return (a, b) if a and b and a <= b else None


def components(row: dict) -> list[dict]:
    """Keep only explicitly observed components, with their original document pages.

    A missing amount or a partial period is never filled by a day-based allocation.
    Older saved results can expose their already extracted Sub Periods too.
    """
    result = []
    raw = row.get("Sub Periods")
    if isinstance(raw, str):
        for part in raw.split(";"):
            fields = part.strip().split("|")
            if len(fields) != 5 or not (period := _period(fields[0], fields[1])):
                continue
            units, rate, amount = (number(value) for value in fields[2:])
            result.append(
                {
                    "kind": "energy",
                    "period_from": period[0],
                    "period_to": period[1],
                    "units": units,
                    "rate": rate,
                    "amount": money(amount),
                    "detail": "Extracted energy component, excluding VAT",
                }
            )
    extracted_energy = bool(result)
    facts = row.get("Rebilling Components") or component_facts(identity(row.get("Source PDF Text")))
    if isinstance(facts, list):
        for fact in facts:
            if isinstance(fact, dict) and fact.get("kind") in {
                "standing",
                "vat",
                "cancellation",
                "discount",
                "energy",
            }:
                if fact["kind"] == "energy" and extracted_energy:
                    continue
                result.append(
                    {
                        "kind": fact["kind"],
                        "period_from": _date(fact.get("period_from")),
                        "period_to": _date(fact.get("period_to")),
                        "units": number(fact.get("units")),
                        "rate": number(fact.get("rate")),
                        "amount": money(fact.get("amount")),
                        "detail": identity(fact.get("detail")),
                    }
                )
    return result


def component_facts(text: str) -> list[dict]:
    """Read explicitly labelled standing charges, VAT and cancellation amounts only."""
    result: list[dict] = []
    from edf_bill_fetcher.helpers.date_utils import parse_to_display_date

    for match in re.finditer(
        r"([\d,]+(?:\.\d+)?)\s*kWh\s*@\s*Electricity used\s*([\d.]+)p/kWh\s*£([\d,]+\.\d{2})",
        text,
        re.I,
    ):
        periods = list(
            re.finditer(
                r"\((\d{1,2} [A-Za-z]+ \d{4})\s*-\s*(\d{1,2} [A-Za-z]+ \d{4})\)",
                text[max(0, match.start() - 2500) : match.start()],
            )
        )
        period = (
            _period(*(parse_to_display_date(value) for value in periods[-1].groups()))
            if periods
            else None
        )
        result.append(
            {
                "kind": "energy",
                "units": number(match[1]),
                "rate": number(match[2]),
                "amount": number(match[3]),
                "period_from": period[0] if period else None,
                "period_to": period[1] if period else None,
                "detail": match[0] + " (excluding VAT)",
            }
        )
    for match in re.finditer(
        r"Standing charge\s*\(?(\d[\d,]*) days?\s*@\s*([\d.]+)p(?: per day|/day)\)?\s*£([\d,]+\.\d{2})",
        text,
        re.I,
    ):
        result.append(
            {
                "kind": "standing",
                "units": number(match[1]),
                "rate": number(match[2]),
                "amount": number(match[3]),
                "detail": match[0],
            }
        )
    for match in re.finditer(
        r"(?<!Cancelled )\bVAT\s*@\s*([\d.]+)%\s+(?:of|on)\s+£[\d,]+\.\d{2}\s+£([\d,]+\.\d{2})",
        text,
        re.I,
    ):
        result.append(
            {
                "kind": "vat",
                "rate": number(match[1]),
                "amount": number(match[2]),
                "detail": match[0],
            }
        )
    for match in re.finditer(
        r"We cancelled your (?:electricity|gas) charges\s*\(excluding VAT\)\s*£([\d,]+\.\d{2})",
        text,
        re.I,
    ):
        result.append(
            {
                "kind": "cancellation",
                "amount": -abs(number(match[1]) or 0),
                "detail": match[0] + " credit (excluding VAT)",
            }
        )
    for match in re.finditer(
        r"Cancelled VAT\s*\([^)]+\)\s*was\s*£([\d,]+\.\d{2})",
        text,
        re.I,
    ):
        result.append(
            {
                "kind": "cancellation",
                "amount": -abs(number(match[1]) or 0),
                "detail": match[0] + " credit",
            }
        )
    for match in re.finditer(
        r"Direct Debit Discount\s*@\s*[\d.]+%\s+of\s+£[\d,]+\.\d{2}\s+£([\d,]+\.\d{2})",
        text,
        re.I,
    ):
        result.append(
            {"kind": "discount", "amount": -abs(number(match[1]) or 0), "detail": match[0]}
        )
    return result


def schedule_digest(items: list[dict], revisions: list[dict], row_id: str) -> str:
    """Bind a position to all compatible ledger evidence and explicit relationships."""
    target = next(item["record"] for item in items if item["id"] == row_id)
    relevant = [
        item for item in items if item["id"] == row_id or compatible(target, item["record"])
    ]
    ids = {item["id"] for item in relevant}
    return hashlib.sha256(
        canonical_text(
            {
                "documents": [
                    {
                        "id": item["id"],
                        "included": item["included"],
                        "record": {key: item["record"].get(key) for key in DOCUMENT_FIELDS},
                    }
                    for item in sorted(relevant, key=lambda item: item["id"])
                ],
                "revisions": sorted(
                    [link for link in revisions if link["older"] in ids or link["newer"] in ids],
                    key=lambda link: (link["older"], link["newer"]),
                ),
            }
        ).encode()
    ).hexdigest()


def validate_schedule(items: list[dict], decisions: object) -> list[dict]:
    """Strictly validate human positions independently of inferred financial liability."""
    if not isinstance(decisions, list) or len(decisions) > 100000:
        raise ValueError("Invalid dispute schedule")
    records = {item["id"]: item["record"] for item in items}
    result, seen = [], set()
    for decision in decisions:
        if not isinstance(decision, dict) or set(decision) != {
            "row_id",
            "source_digest",
            *SCHEDULE_FIELDS,
        }:
            raise ValueError("Invalid dispute schedule fields")
        key = decision["row_id"]
        if (
            not isinstance(key, str)
            or key not in records
            or key in seen
            or records[key].get("Entry Type") not in BILLS
        ):
            raise ValueError("Schedule must reference a unique bill, not a payment or credit")
        seen.add(key)
        status = decision["status"]
        if not isinstance(status, str) or status not in STATUSES:
            raise ValueError("Invalid dispute position status")
        if not isinstance(decision["source_digest"], str) or not re.fullmatch(
            r"[0-9a-f]{64}", decision["source_digest"]
        ):
            raise ValueError("Invalid dispute position evidence digest")
        for field in ("basis", "edf_request"):
            if not isinstance(decision[field], str) or len(decision[field]) > 4000:
                raise ValueError("Dispute explanations must be text of at most 4000 characters")
        if status != "unreviewed" and not decision["basis"].strip():
            raise ValueError("Explain the evidence and scope of this dispute decision")
        normalized = dict(decision)
        for field in ("challenged_amount", "proposed_reduction"):
            value = decision[field]
            if value is None or value == "":
                normalized[field] = None
                continue
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
                normalized[field] = format(amount.quantize(Decimal(".01")), ".2f")
            except InvalidOperation:
                raise ValueError(
                    "Dispute amounts must be non-negative money with at most two decimal places"
                ) from None
        challenged, reduction = normalized["challenged_amount"], normalized["proposed_reduction"]
        if reduction is not None and (
            challenged is None or Decimal(reduction) > Decimal(challenged)
        ):
            raise ValueError("Proposed reduction needs a challenged amount and cannot exceed it")
        value = decision["first_demand_date"]
        normalized["first_demand_date"] = None if value in (None, "") else _date(value)
        if value not in (None, "") and normalized["first_demand_date"] is None:
            raise ValueError("First demand date must be DD/MM/YYYY or YYYY-MM-DD")
        if (
            normalized["first_demand_date"] is not None
            and date.fromisoformat(normalized["first_demand_date"]).year < 2
        ):
            raise ValueError("First demand date must allow a preceding 12-month reference")
        result.append(normalized)
    return result


def _compare(left: dict, right: dict) -> list[dict]:
    earlier, later = components(left), components(right)
    comparisons = []
    for part in later:
        if part["kind"] != "energy":
            continue
        period = part["period_from"], part["period_to"]
        matches = [
            old
            for old in earlier
            if old["kind"] == "energy" and (old["period_from"], old["period_to"]) == period
        ]
        unique_later = (
            sum(
                (p["period_from"], p["period_to"]) == period and p["kind"] == "energy"
                for p in later
            )
            == 1
        )
        if not all(period) or len(matches) != 1 or not unique_later:
            comparisons.append(
                {
                    "period_from": period[0],
                    "period_to": period[1],
                    "status": "unallocated",
                    "charge_difference": None,
                    "units_effect": None,
                    "rate_effect": None,
                    "residual": None,
                }
            )
            continue
        old = matches[0]
        values = [
            old["units"],
            part["units"],
            old["rate"],
            part["rate"],
            old["amount"],
            part["amount"],
        ]
        if any(value is None for value in values):
            comparisons.append(
                {
                    "period_from": period[0],
                    "period_to": period[1],
                    "status": "unallocated",
                    "charge_difference": None,
                    "units_effect": None,
                    "rate_effect": None,
                    "residual": None,
                }
            )
            continue
        old_units, new_units, old_rate, new_rate, old_amount, new_amount = (
            Decimal(str(value)) for value in values
        )
        unit_effect = (new_units - old_units) * old_rate / 100
        rate_effect = new_units * (new_rate - old_rate) / 100
        difference = new_amount - old_amount
        comparisons.append(
            {
                "period_from": period[0],
                "period_to": period[1],
                "status": "matched",
                "charge_difference": money(difference),
                "units_effect": money(unit_effect),
                "rate_effect": money(rate_effect),
                "residual": money(
                    difference
                    - Decimal(money(unit_effect) or "0")
                    - Decimal(money(rate_effect) or "0")
                ),
            }
        )
    later_periods = {
        (part["period_from"], part["period_to"]) for part in later if part["kind"] == "energy"
    }
    for old in earlier:
        if old["kind"] == "energy" and (old["period_from"], old["period_to"]) not in later_periods:
            comparisons.append(
                {
                    "period_from": old["period_from"],
                    "period_to": old["period_to"],
                    "status": "unallocated",
                    "charge_difference": None,
                    "units_effect": None,
                    "rate_effect": None,
                    "residual": None,
                }
            )
    return comparisons


def _other_components(left: dict, right: dict) -> list[dict]:
    earlier_period = _period(left.get("Period From"), left.get("Period To"))
    same_period = earlier_period is not None and earlier_period == _period(
        right.get("Period From"), right.get("Period To")
    )
    before, after = components(left), components(right)
    result = []
    for kind in ("standing", "vat", "discount"):
        amounts = []
        for parts in (before, after):
            matching = [part["amount"] for part in parts if part["kind"] == kind]
            amounts.append(
                sum((Decimal(value) for value in matching), Decimal("0"))
                if matching and all(value is not None for value in matching)
                else None
            )
        old, new = amounts
        if old is not None or new is not None:
            result.append(
                {
                    "kind": kind,
                    "earlier_amount": money(old),
                    "later_amount": money(new),
                    "difference": money(new - old)
                    if same_period and old is not None and new is not None
                    else None,
                }
            )
    return result


def build_ledger(items: list[dict], revision_model: dict, decisions: object = None) -> dict:
    """Compute the complete document history once; renderers only consume this result."""
    decisions = validate_schedule(items, [] if decisions is None else decisions)
    positions = {decision["row_id"]: decision for decision in decisions}
    explicit = [
        {key: link[key] for key in ("older", "newer", "kind", "status", "reason")}
        for link in revision_model["relationships"]
        if link["status"] != "suggested"
    ]
    links = revision_model["relationships"]
    indexed = {item["id"]: item for item in items}
    documents = []
    for item in items:
        row, key = item["record"], item["id"]
        if row.get("Entry Type") not in BILLS | {"Credit"}:
            continue
        position = positions.get(key)
        stale = bool(
            position and position["source_digest"] != schedule_digest(items, explicit, key)
        )
        dates = {
            name: _date(row.get(field))
            for name, field in (
                ("invoice_date", "Date"),
                ("reading_date", "Reading Date"),
                ("period_from", "Period From"),
                ("period_to", "Period To"),
            )
        }
        first = position["first_demand_date"] if position else None
        cutoff = None
        if first:
            when = date.fromisoformat(first)
            cutoff = when.replace(
                year=when.year - 1, day=min(when.day, 28) if when.month == 2 else when.day
            ).isoformat()
        parts = components(row)
        documents.append(
            {
                "row_id": key,
                "record": {field: row.get(field) for field in DOCUMENT_FIELDS},
                "included": item["included"],
                **dates,
                "first_demand_date": first,
                "cutoff_date": cutoff,
                "components": parts,
                "position": position,
                "needs_review": stale,
                "incoming": [link for link in links if link["newer"] == key],
                "outgoing": [link for link in links if link["older"] == key],
            }
        )
    documents.sort(
        key=lambda doc: (
            identity(doc["record"].get("Account")),
            identity(doc["record"].get("Fuel")),
            identity(doc["record"].get("Meter")),
            identity(doc["record"].get("Register")),
            doc["invoice_date"] or "9999-12-31",
            doc["row_id"],
        )
    )
    comparisons = [
        {
            "older": link["older"],
            "newer": link["newer"],
            "status": link["status"],
            "same_period": _period(
                link["earlier"].get("Period From"), link["earlier"].get("Period To")
            )
            is not None
            and _period(link["earlier"].get("Period From"), link["earlier"].get("Period To"))
            == _period(link["later"].get("Period From"), link["later"].get("Period To")),
            "other_components": _other_components(
                indexed[link["older"]]["record"], indexed[link["newer"]]["record"]
            ),
            "components": _compare(
                indexed[link["older"]]["record"], indexed[link["newer"]]["record"]
            ),
        }
        for link in links
        if link["kind"] == "replacement"
    ]
    return {
        "version": 1,
        "notice": NOTICE,
        "documents": documents,
        "comparisons": comparisons,
        "suggestions_truncated": revision_model["truncated"],
    }
