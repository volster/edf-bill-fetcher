"""Derive analytical ledger context while preserving every observed identity."""

from __future__ import annotations

import copy
from collections import defaultdict

from edf_bill_fetcher.helpers.identity import account_key
from edf_bill_fetcher.processors.preparation import identity

FIELDS = ("Account", "Fuel", "Meter", "Register")


def fuel_key(value: object) -> str:
    """Normalize equivalent fuel labels, leaving absent or unknown values empty."""
    text = identity(value).casefold()
    if text in {"na", "not applicable"}:
        return ""
    return {"elec": "electricity", "electric": "electricity"}.get(text, text)


def resolve_ledger_context(
    records: list[dict], sources: list[dict]
) -> tuple[list[dict], list[dict]]:
    """Use explicit one-to-one SAP aliases and uniquely supported missing context.

    Derived fields describe grouping within the supplied evidence, not extracted
    facts. Conflicting aliases, fuels, meters and registers remain separate.
    An unassigned account is never inferred from the rest of the case.
    """
    rows = copy.deepcopy(records + sources)
    forward: dict[str, set[str]] = defaultdict(set)
    reverse: dict[str, set[str]] = defaultdict(set)
    evidence: dict[tuple[str, str], set[str]] = defaultdict(set)
    for row in sources:
        legacy = account_key(identity(row.get("SAP Account Number")))
        modern = account_key(identity(row.get("Kraken ID")))
        origin = identity(row.get("Source ID")) or identity(row.get("Source File"))
        if legacy and modern and legacy != modern and origin:
            forward[legacy].add(modern)
            reverse[modern].add(legacy)
            evidence[legacy, modern].add(origin)
    aliases = {}
    for legacy, moderns in forward.items():
        if len(moderns) == 1:
            modern = next(iter(moderns))
            if len(reverse[modern]) == 1 and legacy not in reverse and modern not in forward:
                aliases[legacy] = aliases[modern] = legacy
    for row in rows:
        for field in FIELDS:
            row.pop("Ledger " + field, None)
        row.pop("Ledger Identity Basis", None)
        account = account_key(
            identity(row.get("Account"))
            or identity(row.get("SAP Account Number"))
            or identity(row.get("Kraken ID"))
        )
        row["Ledger Account"] = aliases.get(account, account)
        row["Ledger Fuel"] = fuel_key(row.get("Fuel"))
        row["Ledger Meter"] = identity(row.get("Meter")).casefold()
        row["Ledger Register"] = identity(row.get("Register")).casefold()
    for field, parents in (("Fuel", ()), ("Meter", ("Fuel",)), ("Register", ("Fuel", "Meter"))):
        known: dict[tuple, set[str]] = defaultdict(set)
        for row in rows:
            key = tuple(row["Ledger " + part] for part in ("Account", *parents))
            if all(key) and row["Ledger " + field]:
                known[key].add(row["Ledger " + field])
        for row in rows:
            key = tuple(row["Ledger " + part] for part in ("Account", *parents))
            if all(key) and not row["Ledger " + field] and len(known[key]) == 1:
                row["Ledger " + field] = next(iter(known[key]))
    for row in rows:
        observed = (
            account_key(identity(row.get("Account"))),
            fuel_key(row.get("Fuel")),
            identity(row.get("Meter")).casefold(),
            identity(row.get("Register")).casefold(),
        )
        derived = tuple(row["Ledger " + field] for field in FIELDS)
        if observed != derived:
            origins = sorted(
                {
                    origin
                    for (legacy, modern), ids in evidence.items()
                    if derived[0] == aliases.get(legacy)
                    for origin in ids
                }
            )
            row["Ledger Identity Basis"] = (
                "Analytical grouping only; observed identities are unchanged. "
                "Explicit SAP account aliases and uniquely supported context in the supplied evidence."
                + (" Alias sources: " + ", ".join(origins) if origins else "")
            )
    return rows[: len(records)], rows[len(records) :]
