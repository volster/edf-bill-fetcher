"""Interpret SAP reading provenance without treating raw codes as confirmation."""

from __future__ import annotations

import copy
from collections import defaultdict

from edf_bill_fetcher.processors.preparation import identity, scope


def sap_reading_basis(row: dict) -> str:
    """Prefer explicit source descriptions to ambiguous or contradictory type codes."""
    source = identity(row.get("Read Source")).casefold()
    if "deemed" in source or "settlement registers" in source:
        return "Deemed / estimated"
    if "estimat" in source:
        return "Estimated"
    if any(word in source for word in ("customer", "internet", "own read")):
        return "Customer"
    if "utility" in source or "data collector reading" in source:
        return "Actual"
    # A bare A is a recorded code, not proof of a physical supplier reading.
    return "Estimated" if identity(row.get("Read Type")).upper() == "E" else "Unknown"


def sap_reading_observations(rows: list[dict]) -> list[dict]:
    """Combine identical observations for display, retaining occurrence counts and raw codes."""
    groups: dict[tuple, list[dict]] = defaultdict(list)
    for row in rows:
        key = (
            scope(row),
            row.get("Meter Read Date"),
            row.get("Reading (kWh)"),
            sap_reading_basis(row),
            row.get("Read Source"),
        )
        groups[key].append(row)
    result = []
    for group in groups.values():
        row = copy.deepcopy(group[0])
        row["Reading Basis"] = sap_reading_basis(row)
        row["Reading Occurrences"] = len(group)
        row["Raw Read Types"] = ", ".join(
            sorted({identity(r.get("Read Type")) or "blank" for r in group})
        )
        result.append(row)
    return result
