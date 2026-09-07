"""Run-diff processor: compare two record sets (added / removed / changed).

Pure-pandas, no writers/UI dependency.  ``diff_records`` diffs two
``--records-json``-shaped record lists by a canonical key built from
``key_fields`` (default ``("Date", "Amount (£)", "Source")``) and
returns ``{"added": [...], "removed": [...], "changed": [...]}``
where every output list is sorted deterministically by ``Date``.
"""

from __future__ import annotations

import math
from collections.abc import Mapping, Sequence
from typing import Any

import numpy as np
import pandas as pd

DEFAULT_KEY_FIELDS: tuple[str, ...] = ("Date", "Amount (£)", "Source")

# Unit-separator join so normalized field values can never collide.
_KEY_SEPARATOR = "\x1f"
_AMOUNT_FIELD = "Amount (£)"


def _normalize_key_value(value: Any) -> str:
    """Normalize one key-field value to a canonical string.

    ``None`` and float NaN (plus pandas NA markers) normalize to the
    empty string so two missing values match each other rather than
    matching real data.
    """
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    try:
        if bool(pd.isna(value)):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value)


def _canonical_key(record: Mapping[str, Any], key_fields: Sequence[str]) -> str:
    """Build the canonical match key for one record from its key fields."""
    return _KEY_SEPARATOR.join(_normalize_key_value(record.get(field)) for field in key_fields)


def _is_na(value: Any) -> bool:
    """Return True when ``value`` is None, NaN, NaT, or another pandas NA marker."""
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def _values_equal(left: Any, right: Any) -> bool:
    """NaN-aware scalar equality: two missing values compare equal."""
    if left is right:
        return True
    if _is_na(left) or _is_na(right):
        return _is_na(left) and _is_na(right)
    try:
        return bool(left == right)
    except (TypeError, ValueError):
        return False


def _native(value: Any) -> Any:
    """Convert pandas/numpy scalars back to native Python values."""
    if isinstance(value, np.integer):
        return int(value)
    if isinstance(value, np.floating):
        return float(value)
    if isinstance(value, np.bool_):
        return bool(value)
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    return value


def _date_sort_key(
    record: Mapping[str, Any], key_fields: Sequence[str]
) -> tuple[int, pd.Timestamp, str]:
    """Deterministic sort key: (missing?, parsed Date, canonical key)."""
    date_value = record.get("Date")
    timestamp: pd.Timestamp = pd.NaT
    if date_value is not None and not (isinstance(date_value, float) and math.isnan(date_value)):
        try:
            converted = pd.to_datetime(date_value, errors="coerce")
            if isinstance(converted, pd.Timestamp):
                timestamp = converted
        except (TypeError, ValueError):
            pass
    if pd.isna(timestamp):
        return (1, pd.Timestamp.min, _canonical_key(record, key_fields))
    return (0, timestamp, _canonical_key(record, key_fields))


def _native_record(row: dict[str, Any]) -> dict[str, Any]:
    """Convert every cell of a DataFrame-derived record to native Python."""
    return {field: _native(value) for field, value in row.items()}


def diff_records(
    old_records: list[dict[str, Any]],
    new_records: list[dict[str, Any]],
    key_fields: Sequence[str] = DEFAULT_KEY_FIELDS,
) -> dict[str, list[dict[str, Any]]]:
    """Diff two record sets and return ``added`` / ``removed`` / ``changed``.

    Default matching also requires equal invoice numbers when available.
    Identical occurrences are consumed one-to-one before changed rows are
    paired. Ambiguous groups remain additions/removals.

    Matching happens in two stages, both on canonical string keys
    built from the key fields (None / NaN normalize to an empty
    component):

    1. Exact key match: records whose full keys appear in both sets
       are compared on their remaining (non-key) fields; any
       difference yields a ``changed`` entry.
    2. Amount-relaxed match: when ``"Amount (£)"`` is one of the key
       fields, records left unmatched by stage 1 are re-matched on the
       other key fields alone, so a corrected amount on the same bill
       surfaces as a ``changed`` entry with ``Amount (£)`` in
       ``changed_fields`` rather than a phantom remove+add pair.

    Each ``changed`` entry has the shape ``{"old_row": {...},
    "new_row": {...}, "changed_fields": [names], "old_values": [...],
    "new_values": [...]}`` with the two value lists aligned to
    ``changed_fields``.  Every output list is sorted deterministically
    by ``Date`` (with the canonical key as a tie-breaker), so the same
    input pair always produces the same output regardless of input
    order.
    """
    old_rows = [_native_record(row) for row in old_records]
    new_rows = [_native_record(row) for row in new_records]
    remaining_old = set(range(len(old_rows)))
    remaining_new = set(range(len(new_rows)))
    changed: list[dict[str, Any]] = []

    def identity(row: dict[str, Any], fields: Sequence[str]) -> tuple[str, ...]:
        values = tuple(_normalize_key_value(row.get(f)) for f in fields)
        if tuple(key_fields) == DEFAULT_KEY_FIELDS:
            invoice = _normalize_key_value(row.get("Invoice #")).strip()
            if invoice in ("N/A", "None", "nan"):
                invoice = ""
            values += (invoice,)
        return values

    def row_order(row: dict[str, Any]) -> tuple[tuple[str, str], ...]:
        return tuple(sorted((k, _normalize_key_value(v)) for k, v in row.items()))

    def pair(old_i: int, new_i: int) -> None:
        old, new = old_rows[old_i], new_rows[new_i]
        fields = list(dict.fromkeys([*old, *new]))
        differences = [f for f in fields if not _values_equal(old.get(f), new.get(f))]
        if differences:
            changed.append(
                {
                    "old_row": old,
                    "new_row": new,
                    "changed_fields": differences,
                    "old_values": [old.get(f) for f in differences],
                    "new_values": [new.get(f) for f in differences],
                }
            )
        remaining_old.remove(old_i)
        remaining_new.remove(new_i)

    stages = [tuple(key_fields)]
    relaxed = tuple(f for f in key_fields if f != _AMOUNT_FIELD)
    if _AMOUNT_FIELD in key_fields and relaxed:
        stages.append(relaxed)
    for fields in stages:
        old_groups: dict[tuple[str, ...], list[int]] = {}
        new_groups: dict[tuple[str, ...], list[int]] = {}
        for i in sorted(remaining_old, key=lambda i: row_order(old_rows[i])):
            old_groups.setdefault(identity(old_rows[i], fields), []).append(i)
        for i in sorted(remaining_new, key=lambda i: row_order(new_rows[i])):
            new_groups.setdefault(identity(new_rows[i], fields), []).append(i)
        for key in sorted(old_groups.keys() & new_groups.keys()):
            old_group, new_group = old_groups[key], new_groups[key]
            # Consume identical occurrences first, preserving multiplicity.
            for old_i in old_group:
                for new_i in new_group:
                    if new_i not in remaining_new:
                        continue
                    old, new = old_rows[old_i], new_rows[new_i]
                    if all(_values_equal(old.get(f), new.get(f)) for f in old.keys() | new.keys()):
                        pair(old_i, new_i)
                        break
            unmatched_old = [i for i in old_group if i in remaining_old]
            unmatched_new = [i for i in new_group if i in remaining_new]
            # Ambiguous groups remain additions/removals instead of fabricated changes.
            if len(unmatched_old) == len(unmatched_new) == 1:
                pair(unmatched_old[0], unmatched_new[0])

    def output_order(row: dict[str, Any]) -> tuple:
        return (*_date_sort_key(row, key_fields), row_order(row))

    return {
        "added": sorted((new_rows[i] for i in remaining_new), key=output_order),
        "removed": sorted((old_rows[i] for i in remaining_old), key=output_order),
        "changed": sorted(changed, key=lambda entry: output_order(entry["new_row"])),
    }
