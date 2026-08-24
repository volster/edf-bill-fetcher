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


def _build_changed(merged: pd.DataFrame, excluded_fields: set[str]) -> list[dict[str, Any]]:
    """Turn an old×new merge into ``changed`` entries.

    Each entry reports the full old/new rows plus ``changed_fields``
    (fields outside ``excluded_fields`` that differ) and the aligned
    ``old_values`` / ``new_values`` lists.  The internal ``_key`` join
    column is never surfaced as a record field.
    """
    bases: list[str] = []
    for column in merged.columns:
        if column.endswith("__old"):
            base = column[: -len("__old")]
            if base != "_key" and base not in bases:
                bases.append(base)

    changed: list[dict[str, Any]] = []
    for _, pair in merged.iterrows():
        old_row: dict[str, Any] = {}
        new_row: dict[str, Any] = {}
        changed_fields: list[str] = []
        old_values: list[Any] = []
        new_values: list[Any] = []
        for base in bases:
            old_value = _native(pair[f"{base}__old"])
            new_value = _native(pair[f"{base}__new"])
            old_row[base] = old_value
            new_row[base] = new_value
            if base not in excluded_fields and not _values_equal(old_value, new_value):
                changed_fields.append(base)
                old_values.append(old_value)
                new_values.append(new_value)
        if changed_fields:
            changed.append(
                {
                    "old_row": old_row,
                    "new_row": new_row,
                    "changed_fields": changed_fields,
                    "old_values": old_values,
                    "new_values": new_values,
                }
            )
    return changed


def _native_record(row: dict[str, Any]) -> dict[str, Any]:
    """Convert every cell of a DataFrame-derived record to native Python."""
    return {field: _native(value) for field, value in row.items()}


def diff_records(
    old_records: list[dict[str, Any]],
    new_records: list[dict[str, Any]],
    key_fields: Sequence[str] = DEFAULT_KEY_FIELDS,
) -> dict[str, list[dict[str, Any]]]:
    """Diff two record sets and return ``added`` / ``removed`` / ``changed``.

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
    old_df = pd.DataFrame(list(old_records))
    new_df = pd.DataFrame(list(new_records))
    old_df["_key"] = [_canonical_key(record, key_fields) for record in old_records]
    new_df["_key"] = [_canonical_key(record, key_fields) for record in new_records]

    # Stage 1: exact key matches → compare remaining fields.
    stage1 = old_df.merge(new_df, on="_key", how="inner", suffixes=("__old", "__new"))
    changed = _build_changed(stage1, set(key_fields))
    matched_old_keys: set[str] = set(stage1["_key"])
    matched_new_keys: set[str] = set(stage1["_key"])

    # Stage 2: amount-relaxed re-match for records left unmatched.
    if _AMOUNT_FIELD in key_fields:
        identity_fields = tuple(field for field in key_fields if field != _AMOUNT_FIELD)
        if identity_fields:
            remaining_old = old_df[~old_df["_key"].isin(matched_old_keys)].copy()
            remaining_new = new_df[~new_df["_key"].isin(matched_new_keys)].copy()
            if not remaining_old.empty and not remaining_new.empty:
                remaining_old["_relaxed_key"] = [
                    _canonical_key(record, identity_fields)
                    for record in old_records
                    if _canonical_key(record, key_fields) not in matched_old_keys
                ]
                remaining_new["_relaxed_key"] = [
                    _canonical_key(record, identity_fields)
                    for record in new_records
                    if _canonical_key(record, key_fields) not in matched_new_keys
                ]
                stage2 = remaining_old.merge(
                    remaining_new, on="_relaxed_key", how="inner", suffixes=("__old", "__new")
                )
                if not stage2.empty:
                    changed.extend(_build_changed(stage2, set(identity_fields)))
                    matched_old_keys.update(stage2["_key__old"])
                    matched_new_keys.update(stage2["_key__new"])

    added_df = new_df[~new_df["_key"].isin(matched_new_keys)]
    removed_df = old_df[~old_df["_key"].isin(matched_old_keys)]
    added = [_native_record(row) for row in added_df.drop(columns="_key").to_dict("records")]
    removed = [_native_record(row) for row in removed_df.drop(columns="_key").to_dict("records")]

    return {
        "added": sorted(added, key=lambda row: _date_sort_key(row, key_fields)),
        "removed": sorted(removed, key=lambda row: _date_sort_key(row, key_fields)),
        "changed": sorted(changed, key=lambda entry: _date_sort_key(entry["new_row"], key_fields)),
    }
