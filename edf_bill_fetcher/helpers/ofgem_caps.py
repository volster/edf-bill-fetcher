"""Canonical OFGEM Default Tariff Cap loader — single source of truth.

The cap table lives in ``edf_bill_fetcher/data/ofgem_caps.json`` (packaged
resource); this module is the only place that reads it.  A new quarterly
cap is a pure data edit (add a quarter + bump ``as_of``) — no code change.
"""

from __future__ import annotations

import json
import os
import warnings
from importlib import resources
from typing import Any, cast


def _load_file(path: str | os.PathLike | None) -> dict[str, Any]:
    """Read and parse the caps JSON from ``path`` or the packaged resource."""
    if path is None:
        resource = resources.files("edf_bill_fetcher.data") / "ofgem_caps.json"
        with resource.open("r", encoding="utf-8") as fh:
            return cast(dict[str, Any], json.load(fh))
    with open(path, encoding="utf-8") as fh:
        return cast(dict[str, Any], json.load(fh))


def load_ofgem_caps(
    auto_carry: bool = True,
    path: str | os.PathLike | None = None,
) -> tuple[dict[str, dict], dict | None]:
    """Load OFGEM Default Tariff Cap data.

    Returns a ``(caps, latest_known)`` tuple:
    * ``caps`` maps period string (e.g., '2023-Q4') to cap values:
      ``{'unit_rate': p_per_kwh, 'standing_charge': p_per_day}``.
      The dict never carries sentinel keys — iterating ``caps.items()``
      yields only real quarters.
    * ``latest_known`` is the most recent published cap (the last quarter
      present in the data file) when ``auto_carry`` is True, else ``None``.

    Raises ``json.JSONDecodeError`` on malformed JSON and ``ValueError``
    on a missing/empty ``quarters`` object or a quarter missing numeric
    ``unit_rate``/``standing_charge``.  Warns (``UserWarning``) if an
    ``is_carry`` quarter's values differ from the previous quarter — a
    maintainer editing mistake surfaces at load without crashing a report.
    """
    data = _load_file(path)
    quarters = data.get("quarters")
    if not quarters or not isinstance(quarters, dict):
        raise ValueError("OFGEM caps file must contain a non-empty 'quarters' object")

    caps: dict[str, dict] = {}
    sorted_keys = sorted(quarters)
    for i, key in enumerate(sorted_keys):
        entry = quarters[key]
        if (
            not isinstance(entry, dict)
            or "unit_rate" not in entry
            or "standing_charge" not in entry
        ):
            raise ValueError(f"OFGEM quarter {key!r} missing 'unit_rate'/'standing_charge'")
        unit_rate = float(entry["unit_rate"])
        standing_charge = float(entry["standing_charge"])
        if entry.get("is_carry") is True and i > 0:
            prev = caps[sorted_keys[i - 1]]
            if (unit_rate, standing_charge) != (prev["unit_rate"], prev["standing_charge"]):
                warnings.warn(
                    f"OFGEM quarter {key!r} is marked carry but differs from "
                    f"{sorted_keys[i - 1]!r}",
                    UserWarning,
                    stacklevel=2,
                )
        caps[key] = {"unit_rate": unit_rate, "standing_charge": standing_charge}

    latest_known = caps[sorted_keys[-1]] if auto_carry else None
    return caps, latest_known
