"""Explicit legacy scalar conversion at the immutable snapshot boundary."""

from __future__ import annotations

import math
from datetime import date, datetime
from decimal import Decimal
from typing import Any

import numpy as np
import pandas as pd


def to_snapshot_json(
    value: object, *, path: str = "$", diagnostics: list[dict] | None = None
) -> Any:
    """Convert known legacy values to JSON, rejecting unsupported objects by path."""
    if value is None or value is pd.NA or value is pd.NaT:
        return None
    if isinstance(value, np.generic):
        return to_snapshot_json(value.item(), path=path, diagnostics=diagnostics)
    if isinstance(value, (str, bool, int)):
        return value
    if isinstance(value, (float, Decimal)):
        if not math.isfinite(value):
            if diagnostics is not None:
                diagnostics.append({"code": "NONFINITE_VALUE", "path": path})
            return None
        return str(value) if isinstance(value, Decimal) else value
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, dict):
        if any(not isinstance(key, str) for key in value):
            raise ValueError(f"Non-string key at {path}")
        return {
            key: to_snapshot_json(item, path=f"{path}.{key}", diagnostics=diagnostics)
            for key, item in value.items()
        }
    if isinstance(value, (list, tuple, np.ndarray)):
        return [
            to_snapshot_json(item, path=f"{path}[{index}]", diagnostics=diagnostics)
            for index, item in enumerate(value)
        ]
    if isinstance(value, (set, frozenset)) and path.endswith(".seen_pdf_hashes"):
        if all(isinstance(item, str) for item in value):
            return sorted(value)
    raise ValueError(f"Unsupported snapshot value at {path}: {type(value).__name__}")
