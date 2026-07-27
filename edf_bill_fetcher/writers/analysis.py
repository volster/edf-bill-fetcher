"""Analysis sheet writers — back-billing, rebilling, meter readings, contract history."""

from __future__ import annotations

# Compatibility re-export — all implementation lives in ``edf_collector.py``.
# Task 5 will extract these into this submodule file directly.
from edf_collector import (
    EDF_NAVY,
    EDF_ORANGE,
    MEDIUM_GREY,
    _disclosed_label,
    _reading_type_to_aem,
    _safe_to_datetime,
    build_evidence_index,
    detect_back_billing,
    detect_meter_rollover,
    detect_rebilling,
    infer_contracts,
)

__all__ = [
    "EDF_NAVY",
    "EDF_ORANGE",
    "MEDIUM_GREY",
    "_disclosed_label",
    "_reading_type_to_aem",
    "_safe_to_datetime",
    "build_evidence_index",
    "detect_back_billing",
    "detect_meter_rollover",
    "detect_rebilling",
    "infer_contracts",
]
