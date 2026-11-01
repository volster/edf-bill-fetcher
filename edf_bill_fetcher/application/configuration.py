"""Validate explicit configuration controls without importing any UI or web framework."""

from __future__ import annotations

import math
import re
from datetime import date
from typing import get_args, get_origin, get_type_hints

from edf_bill_fetcher.models.config import ConfigDict
from edf_bill_fetcher.models.report_sections import REPORT_SECTIONS


def validate_settings(config: dict) -> None:
    """Reject unknown fields, ambiguous types, invalid dates and nonfinite numeric inputs."""
    fields = get_type_hints(ConfigDict)
    if set(config) - fields.keys():
        raise ValueError("Unknown analytical setting")
    for key, value in config.items():
        kind = fields[key]
        if kind is bool and type(value) is not bool:
            raise ValueError(f"{key} must be true or false")
        if kind is float and (
            type(value) not in (int, float) or not math.isfinite(value) or value < 0
        ):
            raise ValueError(f"{key} must be a nonnegative finite number")
        if kind is int and (type(value) is not int or not 0 <= value <= 1_000_000):
            raise ValueError(f"{key} must be a bounded nonnegative integer")
        if kind is str and not isinstance(value, str):
            raise ValueError(f"{key} must be text")
        if get_origin(kind) is list and (
            not isinstance(value, list)
            or any(not isinstance(item, get_args(kind)[0]) for item in value)
        ):
            raise ValueError(f"{key} must be a list of text values")
    if "report_sections" in config:
        sections = config["report_sections"]
        if len(set(sections)) != len(sections) or not set(sections) <= (
            {section.key for section in REPORT_SECTIONS} | {"cover", "toc"}
        ):
            raise ValueError("Unknown or duplicate report section")
    if "as_of" in config:
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", config["as_of"]):
            raise ValueError("Valuation date must use YYYY-MM-DD")
        date.fromisoformat(config["as_of"])
