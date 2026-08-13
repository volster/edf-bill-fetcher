"""Typed dataclasses for SAP ↔ EDF matching and back-billing events."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

import pandas as pd

from edf_bill_fetcher.helpers.date_utils import TimestampOrNaT


@dataclass
class SapBackBillingEvent:
    """One SAP clearing event containing one or more underlying SAP rows.

    Populated by ``detect_sap_back_billing_events``.  The underlying SAP
    rows are retained so the writer can render them as a collapsible
    sub-block beneath each event summary on sheet 'SAP Back-billing
    Events'.
    """

    clearing_doc: str
    clearing_date: TimestampOrNaT
    clearing_reason: str
    rows: list[dict[str, Any]] = field(default_factory=list)
    net_amount: float = 0.0
    has_credit_for_consum_billing: bool = False
    has_account_maintenance: bool = False
    largest_single_posting: float = 0.0
    posting_date_range: tuple[str, str] = ("", "")
    evidence_trail: str = ""
    matched_edf_invoice: str | None = None
    _cluster_unmatched_tag: dict[str, str] | None = field(default=None, repr=False)


@dataclass
class SapEdfMatch:
    """One (SAP event × matched EDF candidate) pair.

    Populated by ``match_sap_events_to_edf``.  Only SAP events that
    produced at least one EDF candidate at Low confidence or above
    appear in the returned list — unmatched events remain on Sheet 1
    only.
    """

    event: SapBackBillingEvent
    edf_record: dict[str, Any]
    confidence_band: str
    confidence_score: int
    amount_delta: float
    date_delta_days: int
    notes: str
