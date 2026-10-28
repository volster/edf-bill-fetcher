"""Immutable financial and analysis results independent of renderers."""

from __future__ import annotations

import hashlib
import math
import re
from dataclasses import dataclass
from decimal import Decimal

from edf_bill_fetcher.domain.evidence import PreparedSnapshot
from edf_bill_fetcher.domain.json_values import FrozenJSON, canonical_text, object_value, rows_value


@dataclass(frozen=True)
class FinancialSummary:
    """Supported legacy totals with explicit net-credit and balance meanings."""

    charges: Decimal
    net_credits: Decimal
    opening_balance: Decimal | None
    closing_balance: Decimal | None
    mean_daily_charge: float

    def __post_init__(self) -> None:
        """Require finite totals and retain unavailable balances."""
        for value in (self.charges, self.net_credits, self.opening_balance, self.closing_balance):
            if value is not None and (not isinstance(value, Decimal) or not value.is_finite()):
                raise ValueError("Financial summary requires finite decimal money")
        if not math.isfinite(self.mean_daily_charge):
            raise ValueError("Daily charge must be finite")


def analytical_fingerprint(prepared: PreparedSnapshot, methods: FrozenJSON) -> str:
    """Identify supported inputs independently of run locations and output options."""
    # Positive analytical configuration list; rendering controls do not affect results.
    config = object_value(prepared.effective_config)
    included = (
        "as_of",
        "analysis_min",
        "min_amount",
        "filter_below",
        "use_dedup",
        "amalgamate_duplicates",
        "credit_interest_rate",
        "credit_hold_days",
        "use_acc_filter",
        "acc_num",
        "scan_sap_dumps",
        "generate_reconciliation_sheet",
        "ofgem_caps_path",
        "ofgem_auto_carry",
    )
    state = object_value(prepared.extraction.legacy_state)
    sap = {
        key: state.get(key, [])
        for key in ("sap_contract_rows", "sap_meter_rows", "sap_financial_rows")
    }
    facts = {
        key: {
            field: metadata.get(field)
            for field in (
                "Content SHA256",
                "Content Hash",
                "Parent Source ID",
                "Page From",
                "Page To",
            )
            if field in metadata
        }
        for key, metadata in state.get("source_manifest", {}).items()
        if isinstance(metadata, dict)
    }
    ignored = {"Original Path", "Source Path", "File Path", "Copied Path", "Output Path"}
    records = [
        {key: value for key, value in row.items() if key not in ignored}
        for row in rows_value(prepared.records)
    ]
    duplicates = [
        {key: value for key, value in row.items() if key not in ignored}
        for row in rows_value(prepared.duplicates)
    ]
    payload: dict = {
        "records": records,
        "duplicates": duplicates,
        "effective_config": {key: config[key] for key in included if key in config},
        "relevant_legacy_state": {"sap": sap, "sources": facts},
        "methods": object_value(methods),
    }
    if state.get("bill_revision_decisions"):
        payload["relevant_legacy_state"]["bill_revision_decisions"] = state[
            "bill_revision_decisions"
        ]
    if state.get("dispute_schedule"):
        payload["relevant_legacy_state"]["dispute_schedule"] = state["dispute_schedule"]
    claims = state.get("claim_reconciliation")
    claim_context = claims.get("context") if isinstance(claims, dict) else None
    claim_context = claim_context if isinstance(claim_context, dict) else {}
    if isinstance(claims, dict) and (
        claims.get("groups")
        or claims.get("separations")
        or claim_context.get("complaint_summary")
        or claim_context.get("requested_outcome")
        or claim_context.get("include_pack") is False
    ):
        payload["relevant_legacy_state"]["claim_reconciliation"] = claims
    return hashlib.sha256(canonical_text(payload).encode()).hexdigest()


@dataclass(frozen=True)
class AnalysisSnapshot:
    """One computed result that can be preserved or rendered repeatedly."""

    prepared: PreparedSnapshot
    summary: FinancialSummary
    models: FrozenJSON
    methods: FrozenJSON
    input_fingerprint: str

    def __post_init__(self) -> None:
        """Validate transport shapes and the analytical fingerprint."""
        object_value(self.models)
        object_value(self.methods)
        if (
            not isinstance(self.input_fingerprint, str)
            or not re.fullmatch(r"[0-9a-f]{64}", self.input_fingerprint)
            or self.input_fingerprint != analytical_fingerprint(self.prepared, self.methods)
        ):
            raise ValueError("Analytical fingerprint does not match preserved inputs")
