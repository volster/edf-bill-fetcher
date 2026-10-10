"""Compute shared financial results from one preserved preparation snapshot."""

from __future__ import annotations

import pandas as pd

from edf_bill_fetcher.domain.analysis import (
    AnalysisSnapshot,
    FinancialSummary,
    analytical_fingerprint,
)
from edf_bill_fetcher.domain.evidence import PreparedSnapshot
from edf_bill_fetcher.domain.json_values import freeze_json, object_value, rows_value
from edf_bill_fetcher.domain.money import parse_money
from edf_bill_fetcher.processors.financial import (
    balance_extremes,
    financial_totals,
    mean_daily_rate,
)
from edf_bill_fetcher.processors.preparation import analysis_frame, normalize_frame


def analyse_prepared(prepared: PreparedSnapshot) -> AnalysisSnapshot:
    """Compute a detached summary using unchanged financial-role rules."""
    frame = normalize_frame(pd.DataFrame(rows_value(prepared.records)))
    selected = analysis_frame(frame, object_value(prepared.effective_config))
    charges, credits = financial_totals(frame)
    opening, closing = balance_extremes(frame)
    charge_money, credit_money = parse_money(charges), parse_money(credits)
    if charge_money is None or credit_money is None:
        raise ValueError("Financial computation returned an unavailable total")
    summary = FinancialSummary(
        charge_money,
        credit_money,
        parse_money(opening),
        parse_money(closing),
        mean_daily_rate(selected),
    )
    methods = freeze_json(
        {"preparation": "legacy-v1", "financial": "legacy-v1", "models": "snapshot-v1"}
    )
    return AnalysisSnapshot(
        prepared, summary, freeze_json({}), methods, analytical_fingerprint(prepared, methods)
    )
