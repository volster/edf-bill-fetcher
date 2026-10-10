"""Compute shared results once, preserving legacy financial and matching rules."""

from __future__ import annotations

import hashlib
from importlib import resources

import pandas as pd

from edf_bill_fetcher.domain.analysis import (
    AnalysisSnapshot,
    FinancialSummary,
    analytical_fingerprint,
)
from edf_bill_fetcher.domain.evidence import PreparedSnapshot
from edf_bill_fetcher.domain.json_values import freeze_json, object_value, rows_value
from edf_bill_fetcher.domain.money import parse_money
from edf_bill_fetcher.helpers.date_utils import parse_to_sort_date
from edf_bill_fetcher.helpers.identity import account_key
from edf_bill_fetcher.io.writers.analysis import run_analysers
from edf_bill_fetcher.models import report_models as models
from edf_bill_fetcher.models.snapshot_adapter import to_snapshot_json
from edf_bill_fetcher.processors.compensation import estimate_compensation
from edf_bill_fetcher.processors.detection import (
    compute_transitive_domination,
    compute_unlawful_union_total,
)
from edf_bill_fetcher.processors.financial import (
    balance_extremes,
    bill_daily_rates,
    financial_totals,
    mean_daily_rate,
)
from edf_bill_fetcher.processors.matching import (
    _build_bb_clusters,
    analyse_sap_back_billing,
    match_sap_events_to_edf,
)
from edf_bill_fetcher.processors.preparation import (
    analysis_frame,
    balance_frame,
    identity,
    monthly_balance_frame,
    normalize_frame,
    number,
    scope,
)
from edf_bill_fetcher.writers._helpers import (
    detect_sap_back_billing_events,
    handle_cluster_unmatched,
)


def _compatible(left: dict, right: dict) -> bool:
    for key in ("Account", "Meter", "Register", "Fuel"):
        a, b = identity(left.get(key)), identity(right.get(key))
        if key == "Account":
            a = account_key(a or left.get("Kraken ID") or left.get("SAP Account Number"))
            b = account_key(b)
        if a and b and a.casefold() != b.casefold():
            return False
    return True


def _reconciliation(state: dict, frame: pd.DataFrame, contracts: pd.DataFrame) -> dict:
    """Port the legacy one-to-one matching rule, storing source indexes, not cells."""
    evidence = frame.to_dict("records")
    meter: list[dict] = []
    transactions: list[dict] = []
    inferred_contract: list[dict] = []
    for idx, record in enumerate(evidence):
        for key, target in (
            ("Register Reading (kWh)", meter),
            ("Transaction Amount (£)", transactions),
        ):
            value = number(record.get(key))
            if value is not None:
                target.append({**record, "_index": idx, "_value": value})
    for idx, record in enumerate(contracts.to_dict("records")):
        if identity(record.get("Product Code")):
            inferred_contract.append(
                {
                    **record,
                    "Date": record.get("Contract From"),
                    "_index": idx,
                    "_value": identity(record.get("Product Code")),
                }
            )
    specs = (
        (
            "Contract",
            state.get("sap_contract_rows", []),
            inferred_contract,
            "Contract From",
            "Product Code",
        ),
        ("Meter Read", state.get("sap_meter_rows", []), meter, "Meter Read Date", "Reading (kWh)"),
        ("Financial", state.get("sap_financial_rows", []), transactions, "Posting Date", "Amount"),
    )
    summaries: list[dict] = []
    details: list[dict] = []
    for entity, sap_items, inferred, date_key, value_key in specs:
        remaining = list(range(len(inferred)))
        matched = unmatched_sap = 0
        start = len(details)
        for source_idx, sap in enumerate(sap_items):
            sap_date = parse_to_sort_date(sap.get(date_key))
            value = (
                identity(sap.get(value_key)) if entity == "Contract" else number(sap.get(value_key))
            )
            if entity == "Financial" and value is not None:
                kind = str(sap.get("Document Type", "")).lower()
                if "payment" in kind or "credit" in kind:
                    value = -abs(value)
            matches = []
            for idx in remaining:
                record = inferred[idx]
                date = parse_to_sort_date(record.get("Date"))
                if pd.isna(sap_date) or pd.isna(date) or not _compatible(sap, record):
                    continue
                delta = abs((sap_date - date).days)
                equal = (
                    value == record["_value"]
                    if entity == "Contract"
                    else value is not None and abs(value - record["_value"]) <= 0.005
                )
                if delta <= 7 and equal:
                    matches.append((delta, idx))
            if matches:
                remaining.remove(min(matches)[1])
                matched += 1
                continue
            unmatched_sap += 1
            details.append(
                {
                    "entity": entity,
                    "side": "sap",
                    "index": source_idx,
                    "values": [
                        entity,
                        "Unmatched" if inferred else "Not assessed",
                        sap.get(date_key),
                        value,
                        None,
                        None,
                        "No match by date, value and ledger identity"
                        if inferred
                        else "No comparable inferred values available",
                    ],
                }
            )
        for idx in remaining:
            record = inferred[idx]
            details.append(
                {
                    "entity": entity,
                    "side": "edf",
                    "index": record["_index"],
                    "values": [
                        entity,
                        "Unmatched",
                        None,
                        None,
                        record.get("Date"),
                        record["_value"],
                        "No comparable SAP match",
                    ],
                }
            )
        verdict = (
            "NOT ASSESSED"
            if not sap_items or not inferred
            else "MATCHED"
            if not unmatched_sap and not remaining
            else "REVIEW"
        )
        summaries.append(
            {
                "values": [
                    entity,
                    len(sap_items),
                    len(inferred),
                    matched,
                    unmatched_sap,
                    len(remaining),
                    verdict,
                ],
                "detail_index": start,
            }
        )
    return {"summaries": summaries, "details": details}


def _computed_models(
    prepared: PreparedSnapshot, frame: pd.DataFrame, selected: pd.DataFrame, mean: float
) -> dict:
    cfg, state = (
        object_value(prepared.effective_config),
        object_value(prepared.extraction.legacy_state),
    )
    payment = models.compute_payment_analysis(frame)
    quality = models.compute_data_quality_report(
        frame, duplicate_count=len(rows_value(prepared.duplicates))
    )
    quality.duplicate_count = len(rows_value(prepared.duplicates))
    quality.duplicate_rate = (
        quality.duplicate_count / quality.total_records if quality.total_records else 0.0
    )
    statistical = models.compute_statistical_analysis(frame)
    forecast = models.compute_forecast(frame)
    ofgem = models.compute_ofgem_comparison(frame, cfg)
    tariff = models.compute_tariff_analysis(frame)
    dispute = models.compute_dispute_analysis(selected, mean)
    analyses = run_analysers(selected, evidence_index_df=frame)
    domination = compute_transitive_domination(analyses["rebilling"], analyses["back_billing"])
    events = detect_sap_back_billing_events(state.get("sap_financial_rows", []))
    matches = match_sap_events_to_edf(events, selected.to_dict("records"))
    clusters = _build_bb_clusters(analyses["back_billing"])
    for event in events:
        candidates = [match for match in matches if match.event is event]
        if candidates:
            best = min(
                candidates,
                key=lambda m: (
                    {"High": 0, "Medium": 1, "Low": 2}.get(m.confidence_band, 3),
                    m.date_delta_days,
                ),
            )
            event.matched_edf_invoice = str(best.edf_record.get("Invoice #") or "") or None
        else:
            tag = handle_cluster_unmatched(event, clusters)
            if tag is not None:
                event.matched_edf_invoice = tag["Matched EDF Invoice #"]
                event._cluster_unmatched_tag = tag
    position = analyse_sap_back_billing(events, selected, analyses["back_billing"])
    balance = balance_frame(selected)
    statistical_frame = balance_frame(frame)
    if len({scope(row) for row in statistical_frame.to_dict("records")}) > 1:
        statistical_frame = statistical_frame.iloc[:0]
    balance["year"] = balance["_dt"].dt.year
    yearly = (
        balance.groupby("year")
        .agg(
            count=("Amount (£)", "count"),
            avg_bal=("Amount (£)", "mean"),
            peak=("Amount (£)", "max"),
            low=("Amount (£)", "min"),
        )
        .reset_index()
    )
    live = analyses["back_billing"]
    if not live.empty:
        live = live[~live["Invoice #"].astype(str).isin(domination)]
    source_indexes = {id(row): idx for idx, row in enumerate(state.get("sap_financial_rows", []))}
    sap = {
        "analyses": analyses,
        "domination": domination,
        "events": events,
        "matches": matches,
        "position": position,
        "event_source_indexes": [
            [source_indexes[id(row)] for row in event.rows] for event in events
        ],
        "union_total": compute_unlawful_union_total(live),
        "frames": {
            "evidence": frame,
            "analysis": selected,
            "balance": balance,
            "statistical": statistical_frame,
            "monthly": monthly_balance_frame(frame),
            "daily_rates": bill_daily_rates(selected),
            "yearly": yearly,
        },
        "timeline": _timeline(balance, dispute.flags),
    }
    results = {
        "payment": (payment, True, None),
        "data_quality": (quality, True, None),
        "statistical": (statistical, statistical.count > 0, "NO_COMPARABLE_BALANCES"),
        "forecast": (forecast, forecast.n >= 3, "INSUFFICIENT_CONTIGUOUS_MONTHLY_BALANCES"),
        "ofgem": (
            ofgem,
            bool(ofgem.rows) and not ofgem.unavailable_count,
            "MISSING_BENCHMARK_CONTEXT",
        ),
        "tariff": (tariff, not tariff.empty, "NO_TARIFF_EVIDENCE"),
        "dispute": (dispute, True, None),
        "compensation": (estimate_compensation(selected, cfg), True, None),
        "sap": (sap, bool(state.get("sap_financial_rows")), "NO_SAP_FINANCIAL_EVIDENCE"),
        "reconciliation": (
            _reconciliation(state, frame, analyses["contracts"]),
            bool(
                any(
                    state.get(key)
                    for key in ("sap_contract_rows", "sap_meter_rows", "sap_financial_rows")
                )
            ),
            "NO_SAP_EVIDENCE",
        ),
    }
    diagnostics: list[dict] = []
    converted = {
        key: {
            "availability": "available" if available else "unavailable",
            "reason_code": None if available else reason,
            "payload": to_snapshot_json(payload, path=f"$.models.{key}", diagnostics=diagnostics),
        }
        for key, (payload, available, reason) in results.items()
    }
    for key, entry in converted.items():
        if any(item["path"].startswith(f"$.models.{key}.") for item in diagnostics):
            entry["availability"] = "partial"
            entry["reason_code"] = "NONFINITE_METRICS_UNAVAILABLE"
    return converted


def analyse_prepared(prepared: PreparedSnapshot) -> AnalysisSnapshot:
    """Calculate the financial summary and every analytical model exactly once."""
    frame = normalize_frame(pd.DataFrame(rows_value(prepared.records)))
    cfg = object_value(prepared.effective_config)
    selected = analysis_frame(frame, cfg)
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
    reference = resources.files("edf_bill_fetcher.data").joinpath("ofgem_caps.json").read_bytes()
    methods = freeze_json(
        {
            "preparation": "legacy-v1",
            "financial": "legacy-v1",
            "models": "snapshot-v1",
            "reference_sha256": hashlib.sha256(reference).hexdigest(),
            "reference_context": {
                key: cfg.get(key) for key in ("fuel", "payment_method", "tariff_type", "cap_region")
            },
        }
    )
    computed = freeze_json(_computed_models(prepared, frame, selected, summary.mean_daily_charge))
    return AnalysisSnapshot(
        prepared, summary, computed, methods, analytical_fingerprint(prepared, methods)
    )


def _timeline(dfc: pd.DataFrame, flags: list) -> list:
    """Preserve the legacy workbook timeline inference in the analysis boundary."""
    if len(dfc) < 2 or len({scope(row) for row in dfc.to_dict("records")}) != 1:
        return []
    dates_lbl = dfc["Date"].tolist()
    amounts = dfc["Amount (£)"].values.astype(float)
    n = len(amounts)
    timeline_events = []

    # Bookend: first record
    timeline_events.append(
        (dates_lbl[0], "ACCOUNT START", f"First bill on record. Balance: £{amounts[0]:,.2f}.")
    )

    # Top 5 largest balance jumps
    jumps = []
    for i in range(1, n):
        delta = float(amounts[i]) - float(amounts[i - 1])
        days = (dfc.iloc[i]["_dt"] - dfc.iloc[i - 1]["_dt"]).days
        if delta > 0:
            jumps.append((delta, i, days))
    jumps.sort(key=lambda x: x[0], reverse=True)
    for delta, idx, days in jumps[:5]:
        timeline_events.append(
            (
                dfc.iloc[idx]["Date"],
                "LARGE INCREASE",
                f"Balance rose £{delta:,.2f} in {days} days "
                f"(from £{amounts[idx - 1]:,.2f} to £{amounts[idx]:,.2f}).",
            )
        )

    # Billing gaps > 60 days
    for i in range(1, n):
        days = (dfc.iloc[i]["_dt"] - dfc.iloc[i - 1]["_dt"]).days
        if days > 60:
            timeline_events.append(
                (
                    dfc.iloc[i]["Date"],
                    "BILLING GAP",
                    f"{days} days without a bill (previous: {dfc.iloc[i - 1]['Date']}). "
                    f"Balance accumulated unchecked.",
                )
            )

    # Estimated reading runs (reuse existing detection)
    if "Reading" in dfc.columns:
        run = 0
        run_start_date = None
        for i, rv in enumerate(dfc["Reading"].tolist()):
            if str(rv).lower() in ("estimated", "est."):
                run += 1
                if run == 1:
                    run_start_date = dfc.iloc[i]["Date"]
            else:
                if run >= 3:
                    timeline_events.append(
                        (
                            run_start_date,
                            "ESTIMATED READINGS",
                            f"{run} consecutive bills used estimated meter readings.",
                        )
                    )
                run = 0
                run_start_date = None
        if run >= 3:
            timeline_events.append(
                (
                    run_start_date,
                    "ESTIMATED READINGS",
                    f"{run} consecutive estimated readings (ongoing).",
                )
            )

    # Payment events (balance reductions)
    for i in range(1, n):
        delta = float(amounts[i]) - float(amounts[i - 1])
        if delta < -200:
            timeline_events.append(
                (
                    dfc.iloc[i]["Date"],
                    "PAYMENT/CREDIT",
                    f"Balance reduced by £{abs(delta):,.2f} "
                    f"(from £{amounts[i - 1]:,.2f} to £{amounts[i]:,.2f}).",
                )
            )

    # Reconciliation mismatches (from flags)
    for ftype, fdate, _famt, fdetail, _fsev in flags:
        if ftype == "RECONCILIATION MISMATCH":
            timeline_events.append((fdate, "RECONCILIATION", fdetail))

    # Bookend: latest record
    timeline_events.append(
        (
            dates_lbl[-1],
            "CURRENT STATE",
            f"Latest bill on record. Balance: £{amounts[-1]:,.2f}. "
            f"Total increase from first record: £{amounts[-1] - amounts[0]:,.2f}.",
        )
    )

    return timeline_events
