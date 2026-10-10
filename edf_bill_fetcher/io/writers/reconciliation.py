"""Reconcile comparable, scoped observations with stable source sheet coordinates."""

from __future__ import annotations

from typing import Any

import pandas as pd
from openpyxl.styles import Font

from edf_bill_fetcher.helpers.date_utils import parse_to_sort_date
from edf_bill_fetcher.helpers.excel_utils import literal_cell
from edf_bill_fetcher.helpers.identity import account_key
from edf_bill_fetcher.processors.preparation import identity, normalize_frame, number
from edf_bill_fetcher.writers._helpers import _recon_hyperlink

RECON_ROW_OFFSET = 4


def _recon_parse_iso_date(value):
    """Use the same UK/ISO interpretation as ingestion and reporting."""
    return parse_to_sort_date(value)


def _recon_amount_to_float(value):
    """Return a finite numeric amount, keeping missing values unavailable."""
    return number(value)


def _compatible(left, right):
    """Reject conflicting ledger identities; unknown identity is not positive evidence."""
    for key in ("Account", "Meter", "Register", "Fuel"):
        a = identity(left.get(key))
        b = identity(right.get(key))
        if key == "Account":
            a = account_key(a or left.get("Kraken ID") or left.get("SAP Account Number"))
            b = account_key(b)
        if a and b and a.casefold() != b.casefold():
            return False
    return True


def write_reconciliation_sheet(
    ws_summary,
    ws_detail,
    sap_contract,
    inferred_contract,
    sap_meter,
    inferred_meter,
    sap_financial,
    evidence_df,
    account="",
    *, payload=None,
):
    """Match once per source row using date, comparable value and ledger identity."""
    ws_summary.title = "Reconciliation"
    ws_detail.title = "Reconciliation Drill-down"
    for ws in (ws_summary, ws_detail):
        literal_cell(ws, 1, 1, "EDF CROSS-SOURCE RECONCILIATION")
        literal_cell(ws, 2, 1, f"Comparable evidence observations only | Account: {account}")
        ws.freeze_panes = "A4"
    summary_headers = [
        "Entity",
        "SAP-side count",
        "EDF-inferred count",
        "Matched",
        "Unmatched (SAP)",
        "Unmatched (EDF)",
        "Verdict",
        "Drill down",
    ]
    detail_headers = [
        "Entity",
        "Status",
        "SAP Date",
        "SAP Value",
        "Evidence Date",
        "Evidence Value",
        "Notes",
        "Hyperlink",
    ]
    for ws, headers in ((ws_summary, summary_headers), (ws_detail, detail_headers)):
        for col, header_value in enumerate(headers, 1):
            literal_cell(ws, 3, col, header_value).font = Font(bold=True)
    if payload is not None:
        _render_stored(ws_summary, ws_detail, payload)
        return
    evidence = normalize_frame(evidence_df).to_dict("records")
    meter = []
    transactions = []
    for idx, record in enumerate(evidence):
        if number(record.get("Register Reading (kWh)")) is not None:
            meter.append(
                {**record, "_target": idx + 2, "_value": number(record["Register Reading (kWh)"])}
            )
        if number(record.get("Transaction Amount (£)")) is not None:
            transactions.append(
                {**record, "_target": idx + 2, "_value": number(record["Transaction Amount (£)"])}
            )
    contracts = []
    if inferred_contract is not None:
        for idx, record in enumerate(inferred_contract.to_dict("records")):
            # A display tariff name does not establish a SAP product code.
            if identity(record.get("Product Code")):
                contracts.append(
                    {
                        **record,
                        "Date": record.get("Contract From"),
                        "_value": identity(record.get("Product Code")),
                        "_target": idx + 8,
                    }
                )
    detail_row = 4
    specs = [
        (
            "Contract",
            sap_contract,
            contracts,
            "Contract From",
            "Product Code",
            "SAP Contract History",
            "Contract History",
        ),
        (
            "Meter Read",
            sap_meter,
            meter,
            "Meter Read Date",
            "Reading (kWh)",
            "SAP Meter Readings",
            "EDF Evidence Report",
        ),
        (
            "Financial",
            sap_financial,
            transactions,
            "Posting Date",
            "Amount",
            "SAP Financial Transactions",
            "EDF Evidence Report",
        ),
    ]
    for summary_row, (
        entity,
        sap_items,
        inferred,
        date_key,
        value_key,
        sap_sheet,
        inferred_sheet,
    ) in enumerate(specs, 4):
        start = detail_row
        remaining = list(range(len(inferred)))
        matched = unmatched_sap = 0
        for source_idx, sap in enumerate(sap_items):
            sap_date = parse_to_sort_date(sap.get(date_key))
            value: Any = (
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
                idx = min(matches)[1]
                remaining.remove(idx)
                matched += 1
                continue
            unmatched_sap += 1
            status = "Unmatched" if inferred else "Not assessed"
            note = (
                "No comparable inferred values available"
                if not inferred
                else "No match by date, value and ledger identity"
            )
            values = [entity, status, sap.get(date_key), value, None, None, note]
            for col, val in enumerate(values, 1):
                literal_cell(ws_detail, detail_row, col, val)
            _recon_hyperlink(ws_detail, detail_row, 8, sap_sheet, source_idx + 4)
            detail_row += 1
        for idx in remaining:
            record = inferred[idx]
            values = [
                entity,
                "Unmatched",
                None,
                None,
                record.get("Date"),
                record["_value"],
                "No comparable SAP match",
            ]
            for col, val in enumerate(values, 1):
                literal_cell(ws_detail, detail_row, col, val)
            _recon_hyperlink(ws_detail, detail_row, 8, inferred_sheet, record["_target"])
            detail_row += 1
        verdict = (
            "NOT ASSESSED"
            if not sap_items or not inferred
            else "MATCHED"
            if not unmatched_sap and not remaining
            else "REVIEW"
        )
        values = [
            entity,
            len(sap_items),
            len(inferred),
            matched,
            unmatched_sap,
            len(remaining),
            verdict,
        ]
        for col, summary_value in enumerate(values, 1):
            literal_cell(ws_summary, summary_row, col, summary_value)
        _recon_hyperlink(ws_summary, summary_row, 8, ws_detail.title, start)
    ws_summary.auto_filter.ref = "A3:H6"
    ws_detail.auto_filter.ref = f"A3:H{max(3, detail_row - 1)}"
    for ws in (ws_summary, ws_detail):
        for letter in "ABCDEFGH":
            ws.column_dimensions[letter].width = 24 if letter != "G" else 55


def _render_stored(ws_summary, ws_detail, payload):
    """Resolve stored source indexes into this workbook's layout coordinates."""
    sap_sheets = {"Contract": "SAP Contract History", "Meter Read": "SAP Meter Readings", "Financial": "SAP Financial Transactions"}
    for row, item in enumerate(payload["summaries"], 4):
        for col, value in enumerate(item["values"], 1):
            literal_cell(ws_summary, row, col, value)
        _recon_hyperlink(ws_summary, row, 8, ws_detail.title, item["detail_index"] + 4)
    for row, item in enumerate(payload["details"], 4):
        for col, value in enumerate(item["values"], 1):
            literal_cell(ws_detail, row, col, value)
        if item["side"] == "sap":
            sheet, offset = sap_sheets[item["entity"]], 4
        else:
            sheet, offset = ("Contract History", 8) if item["entity"] == "Contract" else ("EDF Evidence Report", 2)
        _recon_hyperlink(ws_detail, row, 8, sheet, item["index"] + offset)
    ws_summary.auto_filter.ref = "A3:H6"
    ws_detail.auto_filter.ref = f"A3:H{max(3, len(payload['details']) + 3)}"
    for ws in (ws_summary, ws_detail):
        for letter in "ABCDEFGH":
            ws.column_dimensions[letter].width = 24 if letter != "G" else 55
