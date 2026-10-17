"""Context-bound bill facts, separate from net ledger movements."""

from __future__ import annotations

import re
from datetime import datetime

from edf_bill_fetcher.helpers.date_utils import parse_to_display_date
from edf_bill_fetcher.processors.patterns import SUB_PERIOD_RE, normalize_charge_table

_DATE = r"\d{1,2}\s+[A-Za-z.]+\s+\d{4}"


def bill_facts(text: str, sub_periods: list[dict]) -> dict:
    """Retain gross replacement charges and useful consumption context.

    A cancellation/rebill can have a negative net total and positive new
    electricity components. These amounts must never be interchangeable.
    """
    facts: dict = {}
    gross = list(
        re.finditer(
            r"Total electricity charges for this period\s+£([\d,]+\.\d{2})(?:\s+(?:in\s+)?(credit|debit)\b)?",
            text,
            re.I,
        )
    )
    if gross:
        facts["Gross Electricity Charges (£)"] = round(
            sum(
                float(match[1].replace(",", ""))
                * (-1 if (match[2] or "").lower() == "credit" else 1)
                for match in gross
            ),
            2,
        )
    usage = re.search(r"Your usage this bill\s+([\d,]+(?:\.\d+)?)\s*kWh", text, re.I)
    if usage:
        facts["Units (kWh)"] = float(usage[1].replace(",", ""))
    elif used := list(
        re.finditer(
            r"Electricity used\s+([\d,]+(?:\.\d+)?)\s+kWh\s*@|([\d,]+(?:\.\d+)?)\s+kWh\s*@\s+Electricity used",
            text,
            re.I,
        )
    ):
        facts["Units (kWh)"] = round(sum(float((m[1] or m[2]).replace(",", "")) for m in used), 6)
    elif sub_periods:
        facts["Units (kWh)"] = sum(p["units_kwh"] for p in sub_periods)
    tariffs = []
    for match in re.finditer(r"\bMeter:\s*\S+\s+Tariff:\s*([^\n]+)", text, re.I):
        tariff = re.split(r"\s+Product\b", match[1], flags=re.I)[0].strip()
        if tariff and tariff not in tariffs:
            tariffs.append(tariff)
    if tariffs:
        facts["Tariff"] = "; ".join(tariffs)
    bases = set()
    normalized = normalize_charge_table(text)
    for label, pattern in (
        ("Estimated", r"\d\s*ESTIMATED\b"),
        ("Actual", r"\d\s*(?:OUR READ|Actual reading|Data collector reading)\b"),
        ("Customer", r"\d\s*(?:YOUR READ|Customer reading)\b"),
    ):
        if re.search(pattern, normalized, re.I):
            bases.add(label)
    if bases:
        facts["Reading"] = "Mixed" if len(bases) > 1 else next(iter(bases))
    meters = {
        m.upper()
        for m in re.findall(r"(?:charges for meter\s+|\bMeter:\s*)([A-Z0-9-]+)", text, re.I)
    }
    if len(meters) > 1:
        return facts
    readings = list(
        re.finditer(
            rf"\b({_DATE})\s+([\d,]+(?:\.\d+)?)\s+(Estimated|Actual|Customer|Data collector) reading",
            text,
            re.I,
        )
    )
    valid_readings = []
    for match in readings:
        try:
            valid_readings.append(
                (datetime.strptime(parse_to_display_date(match[1]), "%d/%m/%Y"), match)
            )
        except ValueError:
            continue
    if valid_readings:
        # Sections can appear in reverse chronological order. The physical
        # last row is not necessarily the closing reading of the invoice.
        latest_date, last = max(valid_readings, key=lambda item: item[0])
        latest = [match for when, match in valid_readings if when == latest_date]
        if len({(float(m[2].replace(",", "")), m[3].lower()) for m in latest}) > 1:
            return facts
        facts.update(
            {
                "Reading Date": parse_to_display_date(last[1]),
                "Register Reading (kWh)": float(last[2].replace(",", "")),
                "Closing Reading Type": "Actual"
                if last[3].lower() == "data collector"
                else last[3].title(),
            }
        )
    else:
        periods = list(SUB_PERIOD_RE.finditer(normalized))
        if periods:
            last_period = periods[-1]
            values = list(
                re.finditer(
                    r"([\d,]+(?:\.\d+)?)\s*(OUR READ|YOUR READ|ESTIMATED)", last_period["mid"], re.I
                )
            )
            if values:
                last_read = values[-1]
                facts.update(
                    {
                        "Reading Date": parse_to_display_date(last_period["pt"]),
                        "Register Reading (kWh)": float(last_read[1].replace(",", "")),
                        "Closing Reading Type": {
                            "OUR READ": "Actual",
                            "YOUR READ": "Customer",
                            "ESTIMATED": "Estimated",
                        }[last_read[2].upper()],
                    }
                )
    return facts


def payment_history(text: str) -> list[dict]:
    """Read dated history entries without treating a cover total as a payment."""
    header = re.search(r"Your payment history", text, re.I)
    if not header:
        return []
    block = re.split(
        r"Frequently asked questions|About your charges|About your tariff",
        text[header.end() :],
        flags=re.I,
    )[0]
    rows: list[dict] = []
    for match in re.finditer(rf"\b({_DATE})\s+£([\d,]+\.\d{{2}})", block, re.I):
        amount = float(match[2].replace(",", ""))
        if amount:
            rows.append(
                {
                    "Date": parse_to_display_date(match[1]),
                    "Amount (£)": -amount,
                    "Transaction Amount (£)": -amount,
                    "Amount Kind": "Transaction",
                    "Entry Type": "Payment",
                    "Invoice #": "N/A",
                    "Details": "Dated payment history on invoice; cover aggregate excluded",
                    "Logic Used": "Invoice Payment History",
                    "Source Row": f"payment-history:{len(rows) + 1}",
                    "Source PDF Text": match[0],
                    "Historical Activity": True,
                }
            )
    return rows
