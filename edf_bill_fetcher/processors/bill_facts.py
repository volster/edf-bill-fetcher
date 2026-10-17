"""Context-bound bill facts, separate from net ledger movements."""

from __future__ import annotations

import re

from edf_bill_fetcher.helpers.date_utils import parse_to_display_date
from edf_bill_fetcher.processors.patterns import SUB_PERIOD_RE

_DATE = r"\d{1,2}\s+[A-Za-z.]+\s+\d{4}"


def bill_facts(text: str, sub_periods: list[dict]) -> dict:
    """Retain gross replacement charges and useful consumption context.

    A cancellation/rebill can have a negative net total and positive new
    electricity components. These amounts must never be interchangeable.
    """
    facts: dict = {}
    gross = re.search(
        r"Total electricity charges for this period\s+£([\d,]+\.\d{2})(?:\s+(?:in\s+)?(credit|debit)\b)?",
        text,
        re.I,
    )
    if gross:
        facts["Gross Electricity Charges (£)"] = float(gross[1].replace(",", ""))
        if (gross[2] or "").lower() == "credit":
            facts["Gross Electricity Charges (£)"] = -facts["Gross Electricity Charges (£)"]
    usage = re.search(r"Your usage this bill\s+([\d,]+(?:\.\d+)?)\s*kWh", text, re.I)
    if usage:
        facts["Units (kWh)"] = float(usage[1].replace(",", ""))
    elif sub_periods:
        facts["Units (kWh)"] = sum(p["units_kwh"] for p in sub_periods)
    bases = set()
    for label, pattern in (
        ("Estimated", r"\bEstimated(?: reading)?\b"),
        ("Actual", r"\b(?:OUR READ|Actual reading)\b"),
        ("Customer", r"\b(?:YOUR READ|Customer reading)\b"),
    ):
        if re.search(pattern, text, re.I):
            bases.add(label)
    if len(bases) > 1:
        facts["Reading"] = "Mixed"
    readings = list(
        re.finditer(
            rf"\b({_DATE})\s+([\d,]+(?:\.\d+)?)\s+(Estimated|Actual|Customer) reading", text, re.I
        )
    )
    if readings:
        last = readings[-1]
        facts.update(
            {
                "Reading Date": parse_to_display_date(last[1]),
                "Register Reading (kWh)": float(last[2].replace(",", "")),
                "Closing Reading Type": last[3].title(),
            }
        )
    else:
        periods = list(SUB_PERIOD_RE.finditer(re.sub(r"\s+", " ", text)))
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
    block = re.split(r"Frequently asked questions|About your", text[header.end() :], flags=re.I)[0]
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
