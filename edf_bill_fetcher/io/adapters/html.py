"""Parse EDF MyAccount account-history exports into evidence observations."""

from __future__ import annotations

import re
from typing import Any

from edf_bill_fetcher.helpers.date_utils import parse_to_display_date
from edf_bill_fetcher.models.records import BillingRecord

__all__ = ["htm_excerpt", "parse_htm_account_history"]

_DATE = r"\d{1,2}\s+\w+\s+\d{4}"
_ROW_START_RE = re.compile(
    rf"\b({_DATE})\s+(We charged your account|You paid us|Reversed account charge|Balance)\b",
    re.IGNORECASE,
)
_MONEY_RE = re.compile(r"\s*£([\d,]+\.\d{2})")
_BALANCE_RE = re.compile(r"Balance\s+£([\d,]+\.\d{2})\s+in\s+(debit|credit)\b", re.I)
_CONSUMPTION_RE = re.compile(
    rf"For\s+([\d,]+)\s+kWh\s+of\s+electricity\s+used\s+between\s+({_DATE})\s+and\s+({_DATE})",
    re.IGNORECASE,
)
_PERIOD_DATES_RE = re.compile(rf"({_DATE})\s+(?:and|to|[-–])\s+({_DATE})", re.I)


def htm_excerpt(text: str, m: re.Match[str], window: int = 400) -> str:
    """Return a small window around a matched source observation."""
    start = max(0, m.start(0) - 20)
    end = min(len(text), m.end(0) + window)
    return text[start:end]


def parse_htm_account_history(text: str) -> list[dict[str, Any]]:
    """Read events in source order without borrowing a neighbouring balance.

    The document can omit a transaction's running balance. Preserve its stated
    transaction value with an unknown balance, rather than consuming the next
    event to fill that gap. Consumption dates inside an event are not starts.
    """
    text = re.sub(r"\s+", " ", text)
    period_dates = [match.span() for match in _PERIOD_DATES_RE.finditer(text)]
    starts = [
        match
        for match in _ROW_START_RE.finditer(text)
        if match.group(2).lower() != "balance"
        or not any(start <= match.start() < end for start, end in period_dates)
    ]
    records: list[dict[str, Any]] = []
    for index, match in enumerate(starts):
        end = starts[index + 1].start() if index + 1 < len(starts) else len(text)
        row_text = text[match.start() : end].strip()
        money_match = _MONEY_RE.match(text, match.end())
        if money_match is None or money_match.end() > end:
            continue
        amount = float(money_match.group(1).replace(",", ""))
        balance_match = _BALANCE_RE.search(row_text)
        balance: float | None = None
        if balance_match:
            balance = float(balance_match.group(1).replace(",", ""))
            if balance_match.group(2).lower() == "credit":
                balance = -balance

        verb = match.group(2).lower()
        period_from = period_to = units = "N/A"
        period_charge: float | None = amount
        if verb == "we charged your account":
            entry_type, details, logic, trace = (
                "New Bill",
                "HTM: charged account",
                "HTM Charge",
                "HTM charge_re",
            )
            consumption = _CONSUMPTION_RE.search(row_text)
            if consumption:
                units = consumption.group(1)
                period_from = parse_to_display_date(consumption.group(2))
                period_to = parse_to_display_date(consumption.group(3))
        elif verb == "you paid us":
            entry_type, details, logic, trace = (
                "Payment",
                "HTM: payment received",
                "HTM Payment",
                "HTM pay_re",
            )
        elif verb == "reversed account charge":
            entry_type, details, logic, trace = (
                "Credit",
                "HTM: reversed account charge",
                "HTM Reversal",
                "HTM rev_re",
            )
        else:
            # A balance observation is not a new credit transaction.
            # Retain the established display classification without inventing
            # a transaction amount from money already present in the account.
            if (
                balance is None
                or balance_match is None
                or balance_match.group(2).lower() != "credit"
            ):
                continue
            entry_type, details, logic, trace = (
                "Credit",
                "HTM: standalone credit balance",
                "HTM StandaloneBalance",
                "HTM bal_re (standalone)",
            )
            period_charge = None

        records.append(
            BillingRecord(
                source="HTM Account History",
                sender="",
                date=parse_to_display_date(match.group(1)),
                period_from=period_from,
                period_to=period_to,
                invoice_num="N/A",
                amount=balance,
                period_charge=period_charge,
                entry_type=entry_type,
                reading="N/A",
                units_kwh=units,
                standing_charge="N/A",
                tariff="N/A",
                attachment_name="N/A",
                details=details,
                logic_used=logic,
                source_pdf_text=row_text[:4000],
                regex_trace=trace,
            ).to_dict()
        )
    return records
