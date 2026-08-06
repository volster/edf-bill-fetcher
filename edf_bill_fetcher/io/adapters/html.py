"""HTML reading adapters — BeautifulSoup-backed HTM account-history parser used by ``EvidenceEngine.process_htm_file``.

This module owns the *file-reading primitive* of the HTM path: the
regex-driven parser that turns the EDF MyAccount ``Payments and
Invoices`` HTML export into the same record-dict shape as the PDF /
PST paths.  ``beautifulsoup4`` is the only mandatory dep (used by
the engine to strip tags before invoking
:func:`parse_htm_account_history`).
"""

from __future__ import annotations

import re
from typing import Any

from edf_bill_fetcher.helpers.date_utils import parse_to_display_date

__all__ = [
    "htm_excerpt",
    "parse_htm_account_history",
]


def htm_excerpt(text: str, m: re.Match[str], window: int = 400) -> str:
    """Return a small window of the HTM source around a regex match.

    Used by :func:`parse_htm_account_history` to populate the
    ``Source PDF Text`` column captured for the analyser tabs' Source
    Excerpt lookup.  Capturing the entire HTM document per-record
    would balloon memory (every record would carry the same ~5-50 KB
    body); a 400-char window around each match is enough for a
    reviewer to see the verb phrase + balance clause that produced
    the row.
    """
    start = max(0, m.start(0) - 20)
    end = min(len(text), m.end(0) + window)
    return text[start:end]


def parse_htm_account_history(text: str) -> list[dict[str, Any]]:
    """Parse the EDF MyAccount 'Payments and Invoices' HTM export.

    Returns a list of record dicts ready for ``process_text`` bypass.
    """
    records: list[dict[str, Any]] = []

    text = re.sub(r"\s+", " ", text)

    charge_re = re.compile(
        r"(\d{1,2}\s+\w+\s+\d{4})\s+We charged your account\s+£([\d,]+\.\d{2})"
        r"(?:\s+For\s+([\d,]+)\s+kWh\s+of\s+electricity\s+used\s+between\s+"
        r"(\d{1,2}\s+\w+\s+\d{4})\s+and\s+(\d{1,2}\s+\w+\s+\d{4}))?"
        r".*?Balance\s+£([\d,]+\.\d{2})\s+in\s+(?:debit|credit)",
        re.IGNORECASE,
    )
    covered: list[tuple[int, int]] = []
    for m in charge_re.finditer(text):
        covered.append((m.start(0), m.end(0)))
        date_str = parse_to_display_date(m.group(1))
        period_from = parse_to_display_date(m.group(4)) if m.group(4) else "N/A"
        period_to = parse_to_display_date(m.group(5)) if m.group(5) else "N/A"
        units = m.group(3) if m.group(3) else "N/A"
        charge_amt = float(m.group(2).replace(",", ""))
        balance = float(m.group(6).replace(",", ""))
        excerpt = htm_excerpt(text, m)
        records.append(
            {
                "Source": "HTM Account History",
                "Sender": "",
                "Date": date_str,
                "Period From": period_from,
                "Period To": period_to,
                "Invoice #": "N/A",
                "Amount (£)": balance,
                "Period Charge (£)": charge_amt,
                "Entry Type": "Ongoing Balance",
                "Reading": "N/A",
                "Units (kWh)": units,
                "Standing Chg (p/day)": "N/A",
                "Tariff": "N/A",
                "Attachment Name": "N/A",
                "Details": "HTM: charged account",
                "Logic Used": "HTM Charge",
                "Source PDF Text": excerpt,
                "_regex_trace": "HTM charge_re",
            }
        )

    pay_re = re.compile(
        r"(\d{1,2}\s+\w+\s+\d{4})\s+You paid us\s+£([\d,]+\.\d{2})"
        r".*?Balance\s+£([\d,]+\.\d{2})\s+in\s+(?:debit|credit)",
        re.IGNORECASE,
    )
    for m in pay_re.finditer(text):
        covered.append((m.start(0), m.end(0)))
        date_str = parse_to_display_date(m.group(1))
        payment_amt = float(m.group(2).replace(",", ""))
        balance = float(m.group(3).replace(",", ""))
        excerpt = htm_excerpt(text, m)
        records.append(
            {
                "Source": "HTM Account History",
                "Sender": "",
                "Date": date_str,
                "Period From": "N/A",
                "Period To": "N/A",
                "Invoice #": "N/A",
                "Amount (£)": balance,
                "Period Charge (£)": payment_amt,
                "Entry Type": "Payment",
                "Reading": "N/A",
                "Units (kWh)": "N/A",
                "Standing Chg (p/day)": "N/A",
                "Tariff": "N/A",
                "Attachment Name": "N/A",
                "Details": "HTM: payment received",
                "Logic Used": "HTM Payment",
                "Source PDF Text": excerpt,
                "_regex_trace": "HTM pay_re",
            }
        )

    rev_re = re.compile(
        r"(\d{1,2}\s+\w+\s+\d{4})\s+Reversed account charge\s+£([\d,]+\.\d{2})"
        r".*?Balance\s+£([\d,]+\.\d{2})\s+in\s+(?:debit|credit)",
        re.IGNORECASE,
    )
    for m in rev_re.finditer(text):
        covered.append((m.start(0), m.end(0)))
        date_str = parse_to_display_date(m.group(1))
        credit_amt = float(m.group(2).replace(",", ""))
        balance = float(m.group(3).replace(",", ""))
        excerpt = htm_excerpt(text, m)
        records.append(
            {
                "Source": "HTM Account History",
                "Sender": "",
                "Date": date_str,
                "Period From": "N/A",
                "Period To": "N/A",
                "Invoice #": "N/A",
                "Amount (£)": balance,
                "Period Charge (£)": credit_amt,
                "Entry Type": "Credit",
                "Reading": "N/A",
                "Units (kWh)": "N/A",
                "Standing Chg (p/day)": "N/A",
                "Tariff": "N/A",
                "Attachment Name": "N/A",
                "Details": "HTM: reversed account charge",
                "Logic Used": "HTM Reversal",
                "Source PDF Text": excerpt,
                "_regex_trace": "HTM rev_re",
            }
        )

    def _inside_covered(start: int, end: int) -> bool:
        for cs, ce in covered:
            if start >= cs and end <= ce:
                return True
        return False

    bal_re = re.compile(
        r"(\d{1,2}\s+\w+\s+\d{4})[^A-Za-z0-9]*?Balance\s+£([\d,]+\.\d{2})\s+in\s+credit\b",
        re.IGNORECASE,
    )
    for m in bal_re.finditer(text):
        if _inside_covered(m.start(0), m.end(0)):
            continue
        covered.append((m.start(0), m.end(0)))
        date_str = parse_to_display_date(m.group(1))
        balance = float(m.group(2).replace(",", ""))
        records.append(
            {
                "Source": "HTM Account History",
                "Sender": "",
                "Date": date_str,
                "Period From": "N/A",
                "Period To": "N/A",
                "Invoice #": "N/A",
                "Amount (£)": balance,
                "Period Charge (£)": "N/A",
                "Entry Type": "Credit",
                "Reading": "N/A",
                "Units (kWh)": "N/A",
                "Standing Chg (p/day)": "N/A",
                "Tariff": "N/A",
                "Attachment Name": "N/A",
                "Details": "HTM: standalone credit balance",
                "Logic Used": "HTM StandaloneBalance",
                "Source PDF Text": htm_excerpt(text, m),
                "_regex_trace": "HTM bal_re (standalone)",
            }
        )

    return records
