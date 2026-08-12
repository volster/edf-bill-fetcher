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
from edf_bill_fetcher.models.records import BillingRecord

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
            BillingRecord(
                source="HTM Account History",
                sender="",
                date=date_str,
                period_from=period_from,
                period_to=period_to,
                invoice_num="N/A",
                amount=balance,
                period_charge=charge_amt,
                entry_type="Ongoing Balance",
                reading="N/A",
                units_kwh=units,
                standing_charge="N/A",
                tariff="N/A",
                attachment_name="N/A",
                details="HTM: charged account",
                logic_used="HTM Charge",
                source_pdf_text=excerpt,
                regex_trace="HTM charge_re",
            ).to_dict()
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
            BillingRecord(
                source="HTM Account History",
                sender="",
                date=date_str,
                period_from="N/A",
                period_to="N/A",
                invoice_num="N/A",
                amount=balance,
                period_charge=payment_amt,
                entry_type="Payment",
                reading="N/A",
                units_kwh="N/A",
                standing_charge="N/A",
                tariff="N/A",
                attachment_name="N/A",
                details="HTM: payment received",
                logic_used="HTM Payment",
                source_pdf_text=excerpt,
                regex_trace="HTM pay_re",
            ).to_dict()
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
            BillingRecord(
                source="HTM Account History",
                sender="",
                date=date_str,
                period_from="N/A",
                period_to="N/A",
                invoice_num="N/A",
                amount=balance,
                period_charge=credit_amt,
                entry_type="Credit",
                reading="N/A",
                units_kwh="N/A",
                standing_charge="N/A",
                tariff="N/A",
                attachment_name="N/A",
                details="HTM: reversed account charge",
                logic_used="HTM Reversal",
                source_pdf_text=excerpt,
                regex_trace="HTM rev_re",
            ).to_dict()
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
            BillingRecord(
                source="HTM Account History",
                sender="",
                date=date_str,
                period_from="N/A",
                period_to="N/A",
                invoice_num="N/A",
                amount=balance,
                period_charge="N/A",
                entry_type="Credit",
                reading="N/A",
                units_kwh="N/A",
                standing_charge="N/A",
                tariff="N/A",
                attachment_name="N/A",
                details="HTM: standalone credit balance",
                logic_used="HTM StandaloneBalance",
                source_pdf_text=htm_excerpt(text, m),
                regex_trace="HTM bal_re (standalone)",
            ).to_dict()
        )

    return records
