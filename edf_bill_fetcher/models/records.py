"""Typed ``BillingRecord`` for EDF billing evidence records.

``BillingRecord`` is the single typed producer of the 19-column evidence
schema shared by every record builder (``collectors/engine.py`` and
``io/adapters/html.py``). Snake-case fields default to ``None`` for unset
values; ``to_dict()`` is the display-key boundary, mapping ``None`` string
fields to ``"N/A"`` and ``cancel_rebill_admitted=None`` to ``False`` so the
sentinel never lives inside the dataclass.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any


def _na(value: str | float | None) -> str | float:
    return value if value is not None else "N/A"


@dataclass
class BillingRecord:
    """One row of billing evidence.

    Fields use snake_case names; ``to_dict()`` emits the 19 display keys.
    """

    source: str
    entry_type: str
    logic_used: str
    date: str | None = None
    period_from: str | None = None
    period_to: str | None = None
    invoice_num: str | None = None
    reading: str | None = None
    units_kwh: str | None = None
    standing_charge: str | None = None
    tariff: str | None = None
    sender: str | None = None
    attachment_name: str | None = None
    details: str | None = None
    source_pdf_text: str | None = None
    regex_trace: str | None = None
    amount: str | float | None = None
    period_charge: str | float | None = None
    cancel_rebill_admitted: bool | None = None

    def to_dict(self) -> dict[str, Any]:
        """Return the 19-key display dict, mapping None to "N/A" / False."""
        return {
            "Source": _na(self.source),
            "Sender": _na(self.sender),
            "Date": _na(self.date),
            "Period From": _na(self.period_from),
            "Period To": _na(self.period_to),
            "Invoice #": _na(self.invoice_num),
            "Amount (£)": _na(self.amount),
            "Period Charge (£)": _na(self.period_charge),
            "Entry Type": _na(self.entry_type),
            "Reading": _na(self.reading),
            "Units (kWh)": _na(self.units_kwh),
            "Standing Chg (p/day)": _na(self.standing_charge),
            "Tariff": _na(self.tariff),
            "Attachment Name": _na(self.attachment_name),
            "Details": _na(self.details),
            "Logic Used": _na(self.logic_used),
            "Source PDF Text": _na(self.source_pdf_text),
            "_regex_trace": _na(self.regex_trace),
            "Cancel/Rebill Admitted": (
                self.cancel_rebill_admitted if self.cancel_rebill_admitted is not None else False
            ),
        }
