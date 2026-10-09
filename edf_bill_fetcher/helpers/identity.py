"""Document identity extraction and exact account matching."""

from __future__ import annotations

import re


def account_key(value: object) -> str:
    """Normalize EDF account renderings without substring matching."""
    return re.sub(r"[^A-Z0-9]", "", str(value).upper()).removeprefix("A")


def document_identity(text: str) -> dict[str, str]:
    """Retain identity and benchmark context actually present in a document."""
    account = re.search(r"Account\s+(?:number|reference)\s*:?\s*(A-\d+|\d[\d ]*\d)", text, re.I)
    meter = re.search(r"Meter\s+(?:serial\s+)?number\s*:?\s*([A-Z0-9-]+)", text, re.I)
    fuel = (
        "electricity"
        if re.search(r"\bElectricity\b", text, re.I)
        else "gas"
        if re.search(r"\bGas\b", text, re.I)
        else ""
    )
    payment = re.search(
        r"(?:Payment\s+(?:type|method))\s*:?\s*(Direct\s+Debit|Cash|Cheque|Prepayment)", text, re.I
    )
    tariff_type = (
        "variable" if re.search(r"\b(?:standard variable|default tariff)\b", text, re.I) else ""
    )
    return {
        "Account": account.group(1).strip() if account else "",
        "Meter": meter.group(1) if meter else "",
        "Fuel": fuel,
        "Payment Method": payment.group(1) if payment else "",
        "Tariff Type": tariff_type,
    }


def account_matches(selected: str, row: dict) -> bool:
    """Compare known account identities, including SAP's alternate identifier."""
    key = account_key(selected)
    return bool(key) and any(
        account_key(row.get(field, "")) == key
        for field in ("Account", "Kraken ID", "SAP Account Number")
    )
