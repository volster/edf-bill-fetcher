"""Document identity extraction and exact account matching."""

from __future__ import annotations

import re


def account_key(value: object) -> str:
    """Normalize EDF account renderings without substring matching."""
    return re.sub(r"[^A-Z0-9]", "", str(value).upper()).removeprefix("A")


def document_identity(text: str) -> dict[str, str]:
    """Retain identity and benchmark context actually present in a document."""
    # Eight-digit bank account numbers in payment instructions are not EDF
    # customer identities. Prefer an explicitly labelled customer account;
    # older two-column covers can put website copy between label and number.
    accounts = []
    for match in re.finditer(
        r"Account\s+(?:number|reference)\s*:?\s*(A-\d+|\d(?:[ ]*\d)*)", text, re.I
    ):
        value = re.sub(r"\s+", "", match[1]).upper()
        context = text[max(0, match.start() - 120) : match.end() + 60]
        if (
            len(value) == 8
            and value.isdigit()
            and re.search(r"\bbank\b|sort\s*code|\bBACS\b", context, re.I)
        ):
            continue
        accounts.append(value)
    if not accounts:
        account_block = re.search(r"Account\s+(?:number|reference)\b(.{0,160})", text, re.I | re.S)
        if account_block:
            accounts = re.findall(r"\bA-\d+\b|(?<!\d)(?:\d[ ]*){12}(?!\d)", account_block[1])
    unique = {re.sub(r"\s+", "", value).upper() for value in accounts}
    account = next(iter(unique)) if len(unique) == 1 else ""
    meters = re.findall(
        r"(?:\bMeter\s*(?:(?:serial\s+)?number\s*)?:|charges\s+for\s+meter\s+|Meter\s+(?:serial\s+)?number\s+)"
        r"\s*([A-Z0-9-]*\d[A-Z0-9-]*)",
        text,
        re.I,
    )
    meter_values = {value.upper() for value in meters}
    meter = next(iter(meter_values)) if len(meter_values) == 1 else ""
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
        "Account": account,
        "Meter": meter,
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
