"""Canonical evidence preparation shared by every export surface."""

from __future__ import annotations

import copy
import hashlib
import math
import warnings
from collections.abc import Mapping
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from typing import Any, cast

import pandas as pd

from edf_bill_fetcher.helpers.date_utils import completeness_score, parse_to_sort_date
from edf_bill_fetcher.helpers.identity import account_key
from edf_bill_fetcher.models.config import ConfigDict

DISPLAY_DEFAULTS: dict[str, Any] = {
    "Source": "Unknown",
    "Sender": "",
    "Date": "N/A",
    "Period From": "N/A",
    "Period To": "N/A",
    "Invoice #": "N/A",
    "Amount (£)": None,
    "Period Charge (£)": None,
    "Entry Type": "Unknown",
    "Reading": "Unknown",
    "Units (kWh)": None,
    "Standing Chg (p/day)": None,
    "Tariff": "N/A",
    "Attachment Name": "N/A",
    "Details": "",
    "Logic Used": "",
    "Sub Periods": "",
    "Cancel/Rebill Admitted": False,
}

SOURCE_PRECEDENCE = {
    "HTM Account History": 0,
    "Local PDF Folder": 1,
    "Statement Reconciliation": 1,
    "PST PDF Attachment": 2,
    "EML PDF Attachment": 2,
    "MSG PDF Attachment": 2,
    "Email Body": 3,
    "Email Body (RTF)": 3,
}


def number(value: Any) -> float | None:
    """Read a finite numeric value without turning missing data into zero."""
    try:
        result = float(str(value).replace("£", "").replace(",", "").strip())
    except (ValueError, TypeError):
        return None
    return result if math.isfinite(result) else None


def money(value: Any) -> float | None:
    """Normalize monetary input to pence using decimal rounding."""
    parsed = number(value)
    if parsed is None:
        return None
    try:
        return float(Decimal(str(parsed)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
    except InvalidOperation:
        return None


def identity(value: Any) -> str:
    """Normalize optional identity fields."""
    s = str(value).strip()
    return "" if s.lower() in {"", "n/a", "none", "nan", "unknown", "<na>"} else s


def normalize_record(record: dict[str, Any]) -> dict[str, Any]:
    """Preserve evidence and distinguish balance observations from transactions.

    Positive transactions increase debt; payments and credits are negative.
    Amount remains the legacy source figure; explicit fields carry its meaning.
    """
    row = {**DISPLAY_DEFAULTS, **copy.deepcopy(record)}
    kind = str(row["Entry Type"])
    amount = money(row.get("Amount (£)"))
    charge = money(row.get("Period Charge (£)"))
    role = identity(row.get("Amount Kind"))
    if not role:
        if str(row["Source"]).startswith("HTM") or kind == "Statement Reconciliation":
            role = "Balance"
        elif "Reconciliation Statement" in str(row["Logic Used"]):
            role = "Transaction"
        elif kind in {"Payment", "Credit", "Refund", "Late Payment"}:
            role = "Transaction"
        elif kind in {"New Bill", "Ongoing Balance"}:
            role = "Balance"
        else:
            role = "Unknown"
    balance = money(row.get("Balance (£)"))
    explicit = row.get("Review Explicit Amounts") is True
    if balance is None and role == "Balance" and not explicit:
        balance = amount
    transaction = money(row.get("Transaction Amount (£)"))
    if transaction is None and not explicit:
        if kind in {"Payment", "Credit"}:
            event_amount = (
                charge if charge is not None else amount if role == "Transaction" else None
            )
            transaction = -abs(event_amount) if event_amount is not None else None
        elif kind == "Refund":
            event_amount = (
                charge if charge is not None else amount if role == "Transaction" else None
            )
            transaction = abs(event_amount) if event_amount is not None else None
        elif kind in {"New Bill", "Charge", "Late Payment"} or (
            kind == "Ongoing Balance" and charge is not None
        ):
            transaction = (
                charge if charge is not None else amount if role == "Transaction" else None
            )
    row.update(
        {
            "Amount (£)": amount,
            "Period Charge (£)": charge,
            "Amount Kind": role,
            "Balance (£)": balance,
            "Transaction Amount (£)": transaction,
        }
    )
    row["Cancel/Rebill Admitted"] = str(row.get("Cancel/Rebill Admitted")).lower() == "true"
    row.setdefault("Account", "")
    row.setdefault("Meter", "")
    row.setdefault("Register", "")
    row.setdefault("Fuel", "")
    ids = row.get("Source IDs")
    if isinstance(ids, str):
        import json

        try:
            ids = json.loads(ids)
        except (ValueError, TypeError):
            ids = []
    row["Source IDs"] = (
        [str(v) for v in ids]
        if isinstance(ids, (list, tuple))
        else [row["Source ID"]]
        if row.get("Source ID")
        else []
    )
    # Only an extracted energy rate is a unit rate. Whole-bill/kWh is not.
    rate = number(row.get("Energy Unit Rate (p/kWh)"))
    if rate is None and row.get("Unit Rate Basis") == "Extracted":
        rate = number(row.get("Unit Rate (p/kWh)"))
    row["Unit Rate (p/kWh)"] = rate
    row["Energy Unit Rate (p/kWh)"] = rate
    row["% Change"] = None
    return row


def normalize_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Return the canonical schema, including for an empty input."""
    rows = [normalize_record(r) for r in df.to_dict("records")]
    if not rows:
        rows = [normalize_record({})]
        result = pd.DataFrame(rows).iloc[:0].copy()
    else:
        result = pd.DataFrame(rows)
    result["_dt"] = pd.to_datetime(result["Date"].apply(parse_to_sort_date))
    return result


def scope(row: dict[str, Any]) -> tuple[str, ...]:
    """Keep accounts, fuels, meters and registers in independent ledgers."""
    return (
        account_key(identity(row.get("Account"))),
        *(identity(row.get(k)).casefold() for k in ("Fuel", "Meter", "Register")),
    )


def _same_event(a: dict[str, Any], b: dict[str, Any]) -> bool:
    linked = identity(a.get("Reconciled Event ID"))
    if linked and linked == identity(b.get("Reconciled Event ID")):
        return True
    if scope(a) != scope(b) or a["Entry Type"] != b["Entry Type"]:
        return False
    ai, bi = identity(a.get("Invoice #")), identity(b.get("Invoice #"))
    if ai and bi and ai != bi:
        return False
    # A statement number identifies several transactions, not one invoice.
    statement = "Reconciliation Statement" in str(
        a.get("Logic Used")
    ) or "Reconciliation Statement" in str(b.get("Logic Used"))
    if ai and ai == bi and not statement and a["Entry Type"] == "New Bill":
        for key in ("Period From", "Period To", "Period Charge (£)", "Balance (£)"):
            left, right = identity(a.get(key)), identity(b.get(key))
            if left and right and left != right:
                return False
        return True
    ad, bd = parse_to_sort_date(a["Date"]), parse_to_sort_date(b["Date"])
    if pd.isna(ad) or pd.isna(bd) or ad != bd:
        return False
    same_period = all(
        identity(a.get(k)) == identity(b.get(k)) for k in ("Period From", "Period To")
    )
    amount_match = all(a.get(k) == b.get(k) for k in ("Balance (£)", "Transaction Amount (£)"))
    if not amount_match or not same_period:
        return False
    transaction_id = identity(a.get("Transaction ID"))
    same_transaction = transaction_id and transaction_id == identity(b.get("Transaction ID"))
    source_row = identity(a.get("Source Row"))
    same_occurrence = (
        source_row
        and source_row == identity(b.get("Source Row"))
        and identity(a.get("Source ID")) == identity(b.get("Source ID"))
        and identity(a.get("Page From")) == identity(b.get("Page From"))
        and bool(identity(a.get("Source ID")))
    )
    complete_period = all(identity(a.get(k)) for k in ("Period From", "Period To"))
    cross_source_bill = (
        not statement
        and a["Entry Type"] == "New Bill"
        and complete_period
        and a.get("Transaction Amount (£)") is not None
        and str(a["Source"]) != str(b["Source"])
    )
    return bool(
        cross_source_bill
        or (ai and ai == bi and not statement)
        or same_transaction
        or same_occurrence
    )


def _link_statement_observations(rows: list[dict[str, Any]]) -> None:
    """Link a uniquely supported historical line to its original invoice.

    Account, consumption period and signed charge must agree. Known meters
    and fuels must also agree; missing account identities never cross-join.
    Ambiguous candidate invoices stay separate and explicitly require review.
    """
    for row in rows:
        # Revalidate generated relationships after a review edit or reload.
        row.pop("Reconciled Event ID", None)
        if row.get("Billing Date Status") == "Corroborated by original invoice":
            row["First Billed Date"] = "N/A"
            row["Date"] = "N/A"
            row.pop("Original Invoice #", None)
            row["Billing Date Status"] = "Unknown: original relationship requires revalidation"
    originals = [r for r in rows if r.get("Document Role") != "Historical account statement"]
    for row in rows:
        if row.get("Document Role") != "Historical account statement":
            continue
        kind = row["Entry Type"]
        if kind not in {"Charge", "Credit"}:
            continue
        account = account_key(identity(row.get("Account")))
        if not account or not all(identity(row.get(k)) for k in ("Period From", "Period To")):
            continue
        same_lines = [
            other
            for other in rows
            if other.get("Document Role") == "Historical account statement"
            and other.get("Entry Type") == kind
            and scope(other) == scope(row)
            and identity(other.get("Source ID") or other.get("Attachment Name"))
            == identity(row.get("Source ID") or row.get("Attachment Name"))
            and all(
                other.get(key) == row.get(key)
                for key in ("Account", "Period From", "Period To", "Transaction Amount (£)")
            )
        ]
        if len(same_lines) > 1:
            row["Reconciliation Review"] = (
                "Repeated identical statement lines; event identity needs review"
            )
            continue
        candidates = []
        for original in originals:
            if str(original.get("Source", "")).startswith("HTM"):
                # Account-history posting/period-end dates are not evidence
                # of the issue date printed on an original invoice.
                continue
            if original["Entry Type"] != ("New Bill" if kind == "Charge" else "Credit"):
                continue
            if account_key(identity(original.get("Account"))) != account:
                continue
            if not identity(original.get("Invoice #")):
                continue
            issued = parse_to_sort_date(original["Date"])
            document_date = parse_to_sort_date(row.get("Document Date"))
            if pd.isna(issued) or (not pd.isna(document_date) and issued > document_date):
                continue
            if any(
                identity(row.get(k))
                and identity(original.get(k))
                and str(row[k]).casefold() != str(original[k]).casefold()
                for k in ("Meter", "Fuel", "Register")
            ):
                continue
            if any(
                parse_to_sort_date(row[k]) != parse_to_sort_date(original[k])
                for k in ("Period From", "Period To")
            ):
                continue
            if row.get("Transaction Amount (£)") != original.get("Transaction Amount (£)"):
                continue
            if kind == "Credit" and parse_to_sort_date(row["Date"]) != parse_to_sort_date(
                original["Date"]
            ):
                continue
            candidates.append(original)
        invoices = {identity(c["Invoice #"]) for c in candidates}
        if len(invoices) != 1 or len({scope(candidate) for candidate in candidates}) != 1:
            if candidates:
                row["Reconciliation Review"] = (
                    "Multiple original invoices match this historical line"
                )
            continue
        original = candidates[0]
        token = hashlib.sha256(
            repr(
                (
                    scope(original),
                    kind,
                    original["Invoice #"],
                    original["Period From"],
                    original["Period To"],
                    original["Transaction Amount (£)"],
                )
            ).encode()
        ).hexdigest()
        row["Reconciled Event ID"] = token
        row["Original Invoice #"] = original["Invoice #"]
        row["First Billed Date"] = original["Date"]
        row["Date"] = original["Date"]
        row["Billing Date Status"] = "Corroborated by original invoice"
        for candidate in candidates:
            candidate["Reconciled Event ID"] = token


@dataclass
class PreparedEvidence:
    """One preserved dataset and its auditable duplicate decisions."""

    records: list[dict[str, Any]]
    duplicates: list[dict[str, Any]]
    config: ConfigDict

    @property
    def frame(self) -> pd.DataFrame:
        """A fresh frame so writers cannot mutate canonical records."""
        return normalize_frame(pd.DataFrame(self.records))


def prepare_evidence(
    records: list[dict[str, Any]], config: Mapping[str, Any] | None = None
) -> PreparedEvidence:
    """Resolve canonical records once, retaining source occurrences and flags."""
    cfg = dict(config or {})
    cfg.setdefault("as_of", pd.Timestamp.today().date().isoformat())
    if pd.isna(parse_to_sort_date(cfg["as_of"])):
        raise ValueError("as_of must be a valid valuation date")
    for key in ("min_amount", "analysis_min", "credit_interest_rate"):
        if key in cfg and (number(cfg[key]) is None or float(cfg[key]) < 0):
            raise ValueError(f"{key} must be a finite non-negative number")
    rows = [normalize_record(r) for r in records]
    _link_statement_observations(rows)
    rows.sort(
        key=lambda r: (
            r.get("Document Role") == "Historical account statement",
            -completeness_score(r),
            SOURCE_PRECEDENCE.get(str(r["Source"]), 99),
            str(r["Source"]),
            str(r["Date"]),
            str(r["Invoice #"]),
        )
    )
    kept: list[dict[str, Any]] = []
    duplicates: list[dict[str, Any]] = []
    for row in rows:
        winner = (
            next((r for r in kept if _same_event(row, r)), None)
            if cfg.get("use_dedup", True)
            else None
        )
        if winner is None:
            kept.append(row)
            continue
        duplicate = copy.deepcopy(row)
        duplicate["Duplicate Of"] = f"{winner['Source']} · {winner['Date']} · {winner['Invoice #']}"
        duplicate["_kept_record"] = winner
        duplicates.append(duplicate)
        winner["Cancel/Rebill Admitted"] |= row["Cancel/Rebill Admitted"]
        winner["Source IDs"] = sorted(set(winner["Source IDs"] + row["Source IDs"]))
        occurrences = winner.setdefault("Source Occurrences", [])
        occurrences.append(
            {
                k: row.get(k)
                for k in (
                    "Source ID",
                    "Attachment Name",
                    "Original Filename",
                    "Content Hash",
                    "Page From",
                    "Page To",
                    "Source PDF Text",
                    "_regex_trace",
                    "Source Row",
                    "Document Date",
                    "First Billed Date",
                    "Original Invoice #",
                    "Document Role",
                )
            }
        )
        if cfg.get("amalgamate_duplicates", False):
            for key, val in row.items():
                if not identity(winner.get(key)) and identity(val):
                    winner[key] = copy.deepcopy(val)

    def order(r: dict[str, Any]) -> tuple:
        date = parse_to_sort_date(r["Date"])
        return (
            pd.isna(date),
            pd.Timestamp.max if pd.isna(date) else date,
            str(r["Invoice #"]),
            str(r["Source"]),
        )

    kept.sort(key=order)
    for row in duplicates:
        winner = row.pop("_kept_record")
        row["_matches_kept_idx"] = next(i for i, r in enumerate(kept) if r is winner)
    return PreparedEvidence(kept, duplicates, cast(ConfigDict, cfg))


def analysis_frame(df: pd.DataFrame, config: Mapping[str, Any] | None = None) -> pd.DataFrame:
    """Apply analysis eligibility consistently without removing raw evidence."""
    work = normalize_frame(df)
    cfg = config or {}
    minimum = float(cfg.get("analysis_min", 0.0))
    magnitude = pd.to_numeric(work["Period Charge (£)"], errors="coerce").abs()
    magnitude = magnitude.fillna(pd.to_numeric(work["Balance (£)"], errors="coerce").abs())
    bill = work["Entry Type"].isin(["New Bill", "Ongoing Balance"])
    bill_date = work["_dt"]
    start = pd.to_datetime(work["Period From"].apply(parse_to_sort_date))
    invalid = bill & (
        (bill_date.isna() & work["Date"].map(identity).ne(""))
        | (start.isna() & work["Period From"].map(identity).ne(""))
    )
    if invalid.any():
        warnings.warn(
            f"{int(invalid.sum())} billing records have unparseable dates; dated analysis is unavailable",
            UserWarning,
            stacklevel=2,
        )
    late = (bill_date - start).dt.days > 365
    keep = ~bill | (magnitude >= minimum) | late
    return work[keep].sort_values("_dt", kind="stable").reset_index(drop=True)


def balance_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Select actual balance observations for balance statistics."""
    if "Entry Type" not in df and "Amount Kind" not in df and "Balance (£)" not in df:
        # The historical analytical API accepts a bare balance series.
        df = df.assign(**{"Amount Kind": "Balance"})
    work = normalize_frame(df)
    work = work[work["_dt"].notna() & work["Balance (£)"].notna()].copy()
    work["Amount (£)"] = work["Balance (£)"]
    return work.sort_values("_dt", kind="stable").reset_index(drop=True)


def monthly_balance_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Regular month-end balances; missing months remain explicit gaps."""
    work = balance_frame(df)
    if work.empty or len({scope(r) for r in work.to_dict("records")}) > 1:
        return work.iloc[:0]
    series = work.set_index("_dt")["Balance (£)"].resample(pd.offsets.MonthEnd()).last()
    return pd.DataFrame(
        {
            "Date": series.index.strftime("%d/%m/%Y"),
            "Amount (£)": series.values,
            "Balance (£)": series.values,
            "Amount Kind": "Balance",
            "_dt": series.index,
        }
    )
