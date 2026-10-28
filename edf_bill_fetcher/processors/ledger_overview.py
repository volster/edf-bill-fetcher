"""Accepted per-ledger financial observations for mixed or linked account evidence."""

from __future__ import annotations

from collections import defaultdict
from decimal import Decimal

import pandas as pd

from edf_bill_fetcher.domain.money import money_text
from edf_bill_fetcher.processors.financial import financial_totals
from edf_bill_fetcher.processors.preparation import balance_frame, identity, number, scope


def ledger_overview(frame: pd.DataFrame) -> list[dict]:
    """Store dated balances and movements separately without inventing a current balance."""
    groups: dict[tuple, list[dict]] = defaultdict(list)
    for row in frame.to_dict("records"):
        if (
            number(row.get("Balance (£)")) is not None
            or number(row.get("Transaction Amount (£)")) is not None
        ):
            groups[scope(row)].append(row)
    result = []
    for key, rows in sorted(groups.items()):
        work = pd.DataFrame(rows)
        charges, credits = financial_totals(work)
        balances = balance_frame(work)
        first = balances.iloc[0] if not balances.empty else None
        last = balances.iloc[-1] if not balances.empty else None
        result.append(
            {
                "scope": list(key),
                "references": sorted({identity(row.get("Account")) for row in rows} - {""}),
                "records": len(rows),
                "charges": money_text(Decimal(str(charges))),
                "net_credits": money_text(Decimal(str(credits))),
                "opening_date": first["_dt"].date().isoformat() if first is not None else None,
                "opening_balance": money_text(Decimal(str(first["Balance (£)"])))
                if first is not None
                else None,
                "closing_date": last["_dt"].date().isoformat() if last is not None else None,
                "closing_balance": money_text(Decimal(str(last["Balance (£)"])))
                if last is not None
                else None,
            }
        )
    return result
