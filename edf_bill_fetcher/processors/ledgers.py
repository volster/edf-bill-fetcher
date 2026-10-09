"""Keep independent account, fuel, meter and register evidence in separate analyses."""

from __future__ import annotations

import inspect
from functools import partial, wraps

import pandas as pd

from edf_bill_fetcher.processors.preparation import normalize_frame, scope


def invoice_reference(record, invoice):
    """Create an unambiguous analysis reference while retaining the original invoice in evidence."""
    label = "/".join(value or "unknown" for value in scope(record))
    return f"{label} :: {invoice}"


def scoped_analysis(function):
    """Run a frame detector independently for each ledger, qualifying mixed-ledger references."""
    signature = inspect.signature(function)

    @wraps(function)
    def wrapped(*args, **kwargs):
        bound = signature.bind(*args, **kwargs)
        data = bound.arguments["df"]
        if data is None or data.empty:
            return function(*args, **kwargs)
        work = normalize_frame(data)
        groups: dict[tuple[str, ...], list[int]] = {}
        for idx, record in enumerate(work.to_dict("records")):
            groups.setdefault(scope(record), []).append(idx)
        if len(groups) <= 1:
            return function(*args, **kwargs)
        results = []
        all_evidence = bound.arguments.get("evidence_df")
        for indexes in groups.values():
            ledger = work.iloc[indexes].copy()
            bound.arguments["df"] = ledger
            if "evidence_df" in bound.arguments and bound.arguments["evidence_df"] is not None:
                evidence = normalize_frame(all_evidence)
                key = scope(ledger.iloc[0].to_dict())
                bound.arguments["evidence_df"] = evidence[
                    [scope(r) == key for r in evidence.to_dict("records")]
                ]
            result = function(*bound.args, **bound.kwargs)
            first = ledger.iloc[0].to_dict()
            for field in ("Account", "Fuel", "Meter", "Register"):
                result[field] = first.get(field, "")
            for field in ("Invoice #", "Killer Invoice", "Killed Invoice"):
                if field in result:
                    result[field] = result[field].map(partial(invoice_reference, first))
            results.append(result)
        return pd.concat(results, ignore_index=True)

    return wrapped
