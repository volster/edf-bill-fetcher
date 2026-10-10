"""Small inspection adapter for independently authored financial examples."""

from datetime import date, datetime, timezone
from uuid import uuid4

from edf_bill_fetcher.application.analysis_service import analyse_prepared
from edf_bill_fetcher.application.legacy_adapter import capture_legacy
from edf_bill_fetcher.application.preparation_service import prepare_snapshot
from edf_bill_fetcher.domain.json_values import freeze_json, object_value, rows_value
from edf_bill_fetcher.domain.money import money_text, parse_money
from edf_bill_fetcher.domain.requests import RunRequest
from edf_bill_fetcher.models.snapshot_adapter import from_snapshot_json


def probe(records: list[dict], config: dict) -> dict:
    """Expose actual snapshot values without duplicating any business computation."""
    cfg = dict(config)
    cfg.setdefault("as_of", date.today().isoformat())
    request = RunRequest(
        uuid4().hex,
        datetime.now(timezone.utc),
        date.fromisoformat(cfg["as_of"]),
        (),
        freeze_json(cfg),
    )
    result = analyse_prepared(prepare_snapshot(capture_legacy(records, {}, request)))
    prepared, summary = result.prepared, result.summary
    models = object_value(result.models)
    payment = from_snapshot_json(models["payment"]["payload"])
    forecast = from_snapshot_json(models["forecast"]["payload"])
    return {
        "record_count": len(rows_value(prepared.records)),
        "duplicate_count": len(rows_value(prepared.duplicates)),
        "charges": money_text(summary.charges),
        "net_credits": money_text(summary.net_credits),
        "payment_count": payment.count,
        "gross_payments_credits": money_text(parse_money(payment.total_paid)),
        "opening_balance": money_text(summary.opening_balance),
        "closing_balance": money_text(summary.closing_balance),
        "mean_daily": summary.mean_daily_charge,
        "forecast_n": forecast.n,
        "records": rows_value(prepared.records),
    }
