"""Small, source-independent accepted analyses for integrity regression tests."""

from datetime import date, datetime, timezone

from edf_bill_fetcher.application.analysis_service import analyse_prepared
from edf_bill_fetcher.application.legacy_adapter import capture_legacy
from edf_bill_fetcher.application.preparation_service import prepare_snapshot
from edf_bill_fetcher.domain.json_values import freeze_json
from edf_bill_fetcher.domain.requests import RunRequest


def snapshot(records, state=None):
    config = {"as_of": "2025-03-01", "analysis_min": 0, "use_dedup": False}
    request = RunRequest(
        "synthetic-integrity",
        datetime(2025, 3, 1, tzinfo=timezone.utc),
        date(2025, 3, 1),
        (),
        freeze_json(config),
    )
    return analyse_prepared(prepare_snapshot(capture_legacy(records, state or {}, request)))
