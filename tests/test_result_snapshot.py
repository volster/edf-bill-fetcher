"""Independent business examples and persistence/failure contracts for shared results."""

from datetime import date, datetime, timezone
from decimal import Decimal
from pathlib import Path
from threading import Event

import pytest

from edf_bill_fetcher.application.analysis_service import analyse_prepared
from edf_bill_fetcher.application.legacy_adapter import capture_legacy
from edf_bill_fetcher.application.preparation_service import prepare_snapshot
from edf_bill_fetcher.application.run_service import run
from edf_bill_fetcher.domain.json_values import freeze_json, object_value, strict_loads
from edf_bill_fetcher.domain.requests import RunRequest
from edf_bill_fetcher.io.persistence.result_codec import decode_result, encode_result
from edf_bill_fetcher.io.persistence.result_store import load_result, save_result
from edf_bill_fetcher.models.report_view import RenderSelection, build_report_view


def _request() -> RunRequest:
    return RunRequest(
        "example",
        datetime(2026, 1, 31, tzinfo=timezone.utc),
        date(2026, 1, 31),
        (),
        freeze_json({"as_of": "2026-01-31", "use_dedup": True}),
    )


def _records() -> list[dict]:
    common = {"Account": "A-100", "Fuel": "Electricity"}
    return [
        {
            **common,
            "Date": "01/01/2026",
            "Entry Type": "New Bill",
            "Amount Kind": "Balance",
            "Amount (£)": 100,
            "Period Charge (£)": 60,
            "Invoice #": "B1",
            "Period From": "01/12/2025",
            "Period To": "31/12/2025",
            "extension": {"note": ["retained"]},
        },
        {**common, "Date": "02/01/2026", "Entry Type": "Payment", "Amount (£)": -10},
        {**common, "Date": "03/01/2026", "Entry Type": "Credit", "Amount (£)": -5},
        {**common, "Date": "04/01/2026", "Entry Type": "Refund", "Amount (£)": 2},
        {
            **common,
            "Date": "31/01/2026",
            "Entry Type": "Ongoing Balance",
            "Amount Kind": "Balance",
            "Amount (£)": 0,
        },
    ]


def _snapshot():
    evidence = capture_legacy(_records(), {"extension": {"unknown": [1, 2]}}, _request())
    return analyse_prepared(prepare_snapshot(evidence))


def test_independent_money_example_preserves_zero_and_distinct_totals():
    snapshot = _snapshot()
    assert snapshot.summary.charges == Decimal("60.00")
    assert snapshot.summary.net_credits == Decimal("13.00")
    assert snapshot.summary.opening_balance == Decimal("100.00")
    assert snapshot.summary.closing_balance == Decimal("0.00")
    payment = build_report_view(snapshot, RenderSelection()).model("payment")
    assert payment.count == 2
    assert payment.total_paid == 15


def test_v3_roundtrip_preserves_originals_without_processing(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    snapshot = _snapshot()
    destination = tmp_path / "result.json"
    save_result(snapshot, destination, result_id="example-result")

    def forbidden(*args, **kwargs):
        raise AssertionError("Loading an accepted result must not calculate")

    monkeypatch.setattr("edf_bill_fetcher.application.analysis_service.analyse_prepared", forbidden)
    monkeypatch.setattr(
        "edf_bill_fetcher.application.preparation_service.prepare_snapshot", forbidden
    )
    restored = load_result(destination)
    assert restored == snapshot
    assert object_value(restored.prepared.extraction.observations[0].raw)["extension"] == {
        "note": ["retained"]
    }
    assert object_value(restored.prepared.extraction.legacy_state)["extension"] == {
        "unknown": [1, 2]
    }
    view = build_report_view(restored, RenderSelection())
    frame = view.frame()
    frame.iloc[0, 0] = "changed"
    assert not frame.equals(view.frame())


@pytest.mark.parametrize("value", [True, -1, "2"])
def test_invalid_stored_counts_are_rejected(value):
    encoded = encode_result(_snapshot(), "example-result")
    encoded["analysis"]["models"]["payment"]["payload"]["fields"]["count"] = value
    with pytest.raises(ValueError):
        decode_result(encoded)


@pytest.mark.parametrize("text", ['{"key":1,"key":2}', '{"amount":NaN}', '{"amount":1e400}'])
def test_strict_json_rejects_ambiguous_or_nonfinite_evidence(text: str):
    with pytest.raises(ValueError):
        strict_loads(text)


def test_result_replace_failure_preserves_previous_bytes(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    destination = tmp_path / "result.json"
    destination.write_bytes(b"previous accepted bytes")

    def fail_replace(*args):
        raise OSError("injected before publication")

    monkeypatch.setattr("edf_bill_fetcher.helpers.artifacts.os.replace", fail_replace)
    with pytest.raises(OSError, match="injected"):
        save_result(_snapshot(), destination, result_id="example-result")
    assert destination.read_bytes() == b"previous accepted bytes"
    assert list(tmp_path.iterdir()) == [destination]


def test_second_renderer_failure_preserves_result_and_first_artifact(tmp_path: Path):
    first, second = tmp_path / "first.html", tmp_path / "second.pdf"
    second.write_bytes(b"previous PDF")
    cleaned = []

    def render_first(view, path):
        Path(path).write_text("complete first report")

    def render_second(view, path):
        raise OSError("second report failed")

    outcome = run(
        _request(),
        {"html": str(first), "pdf": str(second)},
        result_path=str(tmp_path / "result.json"),
        manifest_path=str(tmp_path / "run.json"),
        collect=lambda: (_records(), {}),
        cleanup=lambda: cleaned.append(True),
        renderers={"html": render_first, "pdf": render_second},
    )
    assert outcome.status == "failed"
    assert load_result(tmp_path / "result.json") == outcome.snapshot
    assert first.read_text() == "complete first report"
    assert second.read_bytes() == b"previous PDF"
    assert set(object_value(outcome.manifest)["artifacts"]) == {"html"}
    assert cleaned == [True]


def test_cancel_between_stages_records_truthful_outcome_and_cleanup(tmp_path: Path):
    cancelled = Event()
    cleaned = []

    def progress(event):
        if event.stage == "preparing":
            cancelled.set()

    outcome = run(
        _request(),
        {"html": str(tmp_path / "report.html")},
        result_path=str(tmp_path / "result.json"),
        manifest_path=str(tmp_path / "run.json"),
        collect=lambda: (_records(), {}),
        cancelled=cancelled.is_set,
        progress=progress,
        cleanup=lambda: cleaned.append(True),
    )
    assert outcome.status == "cancelled"
    assert cleaned == [True]
    assert not (tmp_path / "report.html").exists()
    assert strict_loads((tmp_path / "run.json").read_text())["status"] == "cancelled"
