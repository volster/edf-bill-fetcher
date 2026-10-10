"""All artifact formats consume accepted results without analytical recomputation."""

import inspect
import json
import sys
from datetime import date, datetime, timezone
from pathlib import Path

import pytest

from edf_bill_fetcher.domain.json_values import freeze_json
from edf_bill_fetcher.domain.requests import RunRequest
from edf_bill_fetcher.io.persistence.result_codec import migrate_legacy
from edf_bill_fetcher.io.persistence.result_store import load_result, save_result
from edf_bill_fetcher.io.reporters.docx_report import render_docx
from edf_bill_fetcher.io.reporters.html_report import render_html
from edf_bill_fetcher.io.reporters.pdf_report import render_pdf
from edf_bill_fetcher.io.writers.export import render_xlsx
from edf_bill_fetcher.models.report_view import RenderSelection, build_report_view

CASES = json.loads((Path(__file__).parent / "fixtures/financial_contracts.json").read_text())[
    "cases"
]


@pytest.mark.parametrize(
    "case",
    [case for case in CASES if case["id"] in {"F01", "F02", "F06", "F12", "F18"}],
    ids=lambda case: case["id"],
)
def test_roundtripped_result_renders_four_formats_without_calculation(
    case: dict, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    request = RunRequest(
        "render-test",
        datetime(2024, 4, 1, tzinfo=timezone.utc),
        date(2024, 4, 1),
        (),
        freeze_json(case["config"]),
    )
    snapshot = migrate_legacy(case["records"], request)
    source = tmp_path / "result.json"
    save_result(snapshot, source, result_id="render-result")
    original = source.read_bytes()

    def forbidden(*args, **kwargs):
        raise AssertionError("An accepted result must not be recalculated during loading/rendering")

    # Patch defining functions and already-imported aliases in every app module.
    names = {
        "prepare_evidence",
        "prepare_snapshot",
        "analysis_frame",
        "balance_frame",
        "monthly_balance_frame",
        "financial_totals",
        "balance_extremes",
        "mean_daily_rate",
        "bill_daily_rates",
    }
    patched = []
    for module_name, module in list(sys.modules.items()):
        if not module_name.startswith("edf_bill_fetcher.") or module is None:
            continue
        for attr, function in list(vars(module).items()):
            if inspect.isfunction(function) and (
                function.__name__ in names
                or function.__name__.startswith(
                    ("compute_", "detect_", "estimate_", "analyse_", "match_sap", "run_analysers")
                )
            ):
                monkeypatch.setattr(module, attr, forbidden)
                patched.append((module_name, attr))
    assert patched
    restored = load_result(source)
    assert restored == snapshot
    for fmt, render in (
        ("xlsx", render_xlsx),
        ("pdf", render_pdf),
        ("docx", render_docx),
        ("html", render_html),
    ):
        view = build_report_view(restored, RenderSelection((fmt,)))
        destination = tmp_path / ("report." + fmt)
        render(view, str(destination))
        assert destination.stat().st_size > 100
        assert view.snapshot.input_fingerprint == snapshot.input_fingerprint
    assert source.read_bytes() == original
