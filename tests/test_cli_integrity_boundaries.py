"""Public CLI failures preserve inputs; analytical overrides create a new accepted result."""

import json
import sys
from types import SimpleNamespace

import pytest

from edf_bill_fetcher.domain.json_values import object_value
from edf_bill_fetcher.io import cli
from edf_bill_fetcher.io.persistence.result_codec import encode_result
from edf_bill_fetcher.io.persistence.result_store import load_result
from tests.synthetic_analysis import snapshot
from tests.test_financial_identity_boundaries import bill


def write(path, value):
    path.write_text(json.dumps(value), encoding="utf-8")
    return str(path)


def options(tmp_path, records, config=None, engine=None):
    return SimpleNamespace(
        records=write(tmp_path / "records.json", records),
        output=str(tmp_path / "report.html"),
        config=write(tmp_path / "config.json", config) if config is not None else None,
        engine_data=write(tmp_path / "engine.json", engine) if engine is not None else None,
    )


@pytest.mark.parametrize(
    ("records", "config", "engine", "message"),
    [
        ({"version": 4, "records": []}, None, None, "Unsupported evidence"),
        ({"records": [1]}, None, None, "JSON list of objects"),
        ([], [], None, "Configuration must be a JSON object"),
        ([], None, [], "Unsupported engine"),
        ([], None, {"version": 4}, "Unsupported engine"),
    ],
)
def test_legacy_report_input_validation_is_explicit(tmp_path, records, config, engine, message):
    parsed = options(tmp_path, records, config, engine)
    with pytest.raises(ValueError, match=message):
        cli._load_report_inputs(parsed)


def test_legacy_configuration_overrides_preserve_engine_diagnostics(tmp_path):
    parsed = options(
        tmp_path,
        [bill()],
        {"analysis_min": 5},
        {
            "version": 2,
            "config": {"analysis_min": 10},
            "pdf_count": 7,
            "error_log": ["Unparsed attachment"],
        },
    )
    records, config, engine, filtered = cli._load_report_inputs(parsed)
    assert records == [bill()]
    assert config["analysis_min"] == 5
    assert engine.pdf_count == 7
    assert engine.error_log == ["Unparsed attachment"]
    assert filtered == []


def test_report_inputs_never_accept_legacy_pickle_metadata(tmp_path):
    parsed = options(tmp_path, [])
    parsed.engine_data = str(tmp_path / "unsafe.pkl")
    with pytest.raises(ValueError, match="not a safe interchange"):
        cli._load_report_inputs(parsed)


@pytest.mark.parametrize(
    ("records", "config", "engine", "message"),
    [
        ([], [], None, "Configuration must be a JSON object"),
        (42, None, None, "evidence object or array"),
        ([], None, [], "Unsupported engine"),
        ([], None, {"version": 4}, "Unsupported engine"),
        ([bill()], None, {"records": []}, "Conflicting evidence"),
    ],
)
def test_invalid_report_request_writes_no_output(tmp_path, records, config, engine, message):
    parsed = options(tmp_path, records, config, engine)
    with pytest.raises(ValueError, match=message):
        cli._render_cli_report(parsed, "html")
    assert not (tmp_path / "report.html").exists()
    assert not (tmp_path / "report.html.result.json").exists()


def test_accepted_result_rejects_extra_engine_metadata(tmp_path):
    parsed = options(tmp_path, encode_result(snapshot([bill()]), "accepted"), engine={})
    with pytest.raises(ValueError, match="already contains its engine metadata"):
        cli._render_cli_report(parsed, "html")


def test_analytical_override_reanalyses_originals_without_overwriting_accepted_input(tmp_path):
    baseline = snapshot([bill(), bill(**{"Source ID": "copy"})])
    assert str(baseline.summary.charges) == "200.00"
    parsed = options(
        tmp_path,
        encode_result(baseline, "accepted"),
        {"use_dedup": True, "report_sections": ["data_quality"]},
    )
    original_bytes = (tmp_path / "records.json").read_bytes()
    cli._render_cli_report(parsed, "html")
    revised = load_result(parsed.output + ".result.json")
    assert str(revised.summary.charges) == "100.00"
    assert len(revised.prepared.extraction.observations) == 2
    assert revised.prepared.extraction.request.run_id != baseline.prepared.extraction.request.run_id
    assert revised.input_fingerprint != baseline.input_fingerprint
    assert (tmp_path / "records.json").read_bytes() == original_bytes


def test_matching_legacy_engine_records_merge_without_losing_configuration(tmp_path):
    records = [bill()]
    parsed = options(
        tmp_path,
        records,
        {"report_sections": ["data_quality"]},
        {"version": 2, "records": records, "config": {"as_of": "2025-03-01"}, "pdf_count": 2},
    )
    cli._render_cli_report(parsed, "html")
    revised = load_result(parsed.output + ".result.json")
    assert str(revised.summary.charges) == "100.00"
    assert object_value(revised.prepared.extraction.legacy_state)["pdf_count"] == 2


@pytest.mark.parametrize("target", ["report", "result", "manifest"])
def test_report_artifacts_cannot_overwrite_an_input(tmp_path, target):
    parsed = options(tmp_path, encode_result(snapshot([bill()]), "accepted"), {})
    if target == "report":
        parsed.output = parsed.records
    elif target == "result":
        replacement = tmp_path / "report.html.result.json"
        replacement.write_bytes((tmp_path / "records.json").read_bytes())
        parsed.records = str(replacement)
    else:
        parsed.config = write(tmp_path / "report.html.run.json", {})
    before = {p: p.read_bytes() for p in tmp_path.iterdir()}
    with pytest.raises(ValueError, match="overwrite an input"):
        cli._render_cli_report(parsed, "html")
    assert all(p.read_bytes() == content for p, content in before.items())


def test_missing_saved_source_is_reported_without_changing_accepted_figures(tmp_path, capsys):
    state = {
        "source_manifest": {
            "original": {
                "Saved Path": "evidence_files/missing.pdf",
                "Content Hash": "a" * 64,
                "Attachment Name": "missing.pdf",
            }
        }
    }
    baseline = snapshot([bill()], state)
    parsed = options(
        tmp_path, encode_result(baseline, "accepted"), {"report_sections": ["data_quality"]}
    )
    cli._render_cli_report(parsed, "html")
    assert "WARNING:" in capsys.readouterr().err
    assert load_result(parsed.output + ".result.json").summary == baseline.summary


@pytest.mark.parametrize("invalid", ["configuration", "missing_source"])
def test_extract_validation_stops_before_creating_a_workbook(tmp_path, capsys, invalid):
    source = tmp_path / "source.htm"
    source.write_text("Synthetic source", encoding="utf-8")
    args = ["--htm", str(source), "-o", str(tmp_path / "out.xlsx")]
    if invalid == "configuration":
        args += ["--min-amount", "-1"]
    else:
        source.unlink()
    with pytest.raises(SystemExit) as error:
        cli.run_cli_extract(args)
    assert error.value.code == 2
    assert (
        "min_amount" in capsys.readouterr().err
        if invalid == "configuration"
        else not source.exists()
    )
    assert not (tmp_path / "out.xlsx").exists()


@pytest.mark.parametrize("collision", ["source", "duplicate_output"])
def test_extract_destination_collision_cleans_up_before_coordinator_starts(
    tmp_path, monkeypatch, capsys, collision
):
    from edf_bill_fetcher.collectors import engine

    source = tmp_path / "source.htm"
    source.write_text("Preserve original bytes", encoding="utf-8")
    cleaned = []
    monkeypatch.setattr(
        engine,
        "EvidenceEngine",
        lambda *args: SimpleNamespace(cleanup=lambda: cleaned.append(True)),
    )
    output = source if collision == "source" else tmp_path / "out.xlsx"
    args = ["--htm", str(source), "-o", str(output)]
    if collision == "duplicate_output":
        args += ["--records-json", str(output)]
    with pytest.raises(SystemExit) as error:
        cli.run_cli_extract(args)
    assert error.value.code == 1
    assert cleaned == [True]
    assert "must be distinct" in capsys.readouterr().err
    assert source.read_text(encoding="utf-8") == "Preserve original bytes"


def test_invalid_diff_inputs_are_a_failure_instead_of_an_empty_success(tmp_path, capsys):
    old = write(tmp_path / "old.json", {"records": {}})
    new = write(tmp_path / "new.json", [])
    with pytest.raises(SystemExit) as error:
        cli.run_cli_diff([old, new])
    assert error.value.code == 1
    assert "both diff inputs must be" in capsys.readouterr().err


@pytest.mark.parametrize(
    ("value", "expected"), [(None, "-"), (float("nan"), "-"), ("unavailable", "unavailable")]
)
def test_diff_missing_amounts_do_not_turn_into_zero(value, expected):
    assert cli._diff_amount_text(value) == expected


def test_cli_without_qt_explains_how_to_run_headless(monkeypatch, capsys):
    monkeypatch.setattr(sys, "argv", ["edf-collector"])
    monkeypatch.setattr(cli, "HAS_QT", False)
    with pytest.raises(SystemExit) as error:
        cli.main()
    assert error.value.code == 2
    assert "--extract" in capsys.readouterr().err


def test_doctor_dispatch_preserves_arguments_and_exit_status(monkeypatch):
    from edf_bill_fetcher.helpers import capabilities

    called = []

    def doctor(args):
        called.append(args)
        return 7

    monkeypatch.setattr(capabilities, "main", doctor)
    monkeypatch.setattr(sys, "argv", ["edf-collector", "--doctor", "--json"])
    with pytest.raises(SystemExit) as error:
        cli.main()
    assert error.value.code == 7
    assert called == [["--json"]]


def test_gui_probe_dispatch_preserves_failure_status(tmp_path, monkeypatch):
    from edf_bill_fetcher.ui import qt_smoke

    called = []
    report = tmp_path / "smoke.json"

    def check_gui(path, platform):
        called.append((path, platform))
        return 3

    monkeypatch.setattr(qt_smoke, "check_gui", check_gui)
    monkeypatch.setattr(cli, "HAS_QT", True)
    monkeypatch.setattr(sys, "argv", ["edf-collector", "--check-gui", "--report", str(report)])
    with pytest.raises(SystemExit) as error:
        cli.main()
    assert error.value.code == 3
    assert called == [(report, "cocoa" if sys.platform == "darwin" else None)]


def test_regeneration_error_is_a_nonzero_cli_exit(tmp_path, monkeypatch, capsys):
    from edf_bill_fetcher.application import workbook_review

    monkeypatch.setattr(
        workbook_review,
        "regenerate_workbook",
        lambda *args, **kwargs: (
            str(tmp_path),
            SimpleNamespace(status="failed", error="Failed to save analysis"),
        ),
    )
    with pytest.raises(SystemExit) as error:
        cli.run_cli_review(["-i", "review.xlsx", "-o", str(tmp_path)])
    assert error.value.code == 1
    assert "Failed to save analysis" in capsys.readouterr().err
