"""Business commands remain usable and financially equivalent outside the browser."""

import json
from datetime import date, datetime, timezone
from pathlib import Path
from unittest.mock import Mock

import pytest
from PySide6.QtWidgets import QApplication, QComboBox, QDialog, QLineEdit

from edf_bill_fetcher.application.comparison import compare_files, comparison_records
from edf_bill_fetcher.application.configuration import validate_settings
from edf_bill_fetcher.application.intake_execution import main as inspect_main
from edf_bill_fetcher.application.review_cli import main as review_main
from edf_bill_fetcher.application.review_journal import ReviewJournal
from edf_bill_fetcher.application.review_session import ReviewSession
from edf_bill_fetcher.domain.json_values import freeze_json
from edf_bill_fetcher.domain.requests import RunRequest, SourceSelection
from edf_bill_fetcher.io.cli import main, run_cli_diff, run_cli_extract
from edf_bill_fetcher.io.persistence.result_codec import encode_result
from edf_bill_fetcher.io.persistence.result_store import load_result, save_result
from edf_bill_fetcher.models.config import ConfigDict
from edf_bill_fetcher.ui.qt_configuration import ConfigurationDialog
from edf_bill_fetcher.ui.qt_jobs import DesktopJob, RunWorker
from tests.synthetic_analysis import snapshot
from tests.test_mixed_intake import mixed_folder


@pytest.fixture
def application():
    app = QApplication.instance() or QApplication([])
    assert isinstance(app, QApplication)
    yield app
    app.processEvents()


def baseline(tmp_path):
    result = snapshot(
        [
            {
                "Date": "01/02/2024",
                "Entry Type": "New Bill",
                "Invoice #": "ONE",
                "Account": "A-123",
                "Period Charge (£)": 100,
                "Balance (£)": 100,
            }
        ]
    )
    path = tmp_path / "result.json"
    save_result(result, path, result_id="baseline")
    return path, result


def test_cli_journal_undo_redo_and_recovery_are_bound_to_the_actual_draft(tmp_path):
    path, result = baseline(tmp_path)
    row = ReviewSession(result).rows[0]["id"]
    command_file = tmp_path / "command.json"
    draft = tmp_path / "review.json"

    def invoke(operation, arguments):
        command_file.write_text(json.dumps({"operation": operation, "arguments": arguments}))
        review_main(["--result", str(path), "--draft", str(draft), "--command", str(command_file)])

    invoke(
        "update",
        {
            "row_id": row,
            "changes": {"Period Charge (£)": "80.00"},
            "include": True,
            "reason": "Checked original",
            "status": "verified",
        },
    )
    assert ReviewSession.load(result, draft).record(row)["Period Charge (£)"] == 80
    invoke("undo", {})
    assert ReviewSession.load(result, draft).record(row)["Period Charge (£)"] == 100
    invoke("redo", {})
    assert ReviewSession.load(result, draft).record(row)["Period Charge (£)"] == 80
    external = ReviewSession(result)
    external.update(row, {"Period Charge (£)": "70.00"}, True, "Separate verified edit", "verified")
    external.save(draft)
    before = draft.read_bytes()
    with pytest.raises(ValueError, match="no longer matches"):
        invoke("undo", {})
    assert draft.read_bytes() == before
    assert load_result(path) == result


@pytest.mark.parametrize(
    ("history", "position"), [([], 0), ([{}], True), ([{}], -1), ([{}], 1), ([{}] * 102, 0)]
)
def test_invalid_journal_coordinates_never_restore_untrusted_history(history, position):
    with pytest.raises(ValueError):
        ReviewJournal(snapshot([]), history, position)


def test_journal_rejects_undo_without_a_change_or_with_arguments():
    result = snapshot([])
    journal = ReviewJournal(result, [ReviewSession(result).payload()], 0)
    with pytest.raises(ValueError, match="No change"):
        journal.apply("undo", {})
    with pytest.raises(ValueError, match="arguments"):
        journal.apply("redo", {"unexpected": True})


@pytest.mark.parametrize(
    "config",
    [
        {"unknown": 1},
        {"min_amount": float("nan")},
        {"analysis_min": True},
        {"credit_hold_days": 1.2},
        {"as_of": "20250101"},
        {"as_of": "2025-02-30"},
        {"report_sections": ["payment", "payment"]},
        {"report_sections": [True]},
        {"report_sections": "payment"},
        {"acc_num": False},
    ],
)
def test_shared_configuration_rejects_ambiguous_input(config):
    with pytest.raises(ValueError):
        validate_settings(config)


def test_complete_desktop_settings_validate_before_applying(application):
    widget = ConfigurationDialog(
        {
            "credit_interest_rate": 5.5,
            "credit_hold_days": 45,
            "fuel": "Electricity",
            "ofgem_auto_carry": False,
        }
    )
    assert set(widget.controls) == set(ConfigDict.__annotations__) - {"report_sections"}
    rate = widget.controls["credit_interest_rate"]
    assert isinstance(rate, QLineEdit)
    rate.setText("nan")
    widget.accept()
    assert widget.result() != QDialog.DialogCode.Accepted
    assert widget.error.text()
    rate.setText("6.25")
    kind = widget.controls["ofgem_auto_carry"]
    assert isinstance(kind, QComboBox)
    kind.setCurrentIndex(1)
    widget.accept()
    assert widget.result_config["credit_interest_rate"] == 6.25
    assert widget.result_config["credit_hold_days"] == 45
    assert widget.result_config["ofgem_auto_carry"] is True
    widget.close()


def test_desktop_worker_can_inspect_and_compare_without_web_dependencies(tmp_path, application):
    sources = mixed_folder(tmp_path / "sources")
    request = RunRequest(
        "desktop-intake",
        datetime(2025, 3, 1, tzinfo=timezone.utc),
        date(2025, 3, 1),
        (SourceSelection("auto", str(sources)),),
        freeze_json({"as_of": "2025-03-01"}),
    )
    directory = tmp_path / "inspection"
    directory.mkdir()
    worker = RunWorker(DesktopJob(request, str(directory), "Intake", (), (), "intake"))
    outcomes: list = []
    worker.outcome_ready.connect(outcomes.append)
    worker.run()
    assert outcomes[0].status == "complete"
    assert (directory / "Intake.xlsx").is_file()
    assert any(
        row["kind"] == "pdf"
        for row in json.loads((directory / "intake.json").read_text(encoding="utf-8"))
    )
    first, result = baseline(tmp_path)
    second = tmp_path / "second.json"
    save_result(result, second, result_id="other")
    output = tmp_path / "comparison"
    output.mkdir()
    worker = RunWorker(
        DesktopJob(
            request,
            str(output),
            "Comparison",
            ("xlsx",),
            (),
            "diff",
            str(first),
            other_path=str(second),
        )
    )
    outcomes = []
    worker.outcome_ready.connect(outcomes.append)
    worker.run()
    assert outcomes[0].status == "complete"
    assert json.loads((output / "comparison.json").read_text(encoding="utf-8")) == {
        "added": [],
        "removed": [],
        "changed": [],
    }


def test_cli_inspection_and_plan_reuse_preserve_the_independent_charge_total(tmp_path, capsys):
    sources = mixed_folder(tmp_path / "inputs")
    inspection = tmp_path / "inspection"
    inspect_main([str(sources), "--output-dir", str(inspection)])
    with pytest.raises(ValueError, match="fresh empty"):
        inspect_main([str(sources), "--output-dir", str(inspection)])
    config = tmp_path / "config.json"
    config.write_text(
        json.dumps({"as_of": "2025-03-01", "use_domain_filter": False, "use_dedup": False})
    )
    run_cli_extract(
        [
            "--intake-plan",
            str(inspection / "intake.json"),
            "--config",
            str(config),
            "-o",
            str(tmp_path / "evidence.xlsx"),
        ]
    )
    results = [path for path in tmp_path.glob("*/result.json") if path.parent != inspection]
    assert len(results) == 1
    assert load_result(results[0]).summary.charges == 180


def test_comparison_accepts_saved_results_and_cli_preserves_inputs(tmp_path, capsys):
    path, result = baseline(tmp_path)
    second = tmp_path / "second.json"
    second.write_text(json.dumps(encode_result(result, "other")))
    assert compare_files(str(path), str(second)) == {"added": [], "removed": [], "changed": []}
    run_cli_diff([str(path), str(second), "--diff-output", str(tmp_path / "diff.xlsx")])
    assert (tmp_path / "diff.xlsx").is_file()
    assert "Added:   0" in capsys.readouterr().out
    before = path.read_bytes()
    with pytest.raises(ValueError, match="overwrite"):
        run_cli_diff([str(path), str(second), "--diff-output", str(path)])
    assert path.read_bytes() == before
    with pytest.raises(ValueError):
        comparison_records({"records": ["invalid"]})


def test_cli_dispatches_optional_web_and_inspection_without_starting_qt(monkeypatch):
    import sys

    import edf_bill_fetcher.application.intake_execution as intake
    import edf_bill_fetcher.application.review_cli as review
    import edf_bill_fetcher.web.launch as web

    for option, module in [("--web", web), ("--inspect", intake), ("--review-command", review)]:
        callback = Mock()
        monkeypatch.setattr(module, "main", callback)
        monkeypatch.setattr(sys, "argv", ["main.py", option, "--help"])
        main()
        callback.assert_called_once_with(["--help"])


@pytest.fixture
def desktop(application, tmp_path):
    from PySide6.QtCore import QCoreApplication, QEvent

    from edf_bill_fetcher.ui.qt_app import EvidenceWindow
    from tests.test_qt_desktop import wait_until

    widget = EvidenceWindow(tmp_path / "preferences.json")
    wait_until(application, lambda: not widget.extraction_readiness.checking)
    yield widget
    widget.close()
    widget.deleteLater()
    QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
    application.processEvents()


def test_desktop_complete_settings_round_trip_without_losing_overrides(desktop, monkeypatch):
    import edf_bill_fetcher.ui.qt_configuration as configuration_module

    values = {
        "as_of": "2025-02-28",
        "credit_interest_rate": 0.08,
        "credit_hold_days": 30,
        "fuel": "Electricity",
        "ofgem_auto_carry": True,
        "min_amount": 25,
        "use_dedup": False,
    }
    dialog = Mock(result_config=values)
    dialog.exec.return_value = QDialog.DialogCode.Accepted
    factory = Mock(return_value=dialog)
    monkeypatch.setattr(configuration_module, "ConfigurationDialog", factory)
    desktop._advanced_settings()
    config = desktop._configuration()
    assert all(config[key] == value for key, value in values.items())
    assert "as_of" not in factory.call_args.args[0]
    desktop.save_settings()
    desktop._extra_config = {}
    desktop.load_settings()
    assert desktop._configuration()["as_of"] == "2025-02-28"
    assert desktop._configuration()["credit_interest_rate"] == 0.08
    dialog.exec.return_value = QDialog.DialogCode.Rejected
    before = desktop._configuration()
    desktop._advanced_settings()
    assert "as_of" in factory.call_args.args[0]
    assert desktop._configuration() == before


def test_desktop_comparison_requires_two_inputs_and_captures_an_isolated_job(
    desktop, tmp_path, monkeypatch
):
    from PySide6.QtWidgets import QFileDialog

    old, result = baseline(tmp_path)
    new = tmp_path / "later.json"
    save_result(result, new, result_id="later")
    jobs: list = []
    monkeypatch.setattr(desktop, "_begin_job", jobs.append)
    monkeypatch.setattr(desktop.review_page, "allow_leave", lambda: True)
    chooser = Mock(side_effect=[("", ""), (str(old), ""), ("", ""), (str(old), ""), (str(new), "")])
    monkeypatch.setattr(QFileDialog, "getOpenFileName", chooser)
    desktop._busy = True
    desktop._compare_runs()
    chooser.assert_not_called()
    desktop._busy = False
    desktop._compare_runs()
    desktop._compare_runs()
    assert not jobs
    desktop._compare_runs()
    assert len(jobs) == 1
    job = jobs[0]
    assert job.mode == "diff"
    assert job.input_path == str(old)
    assert job.other_path == str(new)
    assert Path(job.directory).is_dir()
    assert Path(job.directory) != old.parent
    worker = RunWorker(job)
    outcomes: list = []
    worker.outcome_ready.connect(outcomes.append)
    worker.run()
    assert outcomes[0].status == "complete"
    assert (Path(job.directory) / "comparison.json").is_file()


def test_desktop_inspection_rejects_plans_and_uses_the_shared_inventory_worker(
    desktop, tmp_path, monkeypatch
):
    jobs: list = []
    feedback = Mock()
    monkeypatch.setattr(desktop, "_begin_job", jobs.append)
    monkeypatch.setattr(desktop, "_feedback", feedback)
    desktop._inspect_sources()
    assert not jobs
    assert "at least one source" in feedback.call_args.args[0]
    plan = tmp_path / "intake.json"
    plan.write_text("[]", encoding="utf-8")
    desktop.sources["intake_plan"].entry.setText(str(plan))
    desktop._inspect_sources()
    assert not jobs
    assert "already inspected" in feedback.call_args.args[0]
    desktop.sources["intake_plan"].entry.clear()
    folder = mixed_folder(tmp_path / "mixed")
    desktop.sources["auto_dir"].entry.setText(str(folder))
    desktop._busy = True
    desktop._inspect_sources()
    assert not jobs
    desktop._busy = False
    desktop.inspect_button.click()
    assert len(jobs) == 1
    assert jobs[0].mode == "intake"
    assert not jobs[0].formats
    worker = RunWorker(jobs[0])
    outcomes: list = []
    worker.outcome_ready.connect(outcomes.append)
    worker.run()
    assert outcomes[0].status == "complete"
    inventory = json.loads((Path(jobs[0].directory) / "intake.json").read_text(encoding="utf-8"))
    assert sum(row["status"] == "ready" for row in inventory) == 3
