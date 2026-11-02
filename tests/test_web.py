"""HTTP boundaries and durable review share independently expected core results."""

import hashlib
import io
import json
import os
import subprocess
import sys
import time
from pathlib import Path
from unittest.mock import Mock
from zipfile import ZipFile

import pytest
from fastapi.testclient import TestClient

from edf_bill_fetcher.application.review_commands import apply_command
from edf_bill_fetcher.application.review_session import ReviewSession
from edf_bill_fetcher.domain.json_values import object_value
from edf_bill_fetcher.io.persistence.result_store import load_result, save_result
from edf_bill_fetcher.web.api import Dispatcher, allowed_path, create_app
from edf_bill_fetcher.web.contracts import Job, Source
from edf_bill_fetcher.web.storage import Conflict, Store, contained
from edf_bill_fetcher.web.worker import Cancellation, execute, source_path, unpack_bundle
from tests.synthetic_analysis import snapshot

PASSWORD = "correct-horse-evidence"


@pytest.fixture
def web(tmp_path):
    app = create_app(
        tmp_path / "server", roots={"bills": str(tmp_path / "bills")}, start_worker=False
    )
    (tmp_path / "bills").mkdir()
    store = app.state.store
    store.user("reviewer", PASSWORD)
    with TestClient(app) as client:
        response = client.post("/api/login", json={"username": "reviewer", "password": PASSWORD})
        assert response.status_code == 200
        client.headers["x-csrf-token"] = response.json()["csrf"]
        yield client, store
    store.engine.dispose()


def accepted(store, *, name="Synthetic"):
    entry = store.create(
        "job",
        {
            "status": "complete",
            "cancel": False,
            "created": 1,
            "request": {"name": name, "mode": "extract", "formats": ["xlsx"]},
            "error": None,
        },
    )
    result = snapshot(
        [
            {
                "Date": "01/02/2024",
                "Entry Type": "New Bill",
                "Invoice #": "OLD",
                "Account": "A-123",
                "Fuel": "Electricity",
                "Period From": "01/01/2024",
                "Period To": "31/01/2024",
                "Period Charge (£)": 100,
                "Balance (£)": 100,
            },
            {
                "Date": "01/03/2024",
                "Entry Type": "New Bill",
                "Invoice #": "NEW",
                "Account": "A-123",
                "Fuel": "Electricity",
                "Period From": "01/01/2024",
                "Period To": "31/01/2024",
                "Period Charge (£)": 150,
                "Balance (£)": 250,
            },
        ]
    )
    save_result(result, store.directory(entry) / "result.json", result_id=entry["id"])
    return entry, result


@pytest.mark.parametrize(
    "path", ["../outside", "/etc/passwd", "C:/secret", "a\\b", "", ".", "a/../../b"]
)
def test_paths_cannot_escape_managed_root(tmp_path, path):
    with pytest.raises(ValueError):
        contained(tmp_path, path)


def test_catalogue_is_durable_and_conflicting_revisions_do_not_overwrite(tmp_path):
    first = Store(tmp_path)
    entry = first.create("test", {"value": 1})
    saved = first.replace(entry, {"value": 2})
    with pytest.raises(Conflict):
        first.replace(entry, {"value": 3})
    second = Store(tmp_path)
    assert second.get(saved["id"])["data"] == {"value": 2}
    with pytest.raises(KeyError):
        second.get(saved["id"], "job")
    first.log(entry["id"], "First stage")
    first.log("other", "Unrelated stage")
    first.log(entry["id"], "Second stage")
    events = second.events(entry["id"])
    assert [event["message"] for event in events] == ["First stage", "Second stage"]
    assert [event["message"] for event in second.events(entry["id"], events[0]["sequence"])] == [
        "Second stage"
    ]
    first.engine.dispose()
    second.engine.dispose()


def test_authentication_csrf_origin_logout_and_expiry(web):
    client, store = web
    assert client.get("/api/session").json()["username"] == "reviewer"
    assert (
        client.post(
            "/api/assets",
            json={"name": "Test", "kind": "pdf_dir"},
            headers={"x-csrf-token": "wrong"},
        ).status_code
        == 403
    )
    assert (
        client.post(
            "/api/assets",
            json={"name": "Test", "kind": "pdf_dir"},
            headers={"origin": "https://elsewhere.invalid"},
        ).status_code
        == 403
    )
    assert client.get("/api/openapi.json").status_code == 200
    assert client.get("/api/schema").json()["review_fields"]["Period Charge (£)"] == "money"
    assert client.post("/api/logout", json={}).status_code == 200
    assert client.get("/api/jobs").status_code == 401
    assert (
        client.post("/api/login", json={"username": "nobody", "password": PASSWORD}).status_code
        == 401
    )
    assert (
        client.post("/api/login", json={"username": "reviewer", "password": "wrong"}).status_code
        == 401
    )
    with pytest.raises(ValueError):
        store.user("empty", "")


def test_signin_rate_is_bounded(web):
    client, _ = web
    for _ in range(10):
        assert (
            client.post(
                "/api/login", json={"username": "reviewer", "password": "wrong"}
            ).status_code
            == 401
        )
    assert (
        client.post("/api/login", json={"username": "reviewer", "password": PASSWORD}).status_code
        == 429
    )


def test_upload_resume_offsets_sealing_and_input_validation(web):
    client, store = web
    entry = client.post("/api/assets", json={"kind": "pdf_dir", "name": "Bills"}).json()
    identity = entry["id"]
    url = f"/api/assets/{identity}/content"
    assert client.get(
        f"/api/assets/{identity}/offset", params={"path": "2024/bill.pdf"}
    ).json() == {"offset": 0, "sha256": None}
    assert client.put(url, params={"path": "2024/bill.pdf"}, content=b"first").json() == {
        "offset": 5
    }
    assert (
        client.put(
            url, params={"path": "2024/bill.pdf", "offset": 0}, content=b"duplicate"
        ).status_code
        == 409
    )
    assert (
        client.put(url, params={"path": "2024/bill.pdf", "offset": 5}, content=b"second").json()[
            "offset"
        ]
        == 11
    )
    assert (
        client.get(f"/api/assets/{identity}/offset", params={"path": "2024/bill.pdf"}).json()[
            "offset"
        ]
        == 11
    )
    assert client.put(url, params={"path": "../escape"}, content=b"secret").status_code == 400
    assert client.post(f"/api/assets/{identity}/seal", json={}).status_code == 200
    assert client.put(url, params={"path": "other.pdf"}, content=b"changed").status_code == 409
    assert (
        source_path(store, Source(kind="pdf_dir", asset_id=identity), {})
        .joinpath("2024/bill.pdf")
        .read_bytes()
        == b"firstsecond"
    )
    assert len(client.get("/api/assets").json()) == 1
    with pytest.raises(ValueError, match="wrong source type"):
        source_path(store, Source(kind="pst", asset_id=identity), {})
    assert client.post("/api/assets", json={"kind": "unknown", "name": "bad"}).status_code == 422


@pytest.mark.parametrize(
    ("kind", "files"), [("pst", []), ("pst", ["a.pst", "b.pst"]), ("pdf_dir", [])]
)
def test_invalid_upload_collection_is_not_finalized(web, kind, files):
    client, store = web
    entry = store.create("asset", {"kind": kind, "name": "Input", "sealed": False, "files": files})
    assert client.post(f"/api/assets/{entry['id']}/seal", json={}).status_code == 400
    with pytest.raises(ValueError, match="incomplete"):
        source_path(store, Source(kind=kind, asset_id=entry["id"]), {})


@pytest.mark.parametrize(
    "payload",
    [
        {"mode": "extract"},
        {"mode": "render"},
        {"mode": "diff", "parent_id": "one"},
        {"mode": "preview", "parent_id": "one"},
        {"mode": "review", "parent_id": "one"},
        {"mode": "extract", "sources": [{"kind": "result", "asset_id": "one"}]},
        {"mode": "render", "parent_id": "one", "formats": ["html", "html"]},
        {"mode": "render", "parent_id": "one", "sections": ["unknown"]},
        {"mode": "render", "parent_id": "one", "sections": ["payment", "payment"]},
        *[
            {"mode": "render", "parent_id": "one", "config": config}
            for config in [
                {"misspelled": True},
                {"use_dedup": "yes"},
                {"min_amount": -1},
                {"credit_hold_days": 1.5},
                {"credit_hold_days": -1},
                {"acc_num": 123},
                {"as_of": "20250301"},
                {"as_of": "2025-02-30"},
            ]
        ],
    ],
)
def test_invalid_captured_jobs_fail_before_queueing(payload):
    with pytest.raises(ValueError):
        Job.model_validate(payload)


def test_server_roots_browsing_and_traversal(web, tmp_path):
    client, store = web
    path = tmp_path / "bills" / "2024"
    path.mkdir()
    (path / "bill.pdf").write_bytes(b"source")
    assert client.get("/api/roots").json() == [{"id": "bills", "name": "bills"}]
    assert client.get("/api/roots/bills/browse").json()[0]["directory"] is True
    assert (
        client.get("/api/roots/bills/browse", params={"relative": "2024"}).json()[0]["name"]
        == "bill.pdf"
    )
    assert client.get("/api/roots/unknown/browse").status_code == 404
    assert client.get("/api/roots/bills/browse", params={"relative": "../"}).status_code == 400
    source = Source(kind="pdf_dir", root_id="bills", relative="2024")
    assert source_path(store, source, {"bills": str(tmp_path / "bills")}) == path
    with pytest.raises(ValueError, match="Unknown source"):
        source_path(store, source, {})


def test_job_queue_cancel_and_durable_progress_replay(web):
    client, store = web
    parent, _ = accepted(store)
    response = client.post(
        "/api/jobs", json={"mode": "render", "parent_id": parent["id"], "formats": ["html"]}
    )
    assert response.status_code == 202
    identity = response.json()["id"]
    assert client.get("/api/jobs/" + identity).json()["data"]["status"] == "queued"
    assert client.get("/api/jobs").status_code == 200
    token = Cancellation(store, identity)
    assert not token.is_set()
    store.log(identity, "Reading original evidence")
    response = client.post(f"/api/jobs/{identity}/cancel", json={})
    assert response.json()["data"]["status"] == "cancelled"
    assert token.is_set()
    events = client.get(f"/api/jobs/{identity}/events")
    assert "Reading original evidence" in events.text
    assert "event: finished" in events.text
    cursor = store.events(identity)[0]["sequence"]
    assert (
        "Reading original evidence"
        not in client.get(
            f"/api/jobs/{identity}/events", headers={"last-event-id": str(cursor)}
        ).text
    )
    assert (
        client.get(f"/api/jobs/{identity}/events", headers={"last-event-id": "bad"}).status_code
        == 400
    )


def test_review_edits_conflicts_undo_redo_and_publication_binding(web):
    client, store = web
    parent, baseline = accepted(store)
    url = f"/api/jobs/{parent['id']}/review"
    initial = client.get(url).json()
    row = initial["rows"][0]["id"]
    command = {
        "revision": 0,
        "operation": "update",
        "arguments": {
            "row_id": row,
            "changes": {"Period Charge (£)": "80.00"},
            "include": True,
            "reason": "Checked the original",
            "status": "verified",
        },
    }
    edited = client.post(url, json=command)
    assert edited.status_code == 200, edited.text
    assert edited.json()["revision"] == 1
    assert edited.json()["summary"]["charges"] == "250.00"
    assert edited.json()["rows"][0]["record"]["Period Charge (£)"] == 80
    assert client.post(url, json=command).status_code == 409
    undo = client.post(url, json={"revision": 1, "operation": "undo"}).json()
    assert undo["rows"][0]["record"]["Period Charge (£)"] == 100
    redo = client.post(url, json={"revision": 2, "operation": "redo"}).json()
    assert redo["rows"][0]["record"]["Period Charge (£)"] == 80
    assert client.post(url, json={"revision": 3, "operation": "not_a_method"}).status_code == 400
    assert (
        client.post(url, json={"revision": 3, "operation": "update", "arguments": {}}).status_code
        == 400
    )
    assert load_result(store.directory(parent) / "result.json") == baseline
    queued = client.post(
        "/api/jobs",
        json={
            "mode": "preview",
            "parent_id": parent["id"],
            "draft": redo["draft"],
            "formats": ["html"],
        },
    ).json()
    execute(store, queued["id"], {})
    assert store.get(queued["id"])["data"]["status"] == "complete"
    preview = load_result(store.directory(queued) / "result.json")
    assert str(preview.summary.charges) == "230.00"
    assert load_result(store.directory(parent) / "result.json") == baseline


def test_pack_downloads_are_contained_and_preserve_original_bytes(web):
    client, store = web
    parent, _ = accepted(store)
    directory = store.directory(parent)
    (directory / "evidence_files").mkdir()
    (directory / "evidence_files" / "bill.pdf").write_bytes(b"original")
    identity = parent["id"]
    paths = {item["path"] for item in client.get(f"/api/jobs/{identity}/artifacts").json()}
    assert paths == {"result.json", "evidence_files/bill.pdf"}
    assert (
        client.get(
            f"/api/jobs/{identity}/download", params={"path": "../catalogue.sqlite"}
        ).status_code
        == 400
    )
    assert (
        client.get(f"/api/jobs/{identity}/download", params={"path": "missing"}).status_code == 404
    )
    downloaded = client.get(
        f"/api/jobs/{identity}/download", params={"path": "evidence_files/bill.pdf"}
    )
    assert downloaded.content == b"original"
    with ZipFile(io.BytesIO(client.get(f"/api/jobs/{identity}/pack").content)) as archive:
        assert archive.read("evidence_files/bill.pdf") == b"original"
        assert json.loads(archive.read("result.json"))["schema_version"] == 3
    store.patch(identity, status="running")
    assert client.get(f"/api/jobs/{identity}/pack").status_code == 409
    assert client.get("/api/jobs/unknown").status_code == 404


@pytest.mark.parametrize("entry", ["../secret", "/absolute", "C:/secret", "a\\secret"])
def test_portable_pack_import_rejects_escaping_members(tmp_path, entry):
    path = tmp_path / "input.zip"
    with ZipFile(path, "w") as archive:
        archive.writestr(entry, b"secret")
    with pytest.raises(ValueError):
        unpack_bundle(path, tmp_path / "expanded")


def test_worker_cancel_failure_diff_and_saved_rerender(web):
    client, store = web
    parent, baseline = accepted(store)
    second, _ = accepted(store, name="Later")
    queued = client.post(
        "/api/jobs", json={"mode": "diff", "parent_id": parent["id"], "other_id": second["id"]}
    ).json()
    execute(store, queued["id"], {})
    assert store.get(queued["id"])["data"]["status"] == "complete"
    assert json.loads(
        (store.directory(queued) / "comparison.json").read_text(encoding="utf-8")
    ) == {
        "added": [],
        "removed": [],
        "changed": [],
    }
    queued = client.post(
        "/api/jobs", json={"mode": "render", "parent_id": parent["id"], "formats": ["html"]}
    ).json()
    execute(store, queued["id"], {})
    assert store.get(queued["id"])["data"]["status"] == "complete"
    assert load_result(store.directory(queued) / "result.json") == baseline
    assert (store.directory(queued) / "Evidence_Report.html").is_file()
    cancelled = client.post("/api/jobs", json={"mode": "render", "parent_id": parent["id"]}).json()
    client.post(f"/api/jobs/{cancelled['id']}/cancel", json={})
    execute(store, cancelled["id"], {})
    assert store.get(cancelled["id"])["data"]["status"] == "cancelled"
    failed = store.create(
        "job",
        {
            "request": {"mode": "render", "parent_id": "missing"},
            "status": "running",
            "cancel": False,
        },
    )
    execute(store, failed["id"], {})
    assert store.get(failed["id"])["data"]["status"] == "failed"


def test_review_command_rejects_private_methods_without_mutation():
    session = ReviewSession(snapshot([]))
    original = session.payload()
    with pytest.raises(ValueError):
        apply_command(session, "_commit", {"updates": []})
    assert session.payload() == original


def test_real_pdf_http_extraction_review_and_portable_pack(web):
    client, store = web
    asset = client.post("/api/assets", json={"kind": "pdf_dir", "name": "Synthetic bills"}).json()
    original = (Path(__file__).parent / "fixtures/sample_bill.pdf").read_bytes()
    assert (
        client.put(
            f"/api/assets/{asset['id']}/content", params={"path": "bill.pdf"}, content=original
        ).status_code
        == 200
    )
    client.post(f"/api/assets/{asset['id']}/seal", json={})
    queued = client.post(
        "/api/jobs",
        json={
            "mode": "extract",
            "sources": [{"kind": "pdf_dir", "asset_id": asset["id"]}],
            "formats": ["html"],
            "config": {"analysis_min": 0, "filter_below": False, "as_of": "2025-03-01"},
        },
    ).json()
    execute(store, queued["id"], {})
    assert store.get(queued["id"])["data"]["status"] in {"complete", "review_required"}
    result = load_result(store.directory(queued) / "result.json")
    assert result.prepared.extraction.observations
    sources = client.get(f"/api/jobs/{queued['id']}/sources")
    assert sources.status_code == 200, sources.text
    available = [
        source for entry in sources.json() for source in entry["sources"] if source["available"]
    ]
    assert available
    downloaded = client.get(
        f"/api/jobs/{queued['id']}/source", params={"source_id": available[0]["source_id"]}
    )
    assert downloaded.content == original
    assert downloaded.headers["content-type"] == "application/pdf"
    assert (
        client.get(f"/api/jobs/{queued['id']}/source", params={"source_id": "unknown"}).status_code
        == 404
    )
    workbook = store.directory(queued) / "Evidence.xlsx"
    imported = store.create(
        "asset",
        {"kind": "workbook", "name": "Reviewed Excel", "sealed": True, "files": ["review.xlsx"]},
    )
    (store.directory(imported) / "review.xlsx").write_bytes(workbook.read_bytes())
    reviewed = client.post(
        "/api/jobs",
        json={
            "mode": "workbook",
            "input": {"kind": "workbook", "asset_id": imported["id"]},
            "formats": ["html"],
        },
    ).json()
    execute(store, reviewed["id"], {})
    assert store.get(reviewed["id"])["data"]["status"] in {"complete", "review_required"}
    assert load_result(store.directory(reviewed) / "result.json").summary == result.summary


def test_bundle_import_and_malformed_archive_boundaries(tmp_path):
    source = tmp_path / "pack.zip"
    with ZipFile(source, "w") as archive:
        archive.writestr("pack/", b"")
        archive.writestr("pack/result.json", b"{}")
        archive.writestr("pack/evidence_files/a.pdf", b"original")
    result = unpack_bundle(source, tmp_path / "imported")
    assert result == tmp_path / "imported/pack/result.json"
    assert (result.parent / "evidence_files/a.pdf").read_bytes() == b"original"
    with pytest.raises(ValueError, match="limit"):
        unpack_bundle(source, tmp_path / "limited", max_bytes=1)
    with ZipFile(source, "w") as archive:
        archive.writestr("a.txt", b"no result")
    with pytest.raises(ValueError, match="exactly one"):
        unpack_bundle(source, tmp_path / "missing")
    from zipfile import ZipInfo

    info = ZipInfo("link")
    info.external_attr = 0o120777 << 16
    with ZipFile(source, "w") as archive:
        archive.writestr(info, b"/etc/passwd")
    with pytest.raises(ValueError, match="symbolic"):
        unpack_bundle(source, tmp_path / "symlink")


def test_expired_sessions_and_limits_are_not_bypassed(web):
    client, store = web
    assert (
        client.post(
            "/api/assets", content=b"{}", headers={"content-length": str(9 * 1024 * 1024)}
        ).status_code
        == 413
    )
    session = store.session(client.cookies["edf_session"])
    from sqlalchemy import update

    from edf_bill_fetcher.web.schema import sessions

    with store.engine.begin() as connection:
        connection.execute(update(sessions).where(sessions.c.id == session["id"]).values(expires=0))
    assert client.get("/api/jobs").status_code == 401
    assert not allowed_path(store, {}, str(store.root.parent / "private.pdf"))
    assert allowed_path(store, {}, str(store.root / "job"))
    assert allowed_path(
        store,
        {"shared": str(store.root.parent / "shared")},
        str(store.root.parent / "shared/bill.pdf"),
    )


@pytest.mark.parametrize("mode", ["crash", "launch_failure", "shutdown", "cancel"])
def test_dispatcher_native_crashes_and_shutdown_never_report_success(tmp_path, monkeypatch, mode):
    import edf_bill_fetcher.web.api as api_module

    store = Store(tmp_path)
    job = store.create("job", {"status": "queued", "cancel": False})
    process = Mock(returncode=9)
    process.poll.side_effect = [None, 9] if mode in {"shutdown", "cancel"} else [9]
    if mode in {"shutdown", "cancel"}:
        process.wait.side_effect = [subprocess.TimeoutExpired("worker", 10), None]
    launch = Mock(
        side_effect=OSError("cannot start worker") if mode == "launch_failure" else None,
        return_value=process,
    )
    monkeypatch.setattr(api_module.subprocess, "Popen", launch)
    dispatcher = Dispatcher(store)
    dispatcher.thread.start()
    deadline = time.monotonic() + 5
    while time.monotonic() < deadline:
        status = store.get(job["id"])["data"]["status"]
        if mode == "shutdown" and status == "running":
            dispatcher.stop.set()
        if mode == "cancel" and status == "running":
            store.patch(job["id"], cancel=True)
        if status in {"interrupted", "failed", "cancelled"}:
            break
        time.sleep(0.02)
    dispatcher.stop.set()
    dispatcher.thread.join(3)
    assert store.get(job["id"])["data"]["status"] == (
        "failed" if mode == "launch_failure" else "cancelled" if mode == "cancel" else "interrupted"
    )
    if mode in {"shutdown", "cancel"}:
        process.kill.assert_called_once()
    assert not dispatcher.thread.is_alive()
    store.engine.dispose()


def test_server_startup_marks_old_running_jobs_interrupted_and_locks_workspace(tmp_path):
    from filelock import Timeout

    app = create_app(tmp_path)
    job = app.state.store.create("job", {"status": "running", "cancel": False})
    with TestClient(app):
        assert app.state.store.get(job["id"])["data"]["status"] == "interrupted"
        assert json.loads((tmp_path / "roots.json").read_text(encoding="utf-8")) == {}
        second = create_app(tmp_path)
        with pytest.raises(Timeout), TestClient(second):
            pass
        second.state.store.engine.dispose()
    app.state.store.engine.dispose()


def test_server_cli_creates_accounts_and_captures_explicit_roots(tmp_path, monkeypatch):
    import uvicorn

    from edf_bill_fetcher.web.launch import main

    monkeypatch.setenv("EDF_WEB_PASSWORD", PASSWORD)
    main(["--data", str(tmp_path / "data"), "--init-user", "reviewer"])
    evidence = tmp_path / "bills"
    evidence.mkdir()
    server = Mock()
    monkeypatch.setattr(uvicorn, "run", server)
    main(
        [
            "--data",
            str(tmp_path / "data"),
            "--host",
            "0.0.0.0",
            "--port",
            "8765",
            "--source-root",
            "bills=" + str(evidence),
        ]
    )
    assert server.call_args.kwargs["host"] == "0.0.0.0"
    assert server.call_args.kwargs["port"] == 8765
    assert server.call_args.kwargs["workers"] == 1
    for spec in ["missing-equals", "same=" + str(tmp_path / "missing")]:
        with pytest.raises(SystemExit):
            main(["--data", str(tmp_path / "data"), "--source-root", spec])
    main(["--data", str(tmp_path / "empty")])
    assert server.call_args.args[0].state.store.root == (tmp_path / "empty").resolve()


def test_worker_module_launch_executes_durable_queued_request(tmp_path):
    store = Store(tmp_path)
    job = store.create(
        "job",
        {"status": "queued", "cancel": True, "request": {"mode": "render", "parent_id": "missing"}},
    )
    environment = {**os.environ, "QT_QPA_PLATFORM": "offscreen"}
    completed = subprocess.run(
        [
            sys.executable,
            "-m",
            "edf_bill_fetcher.web.worker",
            "--data",
            str(tmp_path),
            "--job",
            job["id"],
        ],
        capture_output=True,
        text=True,
        timeout=30,
        env=environment,
    )
    assert completed.returncode == 0, completed.stderr
    assert store.get(job["id"])["data"]["status"] == "cancelled"
    store.engine.dispose()


def test_verified_upload_hashes_reject_changed_originals(web):
    client, store = web
    asset = client.post("/api/assets", json={"kind": "pst", "name": "Archive"}).json()
    identity = asset["id"]
    client.put(
        f"/api/assets/{identity}/content", params={"path": "archive.pst"}, content=b"original"
    )
    offset = client.get(f"/api/assets/{identity}/offset", params={"path": "archive.pst"}).json()
    assert offset == {"offset": 8, "sha256": hashlib.sha256(b"original").hexdigest()}
    assert (
        client.post(
            f"/api/assets/{identity}/seal",
            json={"files": {"other.pst": {"size": 8, "sha256": "0" * 64}}},
        ).status_code
        == 400
    )
    assert (
        client.post(
            f"/api/assets/{identity}/seal",
            json={"files": {"archive.pst": {"size": 8, "sha256": "0" * 64}}},
        ).status_code
        == 400
    )
    assert not store.get(identity)["data"]["sealed"]
    assert (
        client.post(
            f"/api/assets/{identity}/seal",
            json={"files": {"archive.pst": {"size": 8, "sha256": offset["sha256"]}}},
        ).status_code
        == 200
    )


def test_chunked_json_cannot_bypass_body_limit(web):
    client, _ = web

    def body():
        yield b"{" + b" " * (8 * 1024 * 1024)
        yield b"}"

    response = client.post("/api/assets", content=body(), headers={"transfer-encoding": "chunked"})
    assert response.status_code == 413


def test_shared_review_execution_matches_cli_and_web_and_preserves_inputs(web, tmp_path):
    from edf_bill_fetcher.application.review_cli import main
    from edf_bill_fetcher.application.review_execution import preview_review, publish_review

    client, store = web
    parent, baseline = accepted(store)
    result_path = store.directory(parent) / "result.json"
    original = result_path.read_bytes()
    session = ReviewSession(baseline)
    row = session.rows[0]["id"]
    command = {
        "operation": "update",
        "arguments": {
            "row_id": row,
            "changes": {"Period Charge (£)": "80.00"},
            "include": True,
            "reason": "Original checked",
            "status": "verified",
        },
    }
    command_file = tmp_path / "command.json"
    command_file.write_text(json.dumps(command))
    draft_file = tmp_path / "draft.json"
    main(
        [
            "--result",
            str(result_path),
            "--draft",
            str(draft_file),
            "--command",
            str(command_file),
            "--preview",
            str(tmp_path / "preview.json"),
            "--publish",
            str(tmp_path / "outputs"),
            "--formats",
            "html",
        ]
    )
    cli_preview = load_result(tmp_path / "preview.json")
    assert str(cli_preview.summary.charges) == "230.00"
    draft = json.loads(draft_file.read_text(encoding="utf-8"))
    assert preview_review(str(result_path), draft).summary == cli_preview.summary
    published = list((tmp_path / "outputs").glob("reviewed_*/result.json"))
    assert len(published) == 1
    assert load_result(published[0]).summary == cli_preview.summary
    with pytest.raises(ValueError, match="fresh empty"):
        publish_review(str(result_path), draft, str(published[0].parent), formats=("html",))
    queued = client.post(
        "/api/jobs",
        json={"mode": "review", "parent_id": parent["id"], "draft": draft, "formats": ["html"]},
    ).json()
    execute(store, queued["id"], {})
    assert store.get(queued["id"])["data"]["status"] in {"complete", "review_required"}
    assert load_result(store.directory(queued) / "result.json").summary == cli_preview.summary
    assert result_path.read_bytes() == original


def test_rerender_analytical_overrides_are_explicit_and_leave_baseline_unchanged(web):
    client, store = web
    parent, baseline = accepted(store)
    queued = client.post(
        "/api/jobs",
        json={
            "mode": "render",
            "parent_id": parent["id"],
            "formats": ["html"],
            "config": {"analysis_min": 1000},
        },
    ).json()
    execute(store, queued["id"], {})
    result = load_result(store.directory(queued) / "result.json")
    assert object_value(result.prepared.effective_config)["analysis_min"] == 1000
    assert result.prepared.extraction.request.as_of == baseline.prepared.extraction.request.as_of
    assert load_result(store.directory(parent) / "result.json") == baseline


def test_folder_uploads_use_per_file_state_and_verify_the_complete_selection(web):
    client, store = web
    asset = client.post("/api/assets", json={"kind": "auto", "name": "Everything"}).json()
    identity = asset["id"]
    url = f"/api/assets/{identity}"
    for path, data, order in [("year/a.pdf", b"first", 0), ("year/b.eml", b"second", 1)]:
        assert client.put(url + "/content", params={"path": path}, content=data).status_code == 200
        revision = store.get(identity)["revision"]
        assert (
            client.put(
                url + "/content", params={"path": path, "offset": len(data)}, content=b"!"
            ).status_code
            == 200
        )
        assert store.get(identity)["revision"] == revision
        assert store.get(identity)["data"]["files"] == []
        digest = hashlib.sha256(data + b"!").hexdigest()
        bad = client.post(
            url + "/file-complete",
            params={"path": path},
            json={"size": len(data) + 1, "sha256": "0" * 64, "batch_id": "batch", "order": order},
        )
        assert bad.status_code == 409
        good = client.post(
            url + "/file-complete",
            params={"path": path},
            json={"size": len(data) + 1, "sha256": digest, "batch_id": "batch", "order": order},
        )
        assert good.status_code == 200, good.text
    assert (
        client.post(url + "/seal", json={"batch_id": "batch", "files_count": 1}).status_code == 409
    )
    assert not store.get(identity)["data"]["sealed"]
    assert (
        client.post(url + "/seal", json={"batch_id": "batch", "files_count": 2}).status_code == 200
    )
    assert store.get(identity)["data"]["files"] == ["year/a.pdf", "year/b.eml"]
    assert len(store.upload_files(identity)) == 2
    assert "files" not in client.get("/api/assets").json()[0]["data"]
    assert (
        client.post(
            url + "/file-complete",
            params={"path": "year/a.pdf"},
            json={
                "size": 6,
                "sha256": hashlib.sha256(b"first!").hexdigest(),
                "batch_id": "batch",
                "order": 0,
            },
        ).status_code
        == 200
    )


def test_mixed_intake_preview_is_paginated_and_reuses_expanded_zip_members(
    web, tmp_path, monkeypatch
):
    import edf_bill_fetcher.application.source_intake as intake_module
    from tests.test_mixed_intake import mixed_folder

    client, store = web
    folder = mixed_folder(tmp_path / "bills" / "mixed")
    parent, _ = accepted(store)
    (folder / "earlier.json").write_bytes((store.directory(parent) / "result.json").read_bytes())
    inspected = client.post(
        "/api/jobs",
        json={
            "mode": "intake",
            "sources": [{"kind": "auto", "root_id": "bills", "relative": "mixed"}],
        },
    ).json()
    execute(store, inspected["id"], {"bills": str(tmp_path / "bills")})
    assert store.get(inspected["id"])["data"]["status"] == "complete"
    response = client.get(f"/api/jobs/{inspected['id']}/intake", params={"limit": 2}).json()
    assert len(response["items"]) == 2
    assert response["counts"]["ready"] == 3
    assert response["counts"]["unsupported"] == 1
    assert response["counts"]["needs_choice"] == 1
    assert all("path" not in row for row in response["items"])
    choice = client.get(
        f"/api/jobs/{inspected['id']}/intake", params={"status": "needs_choice"}
    ).json()["items"][0]
    imported = client.post(
        "/api/jobs",
        json={
            "mode": "render",
            "input": {"kind": "auto", "intake_id": inspected["id"], "relative": choice["id"]},
            "formats": ["html"],
        },
    ).json()
    execute(store, imported["id"], {"bills": str(tmp_path / "bills")})
    assert (
        load_result(store.directory(imported) / "result.json").summary
        == load_result(store.directory(parent) / "result.json").summary
    )
    monkeypatch.setattr(
        intake_module, "discover", Mock(side_effect=AssertionError("Do not expand archives twice"))
    )
    queued = client.post(
        "/api/jobs",
        json={
            "mode": "extract",
            "sources": [{"kind": "auto", "intake_id": inspected["id"], "relative": "__all__"}],
            "formats": ["html"],
            "config": {
                "as_of": "2025-03-01",
                "analysis_min": 0,
                "filter_below": False,
                "use_dedup": False,
                "use_domain_filter": False,
            },
        },
    ).json()
    execute(store, queued["id"], {"bills": str(tmp_path / "bills")})
    assert store.get(queued["id"])["data"]["status"] == "review_required"
    assert load_result(store.directory(queued) / "result.json").summary.charges == 180


@pytest.mark.parametrize("exhausted", [False, True])
def test_worker_status_retries_preserve_concurrent_cancellation(tmp_path, monkeypatch, exhausted):
    store = Store(tmp_path)
    job = store.create("job", {"status": "running", "cancel": False})
    replace = store.replace
    calls = 0

    def racing(entry, data):
        nonlocal calls
        calls += 1
        if exhausted or calls == 1:
            latest = store.get(job["id"])
            replace(latest, {**latest["data"], "cancel": True})
        return replace(entry, data)

    monkeypatch.setattr(store, "replace", racing)
    if exhausted:
        with pytest.raises(Conflict):
            store.patch(job["id"], status="cancelled")
        assert calls == 8
    else:
        updated = store.patch(job["id"], status="cancelled")
        assert updated["data"] == {"status": "cancelled", "cancel": True}
        assert calls == 2
    store.engine.dispose()


@pytest.mark.parametrize(
    ("method", "suffix"),
    [("get", "offset"), ("put", "content"), ("post", "file-complete"), ("post", "seal")],
)
def test_upload_lock_names_cannot_escape_on_windows(web, method, suffix):
    client, _ = web
    response = getattr(client, method)(
        "/api/assets/..%5C..%5Coutside/" + suffix,
        params={"path": "bill.pdf"},
        **(
            {"json": {"size": 0, "sha256": "0" * 64, "batch_id": "batch", "order": 0}}
            if suffix == "file-complete"
            else {"json": {}}
            if method == "post"
            else {}
        ),
    )
    assert response.status_code == 400


def test_legacy_absolute_source_mapping_cannot_escape_review_publication_policy(tmp_path):
    from edf_bill_fetcher.application.workbook_review import bundle_review_sources

    original = tmp_path / "private.pdf"
    contents = b"%PDF-1.7 private original"
    original.write_bytes(contents)
    result = snapshot(
        [
            {
                "Date": "01/02/2024",
                "Entry Type": "New Bill",
                "Period Charge (£)": 100,
                "Attachment Name": "private.pdf",
            }
        ],
        {"saved_evidence_paths": {"private.pdf": str(original)}},
    )
    directory = tmp_path / "public"
    directory.mkdir()
    bundle_review_sources(
        result.prepared,
        str(directory),
        str(tmp_path / "review.xlsx"),
        allowed_source=lambda path: False,
    )
    assert not list((directory / "evidence_files").glob("*.pdf"))
    assert original.read_bytes() == contents


def test_unsafe_custom_analysis_resource_is_rejected_without_mutating_accepted_result(
    web, tmp_path
):
    client, store = web
    parent, result = accepted(store)
    queued = client.post(
        "/api/jobs",
        json={
            "mode": "render",
            "parent_id": parent["id"],
            "config": {"ofgem_caps_path": str(tmp_path / "unapproved.csv")},
            "formats": ["html"],
        },
    ).json()
    execute(store, queued["id"], {})
    failed = store.get(queued["id"])["data"]
    assert failed["status"] == "failed"
    assert "inside a configured source folder" in failed["error"]
    assert load_result(store.directory(parent) / "result.json").summary == result.summary


def test_busy_upload_returns_a_resumable_conflict_instead_of_an_internal_error(web, monkeypatch):
    from unittest.mock import MagicMock

    from filelock import Timeout

    import edf_bill_fetcher.web.api as api_module

    client, _ = web
    asset = client.post("/api/assets", json={"kind": "pdf_dir", "name": "Input"}).json()
    lock = MagicMock()
    lock.__enter__.side_effect = Timeout("synthetic upload lock")
    monkeypatch.setattr(api_module, "FileLock", lambda *args, **kwargs: lock)
    response = client.get(f"/api/assets/{asset['id']}/offset", params={"path": "bill.pdf"})
    assert response.status_code == 409
    assert response.headers["retry-after"] == "1"
    assert "resume the same selection" in response.json()["detail"]


@pytest.mark.parametrize("failure", ["preflight", "disk_full", "quota", "io", "rollback"])
def test_failed_chunk_cannot_publish_partial_bytes_and_can_be_resumed(web, monkeypatch, failure):
    import errno
    from types import SimpleNamespace

    import edf_bill_fetcher.web.api as api_module

    client, store = web
    asset = client.post("/api/assets", json={"kind": "auto", "name": "Storage recovery"}).json()
    url = f"/api/assets/{asset['id']}"
    path = store.directory(asset) / "archive.pst"
    original = b"verified original prefix"
    chunk = b"next chunk"
    assert client.put(url + "/content?path=archive.pst", content=original).status_code == 200
    completion = {
        "size": len(original),
        "sha256": hashlib.sha256(original).hexdigest(),
        "batch_id": "disk-recovery",
        "order": 0,
    }
    assert client.post(url + "/file-complete?path=archive.pst", json=completion).status_code == 200
    actual_open = Path.open

    class FailingWrite:
        def __enter__(self):
            self.stream = actual_open(path, "ab", buffering=0)
            return self

        def write(self, value):
            self.stream.write(value[:3])
            raise OSError(
                errno.EIO
                if failure == "io"
                else getattr(errno, "EDQUOT", errno.ENOSPC)
                if failure == "quota"
                else errno.ENOSPC,
                "synthetic storage failure",
            )

        def __exit__(self, *args):
            self.stream.close()

    def fail_open(target, mode="r", *args, **kwargs):
        if target == path and mode == "ab":
            return FailingWrite()
        if target == path and mode == "r+b" and failure == "rollback":
            raise OSError(errno.EIO, "synthetic rollback failure")
        return actual_open(target, mode, *args, **kwargs)

    with monkeypatch.context() as patch:
        if failure == "preflight":
            patch.setattr(api_module.shutil, "disk_usage", lambda _: SimpleNamespace(free=0))
        else:
            patch.setattr(Path, "open", fail_open)
        response = client.put(
            url + "/content", params={"path": "archive.pst", "offset": len(original)}, content=chunk
        )
    assert response.status_code == (503 if failure == "io" else 507)
    if failure == "preflight":
        assert path.read_bytes() == original
        assert store.upload_files(asset["id"])[0]["data"]["complete"]
        return  # No write started; the earlier verified original is still valid.
    assert not store.upload_files(asset["id"])[0]["data"]["complete"]
    seal = {"batch_id": "disk-recovery", "files_count": 1}
    assert client.post(url + "/seal", json=seal).status_code == 409
    expected = original + (chunk[:3] if failure == "rollback" else b"")
    assert path.read_bytes() == expected
    offset = client.get(url + "/offset?path=archive.pst").json()
    assert offset == {"offset": len(expected), "sha256": hashlib.sha256(expected).hexdigest()}
    assert (
        client.put(
            url + "/content",
            params={"path": "archive.pst", "offset": len(expected)},
            content=(original + chunk)[len(expected) :],
        ).status_code
        == 200
    )
    completion.update(
        size=len(original + chunk), sha256=hashlib.sha256(original + chunk).hexdigest()
    )
    assert client.post(url + "/file-complete?path=archive.pst", json=completion).status_code == 200
    assert client.post(url + "/seal", json=seal).status_code == 200
    assert path.read_bytes() == original + chunk


def test_lost_final_confirmation_can_be_repeated_without_changing_a_sealed_source(web):
    client, store = web
    asset = client.post("/api/assets", json={"kind": "auto", "name": "Lost confirmation"}).json()
    url = f"/api/assets/{asset['id']}"
    completions = {}
    for order, (name, value) in enumerate([("empty.txt", b""), ("bill.pdf", b"original")]):
        assert client.put(url + "/content", params={"path": name}, content=value).status_code == 200
        completion = {
            "size": len(value),
            "sha256": hashlib.sha256(value).hexdigest(),
            "batch_id": "confirmation",
            "order": order,
        }
        completions[name] = completion
        assert (
            client.post(url + "/file-complete", params={"path": name}, json=completion).status_code
            == 200
        )
    seal = {"batch_id": "confirmation", "files_count": 2}
    assert client.post(url + "/seal", json=seal).status_code == 200
    before = store.get(asset["id"])
    files_before = store.upload_files(asset["id"])
    for name, completion in completions.items():
        assert (
            client.post(url + "/file-complete", params={"path": name}, json=completion).status_code
            == 200
        )
    assert client.post(url + "/seal", json=seal).status_code == 200
    assert store.get(asset["id"]) == before
    assert store.upload_files(asset["id"]) == files_before
    assert (
        client.post(url + "/seal", json={"batch_id": "different", "files_count": 2}).status_code
        == 409
    )
    assert (
        client.post(
            url + "/file-complete?path=bill.pdf",
            json={**completions["bill.pdf"], "batch_id": "different"},
        ).status_code
        == 409
    )
    assert (
        client.put(url + "/content?path=bill.pdf&offset=8", content=b"changed").status_code == 409
    )
    assert (store.directory(asset) / "bill.pdf").read_bytes() == b"original"
    # A file appearing outside the upload protocol must not extend a sealed manifest.
    (store.directory(asset) / "unexpected.pdf").write_bytes(b"extra")
    assert (
        client.post(
            url + "/file-complete?path=unexpected.pdf",
            json={
                "size": 5,
                "sha256": hashlib.sha256(b"extra").hexdigest(),
                "batch_id": "confirmation",
                "order": 2,
            },
        ).status_code
        == 409
    )
    assert store.get(asset["id"]) == before


@pytest.mark.parametrize(
    "case",
    ["saved", "unbundled", "changed", "missing", "forbidden", "collision", "existing_conflict"],
)
def test_rerender_preserves_relative_original_paths_and_never_substitutes_changed_copies(
    tmp_path, case
):
    from edf_bill_fetcher.web.worker import copy_accepted_sources

    source_root = tmp_path / "baseline"
    source_root.mkdir()
    original = tmp_path / "original.pdf"
    contents = b"%PDF-1.7 synthetic source bytes"
    original.write_bytes(contents)
    relative = "Evidence.xlsx" if case == "collision" else "evidence_files/A - original.pdf"
    metadata = {
        "Source ID": "source-one",
        "Attachment Name": "original.pdf",
        "Original Filename": "original.pdf",
        "Original Path": str(original),
        "Content Hash": hashlib.sha256(contents).hexdigest(),
    }
    if case not in {"unbundled", "forbidden"}:
        metadata["Saved Path"] = relative
        if case != "missing":
            saved = source_root / relative
            saved.parent.mkdir(parents=True, exist_ok=True)
            saved.write_bytes(b"modified" if case == "changed" else contents)
    result = snapshot(
        [
            {
                "Date": "01/02/2024",
                "Entry Type": "New Bill",
                "Period Charge (£)": 100,
                "Attachment Name": "original.pdf",
                "Source ID": "source-one",
            }
        ],
        {"source_manifest": {"source-one": metadata}},
    )
    output = tmp_path / "published"
    output.mkdir()
    if case == "existing_conflict":
        conflict = output / relative
        conflict.parent.mkdir(parents=True)
        conflict.write_bytes(b"different source")
    if case in {"collision", "existing_conflict"}:
        with pytest.raises(ValueError, match="collid|Conflicting"):
            copy_accepted_sources(result, source_root, output, lambda path: True)
        return
    paths, notes = copy_accepted_sources(
        result, source_root, output, lambda path: case != "forbidden"
    )
    if case in {"changed", "missing", "forbidden"}:
        assert not paths
        assert notes
        assert not list(output.rglob("*.pdf"))
    else:
        assert Path(paths["source-one"]).read_bytes() == contents
        if case == "saved":
            assert Path(paths["source-one"]).relative_to(output).as_posix() == relative
        assert not notes
    assert original.read_bytes() == contents
    assert result.summary.charges == 100


def test_intake_original_downloads_are_scoped_and_reject_changes_after_inspection(web, tmp_path):
    client, store = web
    job = store.create("job", {"status": "complete", "request": {"mode": "intake"}})
    original = tmp_path / "bills" / "original.pdf"
    original.write_bytes(b"%PDF-1.7 synthetic intake original")
    outside = tmp_path / "unapproved.pdf"
    outside.write_bytes(b"private bytes")
    directory = store.directory(job)
    (directory / "intake.json").write_text(
        json.dumps(
            [
                {
                    "id": "original",
                    "path": str(original),
                    "size": original.stat().st_size,
                    "mtime_ns": original.stat().st_mtime_ns,
                },
                {"id": "forbidden", "path": str(outside)},
                {"id": "folder", "path": str(original.parent)},
            ]
        ),
        encoding="utf-8",
    )
    url = f"/api/jobs/{job['id']}/intake-source"
    response = client.get(url, params={"item_id": "original"})
    assert response.content == original.read_bytes()
    assert response.headers["content-disposition"].startswith("attachment;")
    for identity in ["forbidden", "folder", "missing"]:
        assert client.get(url, params={"item_id": identity}).status_code == 404
    original.write_bytes(b"Changed original")
    assert client.get(url, params={"item_id": "original"}).status_code == 409


@pytest.mark.parametrize(
    "case", ["original", "saved", "excluded", "changed_saved", "missing", "forbidden", "no_hash"]
)
def test_single_original_lookup_checks_only_the_requested_identity(tmp_path, monkeypatch, case):
    import edf_bill_fetcher.application.evidence_review as review_module

    original = tmp_path / "original.pdf"
    original.write_bytes(b"%PDF-1.7 source one")
    saved = tmp_path / "evidence_files" / "original.pdf"
    saved.parent.mkdir()
    saved.write_bytes(b"changed" if case == "changed_saved" else original.read_bytes())
    metadata = {
        "Source ID": "one",
        "Attachment Name": "original.pdf",
        "Original Path": str(original),
        "Content Hash": hashlib.sha256(original.read_bytes()).hexdigest(),
    }
    if case in {"saved", "changed_saved"}:
        metadata["Saved Path"] = "evidence_files/original.pdf"
    if case == "no_hash":
        metadata.pop("Content Hash")
    state: dict = {
        "source_manifest": {
            "one": metadata,
            "unrelated": {"Original Path": str(tmp_path / "unrelated-private.pdf")},
        }
    }
    if case == "excluded":
        state = {
            "review_excluded_source_manifest": {"one": metadata},
            "review_source_paths": {"one": str(original)},
        }
    result = snapshot(
        [{"Date": "01/02/2024", "Entry Type": "New Bill", "Period Charge (£)": 100}], state
    )
    if case == "missing":
        original.unlink()
    actual_hash = review_module.artifact_record
    reads: list = []

    def checked(path):
        assert "unrelated-private" not in str(path)
        reads.append(str(path))
        return actual_hash(path)

    monkeypatch.setattr(review_module, "artifact_record", checked)
    value = review_module.resolve_review_source(
        result, str(tmp_path), "one", allowed_source=lambda path: case != "forbidden"
    )
    if case in {"changed_saved", "missing", "forbidden", "no_hash"}:
        assert value is None
    else:
        assert value is not None
        assert Path(value).read_bytes() == original.read_bytes()
    if case in {"forbidden", "no_hash"}:
        assert not reads
    assert review_module.resolve_review_source(result, str(tmp_path), "unknown") is None


@pytest.mark.parametrize("valid", [True, False])
def test_existing_prototype_catalogue_is_adopted_only_when_compatible(tmp_path, valid):
    import sqlite3

    database = sqlite3.connect(tmp_path / "catalogue.sqlite")
    columns = (
        "id TEXT PRIMARY KEY, kind TEXT NOT NULL, revision INTEGER NOT NULL, data JSON NOT NULL"
        if valid
        else "unexpected TEXT"
    )
    database.execute("CREATE TABLE entries (" + columns + ")")
    if valid:
        database.execute(
            "INSERT INTO entries VALUES (?,?,?,?)",
            ("old-job", "job", 3, json.dumps({"status": "complete"})),
        )
    database.commit()
    database.close()
    if valid:
        store = Store(tmp_path)
        assert store.get("old-job")["revision"] == 3
        assert store.get("old-job")["data"]["status"] == "complete"
        store.engine.dispose()
    else:
        with pytest.raises(ValueError, match="Unrecognised catalogue schema"):
            Store(tmp_path)


@pytest.mark.parametrize(
    "payload",
    [
        {"kind": "auto"},
        {"kind": "auto", "root_id": "root"},
        {"kind": "auto", "root_id": "root", "relative": "x", "asset_id": "asset"},
    ],
)
def test_ambiguous_source_locations_are_rejected(payload):
    with pytest.raises(ValueError):
        Source.model_validate(payload)


def test_captured_report_sections_use_shared_keys_and_review_requires_imported_parent():
    job = Job(mode="render", parent_id="parent", config={"report_sections": ["cover", "payment"]})
    assert job.sections == ["cover", "payment"]
    for sections in ["payment", [True]]:
        with pytest.raises(ValueError, match="list of section keys"):
            Job(mode="render", parent_id="parent", config={"report_sections": sections})
    with pytest.raises(ValueError, match="Import the analysis"):
        Job(mode="preview", input=Source(kind="result", asset_id="asset"), draft={})
