"""Independent access, revocation and migration contracts on SQLite and real PostgreSQL."""

import json
import os
import sqlite3
from concurrent.futures import ThreadPoolExecutor
from threading import Barrier
from uuid import uuid4

import pytest
from fastapi.testclient import TestClient
from sqlalchemy import create_engine, text
from sqlalchemy.engine import make_url

from edf_bill_fetcher.web.accounts import Conflict, password_hash, username
from edf_bill_fetcher.web.api import create_app
from edf_bill_fetcher.web.storage import Store
from edf_bill_fetcher.web.worker import execute
from tests.test_web import accepted


@pytest.fixture(params=["sqlite", "postgres"])
def database(request, tmp_path):
    if request.param == "sqlite":
        yield "sqlite:///" + str(tmp_path / "server" / "catalogue.sqlite")
        return
    url = os.environ.get("EDF_TEST_DATABASE_URL")
    if not url:
        pytest.skip("Set EDF_TEST_DATABASE_URL to run real PostgreSQL contracts")
    engine = create_engine(url)
    name = "edf_test_" + uuid4().hex
    with engine.begin() as connection:
        connection.execute(text("CREATE SCHEMA " + name))
    try:
        yield (
            make_url(url)
            .update_query_dict({"options": "-csearch_path=" + name})
            .render_as_string(hide_password=False)
        )
    finally:
        with engine.begin() as connection:
            connection.execute(text("DROP SCHEMA " + name + " CASCADE"))
        engine.dispose()


@pytest.fixture
def server(database, tmp_path):
    originals = tmp_path / "originals"
    originals.mkdir()
    (originals / "account.htm").write_text("<html>synthetic source</html>")
    app = create_app(
        tmp_path / "server",
        database_url=database,
        roots={"originals": str(originals)},
        start_worker=False,
    )
    with TestClient(app) as client:
        yield client, app.state.store, app
    app.state.store.engine.dispose()


def login(client, name="admin", password="admin"):
    response = client.post("/api/login", json={"username": name, "password": password})
    assert response.status_code == 200, response.text
    client.headers["x-csrf-token"] = response.json()["csrf"]
    return response.json()


def finish_setup(client, name="operator", password="p"):
    response = client.post("/api/account/setup", json={"username": name, "password": password})
    assert response.status_code == 200, response.text
    client.headers["x-csrf-token"] = response.json()["csrf"]
    return response.json()


def owner(client):
    login(client)
    return finish_setup(client)


def workspace(client):
    response = client.get("/api/workspaces")
    assert response.status_code == 200, response.text
    selected = client.headers.get("x-workspace-id")
    return next((value for value in response.json() if value["id"] == selected), response.json()[0])


def share(client, name, role, identity=None):
    current = workspace(client)
    identity = identity or current["id"]
    return client.put(
        f"/api/workspaces/{identity}/members",
        json={"username": name, "role": role, "revision": current["revision"]},
    )


def new_user(client, name="reader", admin=False):
    response = client.post(
        "/api/users", json={"username": name, "password": "b", "is_admin": admin}
    )
    assert response.status_code == 201, response.text
    assert not {"digest", "salt"} & response.json().keys()
    return response.json()


def test_first_login_setup_short_password_normalization_and_default_revocation(server):
    client, store, _ = server
    before = login(client)
    assert before["setup_required"] is True
    for path in [
        "/api/jobs",
        "/api/assets",
        "/api/schema",
        "/api/workspaces",
        "/api/users",
        "/api/diagnostics",
    ]:
        assert client.get(path).status_code == 428
    assert (
        client.post(
            "/api/account/setup", json={"username": "admin", "password": "admin"}
        ).status_code
        == 400
    )
    account = finish_setup(client, "  Operator  ", "p")
    assert account["id"] == before["id"]
    assert account["username"] == "operator"
    assert account["setup_required"] is False
    assert client.get("/api/jobs").status_code == 200
    assert workspace(client)["role"] == "owner"
    assert (
        client.post("/api/login", json={"username": "admin", "password": "admin"}).status_code
        == 401
    )
    assert login(client, " OpErAtOr ", "p")["id"] == account["id"]
    assert store.accounts.ensure_bootstrap() is False
    assert [user["username"] for user in store.accounts.users()] == ["operator"]


def test_existing_account_is_not_replaced_by_bootstrap(database, tmp_path):
    store = Store(tmp_path / "server", database)
    original = store.user("Existing", "a")
    app = create_app(tmp_path / "server", database_url=database, start_worker=False)
    with TestClient(app) as client:
        assert (
            client.post("/api/login", json={"username": "admin", "password": "admin"}).status_code
            == 401
        )
        assert login(client, "existing", "a")["id"] == original["id"]
    store.engine.dispose()
    app.state.store.engine.dispose()


@pytest.mark.parametrize(
    "path", ["/api/jobs", "/api/%6Aobs", "/%61pi/assets", "/api/workspaces", "/api/users"]
)
def test_encoded_routes_do_not_bypass_authentication(server, path):
    client, _, _ = server
    assert client.get(path).status_code == 401


def test_workspaces_isolate_every_evidence_endpoint_and_submitted_reference(server):
    client, store, app = server
    owner(client)
    other = new_user(client)
    baseline, _ = accepted(store)
    asset = client.post("/api/assets", json={"kind": "auto", "name": "Private upload"}).json()
    with TestClient(app) as reader:
        login(reader, "reader", "b")
        finish_setup(reader, "reader", "b")
        assert workspace(reader)["role"] == "owner"
        assert workspace(reader)["id"] != workspace(client)["id"]
        assert reader.get("/api/jobs").json() == []
        assert reader.get("/api/assets").json() == []
        for suffix in [
            "",
            "/review",
            "/artifacts",
            "/download?path=result.json",
            "/pack",
            "/events",
            "/sources",
            "/source?source_id=any",
            "/intake",
            "/intake-source?item_id=any",
        ]:
            assert reader.get(f"/api/jobs/{baseline['id']}" + suffix).status_code == 404
        for suffix in ["offset?path=x", "content?path=x"]:
            method = reader.get if suffix.startswith("offset") else reader.put
            assert method(f"/api/assets/{asset['id']}/" + suffix).status_code == 404
        assert reader.post(f"/api/jobs/{baseline['id']}/cancel", json={}).status_code == 404
        for payload in [
            {"mode": "render", "parent_id": baseline["id"]},
            {"mode": "diff", "parent_id": baseline["id"], "other_id": baseline["id"]},
            {"mode": "extract", "sources": [{"kind": "auto", "asset_id": asset["id"]}]},
            {
                "mode": "extract",
                "sources": [{"kind": "auto", "intake_id": baseline["id"], "relative": "__all__"}],
            },
        ]:
            assert reader.post("/api/jobs", json=payload).status_code == 404
        assert (
            reader.get(
                "/api/workspaces", headers={"x-workspace-id": workspace(client)["id"]}
            ).json()[0]["created_by"]
            == other["id"]
        )
        assert (
            reader.get("/api/jobs", headers={"x-workspace-id": workspace(client)["id"]}).status_code
            == 404
        )
        assert reader.get(f"/api/workspaces/{workspace(client)['id']}/members").status_code == 404
        assert reader.get("/api/users").status_code == 403
    private = store.directory(baseline) / "private.pdf"
    private.write_bytes(b"private evidence")
    assert not store.path_allowed(store.accounts.workspaces(other["id"])[0]["id"], str(private), {})


def test_viewer_permissions_are_enforced_and_sharing_changes_affect_live_sessions(server):
    client, store, app = server
    owner(client)
    new_user(client)
    entry, _ = accepted(store)
    (store.directory(entry) / "report.html").write_text("<html>Reviewed evidence</html>")
    identity = workspace(client)["id"]
    assert share(client, "reader", "viewer").status_code == 200
    with TestClient(app) as reader:
        login(reader, "reader", "b")
        finish_setup(reader, "reader", "b")
        reader.headers["x-workspace-id"] = identity
        assert reader.get("/api/jobs").json()[0]["id"] == entry["id"]
        assert reader.get(f"/api/jobs/{entry['id']}/review").status_code == 200
        assert reader.get(f"/api/jobs/{entry['id']}/download?path=report.html").status_code == 200
        assert reader.get(f"/api/jobs/{entry['id']}/pack").status_code == 200
        for path, body in [
            ("/api/assets", {"kind": "auto", "name": "Denied"}),
            ("/api/jobs", {}),
            (f"/api/jobs/{entry['id']}/review", {}),
            (f"/api/jobs/{entry['id']}/cancel", {}),
        ]:
            assert reader.post(path, json=body).status_code == 403
        assert share(client, "reader", "editor").status_code == 200
        own_asset = reader.post("/api/assets", json={"kind": "auto", "name": "Shared upload"})
        assert own_asset.status_code == 201
        assert own_asset.json()["workspace_id"] == identity
        assert share(client, "reader", "viewer").status_code == 200
        assert (
            reader.put(
                f"/api/assets/{own_asset.json()['id']}/content?path=bill.pdf", content=b"changed"
            ).status_code
            == 403
        )
        assert share(client, "reader", None).status_code == 200
        assert reader.get(f"/api/jobs/{entry['id']}/review").status_code == 404


def test_shared_draft_conflict_keeps_original_and_records_actor(server):
    client, store, app = server
    operator = owner(client)
    new_user(client)
    entry, original = accepted(store)
    identity = workspace(client)["id"]
    assert share(client, "reader", "editor").status_code == 200
    with TestClient(app) as reader:
        login(reader, "reader", "b")
        account = finish_setup(reader, "reader", "b")
        reader.headers["x-workspace-id"] = identity
        path = f"/api/jobs/{entry['id']}/review"
        first = client.get(path).json()
        second = reader.get(path).json()
        row = first["rows"][0]["id"]

        def change(revision, amount):
            return {
                "revision": revision,
                "operation": "update",
                "arguments": {
                    "row_id": row,
                    "changes": {"Period Charge (£)": amount},
                    "reason": "Verified synthetic correction",
                    "include": True,
                    "status": "verified",
                },
            }

        assert reader.post(path, json=change(second["revision"], "80")).status_code == 200
        assert client.post(path, json=change(first["revision"], "60")).status_code == 409
        assert client.get(path).json()["rows"][0]["record"]["Period Charge (£)"] == 80
        assert original.summary.charges == 250
        audit = client.get(f"/api/workspaces/{identity}/audit").json()
        assert any(
            event["actor_id"] == account["id"] and event["action"] == "review.updated"
            for event in audit
        )
        assert not any(word in json.dumps(audit) for word in ["salt", "digest", "csrf"])
        assert reader.get(f"/api/workspaces/{identity}/audit").status_code == 403
    assert operator["id"] != account["id"]


def test_stale_sharing_last_owner_and_administrator_has_no_implicit_evidence_access(server):
    client, store, app = server
    operator = owner(client)
    new_user(client, "second", True)
    current = workspace(client)
    assert share(client, "operator", "viewer").status_code == 409
    assert share(client, "operator", None).status_code == 409
    assert share(client, "second", "owner").status_code == 200
    assert (
        client.put(
            f"/api/workspaces/{current['id']}/members",
            json={"username": "second", "role": None, "revision": current["revision"]},
        ).status_code
        == 409
    )
    assert share(client, "second", None).status_code == 200
    baseline, _ = accepted(store)
    with TestClient(app) as second:
        login(second, "second", "b")
        finish_setup(second, "second", "b")
        assert second.get("/api/users").status_code == 200
        assert second.get(f"/api/jobs/{baseline['id']}").status_code == 404
        assert (
            second.put(
                f"/api/workspaces/{current['id']}/members",
                json={
                    "username": "second",
                    "role": "owner",
                    "revision": workspace(client)["revision"],
                },
            ).status_code
            == 404
        )
    assert store.accounts.role(operator["id"], current["id"]) == "owner"


def test_server_folders_require_explicit_grants_and_are_rechecked_for_queued_jobs(server):
    client, store, app = server
    owner(client)
    new_user(client)
    identity = workspace(client)["id"]
    assert client.get("/api/roots").json() == [{"id": "originals", "name": "originals"}]
    with TestClient(app) as reader:
        login(reader, "reader", "b")
        finish_setup(reader, "reader", "b")
        private = workspace(reader)
        assert reader.get("/api/roots").json() == []
        assert reader.get("/api/roots/originals/browse").status_code == 404
        assert (
            reader.put(
                f"/api/workspaces/{private['id']}/roots",
                json={"labels": ["originals"], "revision": private["revision"]},
            ).status_code
            == 403
        )
    queued = client.post(
        "/api/jobs",
        json={
            "mode": "extract",
            "sources": [{"kind": "auto", "root_id": "originals", "relative": "."}],
        },
    ).json()
    current = workspace(client)
    assert (
        client.put(
            f"/api/workspaces/{identity}/roots",
            json={"labels": [], "revision": current["revision"]},
        ).status_code
        == 200
    )
    execute(store, queued["id"], {"originals": str(store.root.parent / "originals")})
    assert store.get(queued["id"])["data"]["status"] == "failed"
    assert "Unknown source folder" in store.get(queued["id"])["data"]["error"]
    assert client.get("/api/roots").json() == []
    assert not store.accounts.ensure_bootstrap(["originals"])
    assert client.get("/api/roots").json() == []


def test_password_changes_reset_disable_and_last_administrator(server):
    client, store, app = server
    operator = owner(client)
    user = new_user(client)
    assert (
        client.put(
            f"/api/users/{operator['id']}",
            json={"revision": operator["revision"], "active": True, "is_admin": False},
        ).status_code
        == 409
    )
    assert (
        client.put(
            f"/api/users/{operator['id']}",
            json={"revision": operator["revision"], "active": False, "is_admin": True},
        ).status_code
        == 400
    )
    with TestClient(app) as reader:
        login(reader, "reader", "b")
        user = finish_setup(reader, "reader", "b")
        for path, payload in [
            ("/api/users", {"username": "unauthorized", "password": "x"}),
            (f"/api/users/{operator['id']}/password", {"password": "x"}),
        ]:
            assert reader.post(path, json=payload).status_code == 403
        assert (
            reader.post(
                "/api/account/setup",
                json={"username": "reader", "password": "c", "current_password": "wrong"},
            ).status_code
            == 403
        )
        assert (
            reader.post(
                "/api/account/setup",
                json={"username": "operator", "password": "c", "current_password": "b"},
            ).status_code
            == 409
        )
        assert (
            client.post(f"/api/users/{user['id']}/password", json={"password": "new"}).status_code
            == 200
        )
        assert reader.get("/api/session").status_code == 401
        assert login(reader, "reader", "new")["setup_required"] is True
        assert reader.get("/api/jobs").status_code == 428
        user = finish_setup(reader, "renamed", "c")
        assert user["id"] == store.accounts.workspaces(user["id"])[0]["created_by"]
        permissions_payload = {"revision": user["revision"], "active": False, "is_admin": False}
        assert client.put(f"/api/users/{user['id']}", json=permissions_payload).status_code == 200
        assert reader.get("/api/session").status_code == 401
        assert (
            reader.post("/api/login", json={"username": "renamed", "password": "c"}).status_code
            == 401
        )
        assert client.put(f"/api/users/{user['id']}", json=permissions_payload).status_code == 409
        permissions_payload["revision"] += 1
        permissions_payload["active"] = True
        assert client.put(f"/api/users/{user['id']}", json=permissions_payload).status_code == 200
        login(reader, "renamed", "c")
        with TestClient(app) as additional:
            login(additional, "renamed", "c")
            response = reader.post(
                "/api/account/setup",
                json={"username": "renamed", "password": "d", "current_password": "c"},
            )
            assert response.status_code == 200
            reader.headers["x-csrf-token"] = response.json()["csrf"]
            assert additional.get("/api/session").status_code == 401
            assert reader.get("/api/session").status_code == 200


def test_simultaneous_owner_demotions_preserve_an_owner(server):
    client, store, _ = server
    first = owner(client)
    second = new_user(client, "second")
    identity = workspace(client)["id"]
    assert share(client, "second", "owner").status_code == 200
    barrier = Barrier(2)

    def demote(user):
        barrier.wait()
        try:
            store.accounts.member(user["id"], identity, user["username"], "viewer")
            return "saved"
        except Conflict:
            return "conflict"

    with ThreadPoolExecutor(2) as executor:
        assert sorted(executor.map(demote, [first, second])) == ["conflict", "saved"]
    assert sorted(store.accounts.role(user["id"], identity) for user in [first, second]) == [
        "owner",
        "viewer",
    ]


@pytest.mark.parametrize("name", ["", "   ", "a" * 101, "bad\x00name"])
def test_invalid_names_and_passwords_are_bounded(name):
    with pytest.raises(ValueError):
        username(name)
    for value in ["", "a" * 1025]:
        with pytest.raises(ValueError):
            password_hash(value)


def test_legacy_catalogue_preserves_hashes_files_roles_and_event_order(tmp_path):
    root = tmp_path / "server"
    root.mkdir()
    salt, digest = password_hash("old")
    connection = sqlite3.connect(root / "catalogue.sqlite")
    connection.execute(
        "CREATE TABLE entries (id TEXT PRIMARY KEY, kind TEXT NOT NULL, revision INTEGER NOT NULL, data JSON NOT NULL)"
    )
    for identity, kind, data in [
        ("existing", "user", {"name": "Existing", "salt": salt, "digest": digest}),
        ("old-job", "job", {"status": "complete"}),
        ("event-1", "event", {"job": "old-job", "message": "first", "at": 1}),
        ("event-2", "event", {"job": "old-job", "message": "second", "at": 1}),
    ]:
        connection.execute(
            "INSERT INTO entries VALUES (?,?,?,?)", (identity, kind, 3, json.dumps(data))
        )
    connection.commit()
    connection.close()
    directory = root / "job" / "old-job"
    directory.mkdir(parents=True)
    (directory / "original.pdf").write_bytes(b"original")
    store = Store(root)
    token, _ = store.login(" existing ", "old")
    assert store.session(token)["user"]["is_admin"] is True
    assert store.accounts.workspaces("existing")[0]["role"] == "owner"
    assert store.get("old-job")["revision"] == 3
    assert store.directory(store.get("old-job")) == directory
    assert (directory / "original.pdf").read_bytes() == b"original"
    assert [event["message"] for event in store.events("old-job")] == ["first", "second"]
    store.engine.dispose()


def test_postgres_import_retains_existing_workspace_and_does_not_overwrite(database, tmp_path):
    if not database.startswith("postgresql"):
        pytest.skip("SQLite-to-PostgreSQL migration requires PostgreSQL")
    root = tmp_path / "server"
    local = Store(root, "sqlite:///" + str(root / "catalogue.sqlite"))
    account = local.user("existing", "e")
    job = local.create("job", {"status": "complete"})
    (local.directory(job) / "original.pdf").write_bytes(b"original bytes")
    local.log(job["id"], "legacy progress")
    local.engine.dispose()
    app = create_app(root, database_url=database, start_worker=False)
    with TestClient(app) as client:
        assert login(client, "existing", "e")["id"] == account["id"]
        assert client.get("/api/jobs").json()[0]["id"] == job["id"]
        assert app.state.store.events(job["id"])[0]["message"] == "legacy progress"
        assert (
            app.state.store.directory(app.state.store.get(job["id"]))
            .joinpath("original.pdf")
            .read_bytes()
            == b"original bytes"
        )
        assert app.state.store.import_sqlite() is False
    app.state.store.engine.dispose()


def test_workspace_creation_rename_stale_requests_and_payload_ownership(server):
    client, _, _ = server
    owner(client)
    created = client.post("/api/workspaces", json={"name": "Complaint review"})
    assert created.status_code == 201
    entry = created.json()
    assert entry["role"] == "owner"
    path = f"/api/workspaces/{entry['id']}"
    settings = client.get(path).json()
    assert settings["revision"] == 0
    assert settings["members"][0]["username"] == "operator"
    assert settings["roots"] == []
    assert client.put(path, json={"name": "Reviewed evidence", "revision": 0}).status_code == 200
    assert client.put(path, json={"name": "Stale replacement", "revision": 0}).status_code == 409
    assert client.put(path, json={"name": " ", "revision": 1}).status_code == 400
    client.headers["x-workspace-id"] = entry["id"]
    assert client.get("/api/jobs").json() == []
    assert client.get("/api/roots").json() == []
    assert (
        client.post(
            "/api/assets", json={"name": "Input", "kind": "auto", "workspace_id": "forged"}
        ).status_code
        == 422
    )
    assert workspace(client)["name"] == "Reviewed evidence"
    assert share(client, "missing-account", "viewer").status_code == 400
    assert (
        client.put(path + "/roots", json={"labels": ["unconfigured"], "revision": 1}).status_code
        == 400
    )


@pytest.mark.parametrize("revoke", ["membership", "account"])
def test_queued_jobs_recheck_their_creator_before_reading_inputs(server, revoke):
    client, store, app = server
    owner(client)
    account = new_user(client)
    current = workspace(client)
    assert share(client, "reader", "editor").status_code == 200
    parent, _ = accepted(store)
    with TestClient(app) as reader:
        login(reader, "reader", "b")
        account = finish_setup(reader, "reader", "b")
        reader.headers["x-workspace-id"] = current["id"]
        queued = reader.post("/api/jobs", json={"mode": "render", "parent_id": parent["id"]})
        assert queued.status_code == 202
        if revoke == "membership":
            assert share(client, "reader", None).status_code == 200
        else:
            assert (
                client.put(
                    f"/api/users/{account['id']}",
                    json={"revision": account["revision"], "active": False, "is_admin": False},
                ).status_code
                == 200
            )
        execute(store, queued.json()["id"], {})
        failed = store.get(queued.json()["id"])
        assert failed["data"]["status"] == "failed"
        assert not (store.directory(failed) / "result.json").exists()


def test_simultaneous_administrator_demotions_preserve_an_administrator(server):
    client, store, _ = server
    first = owner(client)
    second = new_user(client, "second", True)
    barrier = Barrier(2)

    def demote(account):
        barrier.wait()
        try:
            store.accounts.update_user(
                account["id"],
                account["id"],
                revision=account["revision"],
                active=True,
                is_admin=False,
            )
            return "saved"
        except Conflict:
            return "conflict"

    with ThreadPoolExecutor(2) as executor:
        assert sorted(executor.map(demote, [first, second])) == ["conflict", "saved"]
    assert sum(user["active"] and user["is_admin"] for user in store.accounts.users()) == 1


def test_successful_signins_do_not_lock_out_other_users(server):
    client, _, _ = server
    owner(client)
    for _ in range(12):
        login(client, "operator", "p")
    for _ in range(10):
        assert (
            client.post("/api/login", json={"username": "unknown", "password": "wrong"}).status_code
            == 401
        )
    assert (
        client.post("/api/login", json={"username": "unknown", "password": "wrong"}).status_code
        == 429
    )
    assert login(client, "operator", "p")["username"] == "operator"


def test_postgres_prevents_dispatchers_with_different_data_directories(server, database, tmp_path):
    if not database.startswith("postgresql"):
        pytest.skip("Database-wide dispatcher leases require PostgreSQL")
    _, _, _ = server
    first = create_app(tmp_path / "server", database_url=database)
    second = create_app(tmp_path / "second-server", database_url=database)
    with TestClient(first):
        with pytest.raises(Conflict, match="Another app server"):
            with TestClient(second):
                pass
    third = create_app(tmp_path / "second-server", database_url=database)
    with TestClient(third) as client:
        assert client.get("/api/session").status_code == 401


def test_invalid_legacy_username_upgrade_rolls_back_before_retry(tmp_path):
    root = tmp_path / "server"
    root.mkdir()
    salt, digest = password_hash("old")
    connection = sqlite3.connect(root / "catalogue.sqlite")
    connection.execute(
        "CREATE TABLE entries (id TEXT PRIMARY KEY, kind TEXT NOT NULL, revision INTEGER NOT NULL, data JSON NOT NULL)"
    )
    for identity, name in [("first", "Existing"), ("second", " existing ")]:
        connection.execute(
            "INSERT INTO entries VALUES (?,?,?,?)",
            (identity, "user", 0, json.dumps({"name": name, "salt": salt, "digest": digest})),
        )
    connection.commit()
    with pytest.raises(ValueError, match="collide after normalization"):
        Store(root)
    assert connection.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall() == [
        ("entries",)
    ]
    data = {"name": "other", "salt": salt, "digest": digest}
    connection.execute("UPDATE entries SET data=? WHERE id='second'", (json.dumps(data),))
    connection.commit()
    connection.close()
    store = Store(root)
    assert {user["username"] for user in store.accounts.users()} == {"existing", "other"}
    store.engine.dispose()


@pytest.mark.parametrize("location", ["data", "data/workspaces", "."])
def test_server_folders_cannot_bypass_isolation_by_containing_managed_evidence(tmp_path, location):
    with pytest.raises(ValueError, match="separate from the app data"):
        create_app(tmp_path / "data", roots={"bills": str(tmp_path / location)}, start_worker=False)
