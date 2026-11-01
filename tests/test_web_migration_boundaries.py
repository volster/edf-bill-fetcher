"""A legacy queue cannot bypass the new workspace authorization boundary."""

import json
import sqlite3

import pytest

from edf_bill_fetcher.web.storage import Store


@pytest.mark.parametrize("status", ["queued", "running", "complete"])
def test_legacy_jobs_require_fresh_authorization_without_losing_metadata(tmp_path, status):
    database = sqlite3.connect(tmp_path / "catalogue.sqlite")
    database.execute(
        "CREATE TABLE entries (id TEXT PRIMARY KEY, kind TEXT NOT NULL, revision INTEGER NOT NULL, data JSON NOT NULL)"
    )
    database.execute(
        "INSERT INTO entries VALUES (?,?,?,?)",
        ("legacy-job", "job", 7, json.dumps({"status": status, "marker": "retained"})),
    )
    database.commit()
    database.close()
    store = Store(tmp_path)
    entry = store.get("legacy-job", "job")
    assert entry["data"]["status"] == ("complete" if status == "complete" else "interrupted")
    assert entry["data"]["marker"] == "retained"
    assert entry["revision"] == 7
    if status != "complete":
        assert "Create a fresh run" in entry["data"]["error"]
    store.engine.dispose()
