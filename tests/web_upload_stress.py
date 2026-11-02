"""Opt-in HTTP stress check using disposable synthetic evidence and database schemas."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import random
import socket
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from tempfile import TemporaryDirectory
from uuid import uuid4
from zipfile import ZIP_DEFLATED, ZipFile

import httpx
from sqlalchemy import create_engine, text
from sqlalchemy.engine import make_url


def call(client, path, body=None, *, method=None, **kwargs):
    """Fail at the actual HTTP boundary without swallowing unsuccessful requests."""
    response = client.request(
        method or ("POST" if body is not None else "GET"),
        "/api" + path,
        json=body,
        **kwargs,
    )
    assert response.is_success, (path, response.status_code, response.text[:300])
    return response.json()


def sign_in(base, name, password, *, setup=False):
    """Use normal cookies and CSRF, including mandatory new-account setup."""
    client = httpx.Client(base_url=base, timeout=120)
    profile = call(client, "/login", {"username": name, "password": password})
    client.headers["x-csrf-token"] = profile["csrf"]
    if setup:
        profile = call(client, "/account/setup", {"username": name, "password": password})
        client.headers["x-csrf-token"] = profile["csrf"]
    return client


def upload_file(client, asset, path, blocks, batch, order):
    """Stream bounded chunks and verify the resulting original independently."""
    digest = hashlib.sha256()
    size = 0
    for block in blocks:
        response = call(
            client,
            f"/assets/{asset}/content",
            method="PUT",
            params={"path": path, "offset": size},
            content=block,
        )
        digest.update(block)
        size += len(block)
        assert response["offset"] == size
    call(
        client,
        f"/assets/{asset}/file-complete",
        {
            "size": size,
            "sha256": digest.hexdigest(),
            "batch_id": batch,
            "order": order,
        },
        params={"path": path},
    )
    return {"offset": size, "sha256": digest.hexdigest()}


def exercise(base, directory, files, users, large_mib):
    """Check multi-user uploads, isolation, large-file integrity and real ZIP intake."""
    initial = sign_in(base, "admin", "admin")
    profile = call(initial, "/account/setup", {"username": "stress-owner", "password": "s"})
    initial.headers["x-csrf-token"] = profile["csrf"]
    workspace = call(initial, "/workspaces")[0]
    shared = workspace["id"]
    clients = [initial]
    for index in range(1, users):
        name = f"stress-user-{index}"
        call(initial, "/users", {"username": name, "password": "s"})
        client = sign_in(base, name, "s", setup=True)
        workspace = call(initial, f"/workspaces/{shared}")
        call(
            initial,
            f"/workspaces/{shared}/members",
            {
                "username": name,
                "role": "editor",
                "revision": workspace["revision"],
            },
            method="PUT",
        )
        clients.append(client)
    for client in clients:
        client.headers["x-workspace-id"] = shared
    try:

        def folder(index):
            client = clients[index]
            asset = call(client, "/assets", {"kind": "auto", "name": f"Folder {index}"})["id"]
            batch = uuid4().hex
            expected = {}
            for order, number in enumerate(range(index, files, users)):
                path = f"year-{number % 10}/evidence-{number:06}.txt"
                data = f"Synthetic evidence {number}; uploader {index}\n".encode() * 8
                expected[path] = upload_file(client, asset, path, [data], batch, order)
            result = call(
                client, f"/assets/{asset}/seal", {"batch_id": batch, "files_count": len(expected)}
            )
            assert set(result["data"]["files"]) == set(expected)
            assert result["data"]["file_count"] == len(expected)
            for path in list(expected)[:: max(1, len(expected) // 10)]:
                assert (
                    call(client, f"/assets/{asset}/offset", params={"path": path}) == expected[path]
                )
            return asset

        with ThreadPoolExecutor(users) as executor:
            assets = list(executor.map(folder, range(users)))
        # Membership in the shared workspace must not grant access in a private workspace.
        private = next(w for w in call(clients[1], "/workspaces") if w["id"] != shared)
        clients[1].headers["x-workspace-id"] = private["id"]
        assert clients[1].get(f"/api/assets/{assets[0]}/offset?path=x").status_code == 404
        clients[1].headers["x-workspace-id"] = shared

        asset = call(initial, "/assets", {"kind": "auto", "name": "Large original"})["id"]
        block = random.Random(12).randbytes(1024 * 1024)
        large = upload_file(
            initial, asset, "large.pst", (block for _ in range(large_mib)), "large", 0
        )
        assert call(initial, f"/assets/{asset}/offset?path=large.pst") == large
        call(initial, f"/assets/{asset}/seal", {"batch_id": "large", "files_count": 1})

        archive = directory / "synthetic.zip"
        random_bytes = random.Random(17)
        with ZipFile(archive, "w", ZIP_DEFLATED) as bundle:
            for number in range(files):
                bundle.writestr(f"nested/evidence-{number:06}.bin", random_bytes.randbytes(16_384))
        asset = call(initial, "/assets", {"kind": "auto", "name": "Large archive"})["id"]
        with archive.open("rb") as stream:
            expected = upload_file(
                initial,
                asset,
                "synthetic.zip",
                iter(lambda: stream.read(4 * 1024**2), b""),
                "archive",
                0,
            )
        assert call(initial, f"/assets/{asset}/offset?path=synthetic.zip") == expected
        call(initial, f"/assets/{asset}/seal", {"batch_id": "archive", "files_count": 1})
        job = call(
            initial,
            "/jobs",
            {
                "mode": "intake",
                "name": "Archive stress inspection",
                "sources": [{"kind": "auto", "asset_id": asset}],
                "formats": ["xlsx"],
            },
        )
        deadline = time.monotonic() + 300
        while True:
            state = call(initial, f"/jobs/{job['id']}")["data"]
            assert state["status"] not in {"failed", "cancelled", "interrupted"}, state
            if state["status"] in {"complete", "review_required", "empty"}:
                break
            assert time.monotonic() < deadline, "Archive inspection timed out"
            time.sleep(0.25)
        inventory = call(initial, f"/jobs/{job['id']}/intake", params={"limit": 1})
        assert inventory["counts"]["unsupported"] == files, inventory
        assert inventory["total"] == files + 1, inventory
        return {
            "users": users,
            "folder_files": files,
            "large_file_bytes": large["offset"],
            "zip_bytes": archive.stat().st_size,
            "zip_members": files,
        }
    finally:
        for client in clients:
            client.close()


def main():
    """Start a disposable server; PostgreSQL tests use an isolated schema."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--files", type=int, default=2000)
    parser.add_argument("--users", type=int, default=4)
    parser.add_argument("--large-mib", type=int, default=64)
    parser.add_argument("--database-url", default=os.environ.get("EDF_TEST_DATABASE_URL"))
    options = parser.parse_args()
    if (
        not 2 <= options.users <= 10
        or not options.users <= options.files <= 100_000
        or not 1 <= options.large_mib <= 1024
    ):
        parser.error("Use 2–10 users, at least one file per user (up to 100,000), and 1–1024 MiB")
    engine = None
    database_url = ""
    schema = "edf_stress_" + uuid4().hex
    if options.database_url:
        url = make_url(options.database_url)
        if url.get_backend_name() != "postgresql":
            parser.error("Supply a dedicated PostgreSQL test database or omit the URL for SQLite")
        url = url.set(drivername="postgresql+psycopg")
        engine = create_engine(url)
        with engine.begin() as connection:
            connection.execute(text("CREATE SCHEMA " + schema))
        database_url = url.update_query_dict(
            {"options": "-csearch_path=" + schema}
        ).render_as_string(hide_password=False)
    started = time.monotonic()
    try:
        with TemporaryDirectory(prefix="edf-upload-stress-") as temporary:
            directory = Path(temporary)
            with socket.socket() as listener:
                listener.bind(("127.0.0.1", 0))
                port = listener.getsockname()[1]
            base = f"http://127.0.0.1:{port}"
            with (directory / "server.log").open("wb") as log:
                process = subprocess.Popen(
                    [
                        sys.executable,
                        "-m",
                        "edf_bill_fetcher.web.launch",
                        "--data",
                        str(directory / "server"),
                        "--port",
                        str(port),
                    ],
                    env={**os.environ, "EDF_DATABASE_URL": database_url},
                    stdout=log,
                    stderr=log,
                )
                try:
                    with httpx.Client(base_url=base, timeout=2) as probe:
                        for _ in range(300):
                            assert process.poll() is None, (
                                "Server startup failed; inspect test database configuration"
                            )
                            try:
                                if probe.get("/api/session").status_code == 401:
                                    break
                            except httpx.TransportError:
                                pass
                            time.sleep(0.1)
                        else:
                            raise AssertionError("Server startup timed out")
                    result = exercise(
                        base, directory, options.files, options.users, options.large_mib
                    )
                    result["elapsed_seconds"] = round(time.monotonic() - started, 2)
                    result["database"] = "postgresql" if engine else "sqlite"
                    print(json.dumps(result, indent=2))
                finally:
                    process.terminate()
                    try:
                        process.wait(timeout=20)
                    except subprocess.TimeoutExpired:
                        process.kill()
                        process.wait()
    finally:
        if engine:
            with engine.begin() as connection:
                connection.execute(text("DROP SCHEMA " + schema + " CASCADE"))
            engine.dispose()


if __name__ == "__main__":
    main()
