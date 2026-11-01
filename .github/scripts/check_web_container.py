"""Exercise the installed container through HTTP, including native OCR and review."""

import argparse
import hashlib
import io
import json
import subprocess
import sys
import time
from http.cookiejar import CookieJar
from pathlib import Path
from tempfile import TemporaryDirectory
from urllib.error import HTTPError, URLError
from urllib.request import HTTPCookieProcessor, Request, build_opener
from zipfile import ZipFile

from edf_bill_fetcher.io.persistence.result_store import load_result
from edf_bill_fetcher.web.storage import Store


def main():
    """Require an actual uploaded scan to survive extraction, correction and publication."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--fixture", type=Path, default=Path(__file__).with_name("scanned_bill.pdf")
    )
    options = parser.parse_args()
    original = options.fixture.read_bytes()
    with TemporaryDirectory(prefix="edf-container-check-") as temporary:
        root = Path(temporary)
        store = Store(root / "data")
        store.user("smoke", "synthetic-container-password")
        store.engine.dispose()
        with (root / "server.log").open("wb") as output:
            server = subprocess.Popen(
                [
                    sys.executable,
                    "-m",
                    "edf_bill_fetcher.web.launch",
                    "--data",
                    str(root / "data"),
                    "--port",
                    "8769",
                ],
                stdout=output,
                stderr=subprocess.STDOUT,
            )
            client = build_opener(HTTPCookieProcessor(CookieJar()))
            csrf = ""

            def request(path, payload=None, *, method=None, raw=None):
                body = (
                    raw
                    if raw is not None
                    else json.dumps(payload).encode()
                    if payload is not None
                    else None
                )
                call = Request(
                    "http://127.0.0.1:8769" + path,
                    data=body,
                    method=method,
                    headers={"Content-Type": "application/json", "X-CSRF-Token": csrf},
                )
                with client.open(call, timeout=120) as response:
                    content = response.read()
                    return (
                        json.loads(content)
                        if response.headers.get_content_type() == "application/json"
                        else content
                    )

            def wait_job(identity):
                deadline = time.monotonic() + 180
                while time.monotonic() < deadline:
                    value = request("/api/jobs/" + identity)
                    if value["data"]["status"] in {"complete", "review_required"}:
                        return value
                    assert value["data"]["status"] not in {
                        "failed",
                        "cancelled",
                        "interrupted",
                        "empty",
                    }, value
                    time.sleep(0.25)
                raise RuntimeError("Container processing did not finish within 180 seconds")

            try:
                deadline = time.monotonic() + 45
                while True:
                    try:
                        assert b"<html" in request("/").lower()
                        break
                    except (URLError, HTTPError):
                        if server.poll() is not None or time.monotonic() > deadline:
                            raise
                        time.sleep(0.1)
                session = request(
                    "/api/login", {"username": "smoke", "password": "synthetic-container-password"}
                )
                csrf = session["csrf"]
                asset = request("/api/assets", {"name": "Synthetic scan", "kind": "auto"})
                base = "/api/assets/" + asset["id"]
                request(base + "/content?path=scanned.pdf", method="PUT", raw=original)
                request(
                    base + "/seal",
                    {
                        "files": {
                            "scanned.pdf": {
                                "size": len(original),
                                "sha256": hashlib.sha256(original).hexdigest(),
                            }
                        }
                    },
                )
                run = request(
                    "/api/jobs",
                    {
                        "mode": "extract",
                        "name": "Container OCR",
                        "sources": [{"kind": "auto", "asset_id": asset["id"]}],
                        "formats": ["xlsx", "html", "pdf", "docx"],
                        "config": {"as_of": "2025-03-01"},
                    },
                )
                wait_job(run["id"])
                review_path = "/api/jobs/" + run["id"] + "/review"
                review = request(review_path)
                row = review["rows"][0]
                assert row["record"]["Invoice #"] == "KI-12345", row
                assert row["record"]["Text Extraction"] == "OCR", row
                assert row["record"]["Period Charge (£)"] == 125, row
                assert review["summary"]["charges"] == "125.00", review
                edited = request(
                    review_path,
                    {
                        "revision": review["revision"],
                        "operation": "update",
                        "arguments": {
                            "row_id": row["id"],
                            "changes": {"Period Charge (£)": "140.00"},
                            "reason": "Synthetic container correction",
                            "include": True,
                            "status": "needs_information",
                        },
                    },
                )
                published = request(
                    "/api/jobs",
                    {
                        "mode": "review",
                        "parent_id": run["id"],
                        "draft": edited["draft"],
                        "formats": ["xlsx", "html", "pdf", "docx"],
                    },
                )
                wait_job(published["id"])
                result = request("/api/jobs/" + published["id"] + "/download?path=result.json")
                result_path = root / "reviewed.json"
                result_path.write_text(json.dumps(result), encoding="utf-8")
                assert load_result(result_path).summary.charges == 140
                packed = request("/api/jobs/" + published["id"] + "/pack")
                with ZipFile(io.BytesIO(packed)) as archive:
                    for name in [
                        "Evidence.xlsx",
                        "Evidence_Report.html",
                        "Evidence_Report.pdf",
                        "Evidence_Report.docx",
                        "result.json",
                        "run.json",
                    ]:
                        assert name in archive.namelist(), name
                    originals = [
                        name
                        for name in archive.namelist()
                        if name.startswith("evidence_files/") and name.endswith(".pdf")
                    ]
                    assert len(originals) == 1, originals
                    assert archive.read(originals[0]) == original
                sys.stdout.write(
                    "Container HTTP/OCR/review/all-format/source-preservation checks passed\n"
                )
            finally:
                server.terminate()
                try:
                    server.wait(timeout=20)
                except subprocess.TimeoutExpired:
                    server.kill()
                    server.wait()
                sys.stdout.write((root / "server.log").read_text(encoding="utf-8"))


if __name__ == "__main__":
    main()
