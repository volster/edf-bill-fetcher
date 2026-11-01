"""Authenticated household web transport over the existing evidence application."""

from __future__ import annotations

import asyncio
import hashlib
import json
import secrets
import shutil
import subprocess
import sys
import threading
import time
from collections import defaultdict, deque
from contextlib import asynccontextmanager
from dataclasses import asdict
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any
from urllib.parse import urlparse
from zipfile import ZIP_DEFLATED, ZipFile

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from filelock import FileLock, Timeout
from sqlalchemy.exc import IntegrityError

from edf_bill_fetcher.application.evidence_review import (
    build_evidence_review,
    resolve_review_source,
)
from edf_bill_fetcher.application.record_review import REVIEW_FIELDS
from edf_bill_fetcher.application.review_commands import COMMANDS
from edf_bill_fetcher.application.review_journal import ReviewJournal
from edf_bill_fetcher.application.review_session import ReviewSession
from edf_bill_fetcher.domain.json_values import object_value
from edf_bill_fetcher.domain.money import money_text
from edf_bill_fetcher.helpers.capabilities import diagnostic_report
from edf_bill_fetcher.io.adapters.intake import is_pdf_candidate
from edf_bill_fetcher.io.persistence.result_store import load_result
from edf_bill_fetcher.io.persistence.run_manifest import artifact_record
from edf_bill_fetcher.models.config import ConfigDict
from edf_bill_fetcher.models.defaults import EXTRACTION_DEFAULTS
from edf_bill_fetcher.models.report_sections import REPORT_SECTIONS
from edf_bill_fetcher.models.snapshot_adapter import to_snapshot_json
from edf_bill_fetcher.web.contracts import Asset, CompleteFile, Credentials, Edit, Job, Seal
from edf_bill_fetcher.web.storage import Conflict, Store, contained
from edf_bill_fetcher.web.worker import source_path

TERMINAL = {"complete", "review_required", "empty", "failed", "cancelled", "interrupted"}
MAX_CHUNK = 8 * 1024 * 1024


def public_artifact(path: Path, directory: Path) -> bool:
    """Do not publish expanded intake staging or atomic writer temporaries as outputs."""
    relative = path.relative_to(directory)
    return relative.parts[0] not in {"imported", "intake_files"} and not any(
        part.startswith((".edf-", ".complaint-")) for part in relative.parts
    )


def allowed_path(store: Store, roots: dict[str, str], path: str) -> bool:
    """Authorize resolved paths, including symlink targets, before reading originals."""
    resolved = Path(path).resolve()
    return any(
        resolved == root or resolved.is_relative_to(root)
        for root in (store.root, *(Path(value).resolve() for value in roots.values()))
    )


def review_state(store: Store, identity: str) -> tuple[dict, ReviewSession]:
    """Restore the authoritative draft; bounded payload history supplies cross-request undo."""
    job = store.get(identity, "job")
    snapshot = load_result(store.directory(job) / "result.json")
    try:
        draft = store.get("draft-" + identity, "draft")
    except KeyError:
        session = ReviewSession(snapshot)
        try:
            draft = store.create(
                "draft",
                {"history": [session.payload()], "position": 0, "actor": ""},
                "draft-" + identity,
            )
        except IntegrityError:
            draft = store.get("draft-" + identity, "draft")
    return draft, ReviewSession.restore(
        snapshot, draft["data"]["history"][draft["data"]["position"]]
    )


def present_review(entry: dict, session: ReviewSession) -> Any:
    """Expose shared accepted models and detached effective rows, retaining unknown values."""
    summary = session.snapshot.summary
    return to_snapshot_json(
        {
            "revision": entry["revision"],
            "draft": session.payload(),
            "rows": [
                {
                    **row,
                    "decision": session.decision(row["id"]),
                    "status": session.statuses.get(row["id"], {}).get("status", "unreviewed"),
                }
                for row in session.items()
            ],
            "ledger": session.ledger_model(),
            "claims": session.claim_model(),
            "revisions": session.revision_model(),
            "models": object_value(session.snapshot.models),
            "summary": {
                key: money_text(getattr(summary, key))
                for key in ("charges", "net_credits", "opening_balance", "closing_balance")
            },
            "can_undo": entry["data"]["position"] > 0,
            "can_redo": entry["data"]["position"] < len(entry["data"]["history"]) - 1,
        }
    )


class Dispatcher:
    """One bounded process at a time; no analytical work runs on the HTTP event loop."""

    def __init__(self, store: Store) -> None:
        """Capture one catalogue and a cooperative server shutdown token."""
        self.store = store
        self.stop = threading.Event()
        self.thread = threading.Thread(target=self.run, daemon=True)

    def run(self) -> None:
        """Launch queued work and detect native child crashes without publishing success."""
        while not self.stop.wait(0.5):
            for job in self.store.list("job"):
                if self.stop.is_set():
                    break
                if job["data"]["status"] != "queued":
                    continue
                self.store.patch(job["id"], status="running")
                log = self.store.directory(job) / "worker.log"
                try:
                    with log.open("wb") as output:
                        process = subprocess.Popen(
                            [
                                sys.executable,
                                "-m",
                                "edf_bill_fetcher.web.worker",
                                "--data",
                                str(self.store.root),
                                "--job",
                                job["id"],
                            ],
                            stdout=output,
                            stderr=subprocess.STDOUT,
                        )
                        while process.poll() is None:
                            stopping = self.stop.wait(0.5)
                            cancelled = self.store.get(job["id"])["data"].get("cancel")
                            if stopping or cancelled:
                                if not cancelled:
                                    self.store.patch(job["id"], cancel=True)
                                try:
                                    process.wait(timeout=10)
                                except subprocess.TimeoutExpired:
                                    process.kill()
                                    process.wait()
                                if (
                                    cancelled
                                    and self.store.get(job["id"])["data"]["status"] == "running"
                                ):
                                    self.store.patch(
                                        job["id"],
                                        status="cancelled",
                                        error="Worker stopped after the cancellation grace period; any saved artifacts remain available",
                                    )
                        if self.store.get(job["id"])["data"]["status"] == "running":
                            self.store.patch(
                                job["id"],
                                status="interrupted",
                                error=f"Worker exited ({process.returncode}); saved artifacts remain available. Retry in a fresh run.",
                            )
                except OSError as error:
                    self.store.patch(job["id"], status="failed", error=str(error))


def create_app(
    directory: str | Path,
    *,
    roots: dict[str, str] | None = None,
    start_worker: bool = True,
    secure_cookie: bool = False,
) -> FastAPI:
    """Create an independently installable source/Docker server with one shared workspace."""
    store = Store(directory)
    roots = {key: str(Path(value).resolve()) for key, value in (roots or {}).items()}
    attempts: dict[str, deque] = defaultdict(deque)
    login_lock = threading.Lock()
    dispatcher = Dispatcher(store)
    server_lock = FileLock(str(store.root / "server.lock"), timeout=0)

    @asynccontextmanager
    async def lifespan(app):
        if start_worker:
            server_lock.acquire()
            try:
                (store.root / "roots.json").write_text(json.dumps(roots), encoding="utf-8")
                for job in store.list("job"):
                    if job["data"]["status"] == "running":
                        store.patch(
                            job["id"],
                            status="interrupted",
                            error="Server stopped during this job. Retry creates a fresh run.",
                        )
                dispatcher.thread.start()
                yield
            finally:
                dispatcher.stop.set()
                await asyncio.to_thread(dispatcher.thread.join, 15)
                server_lock.release()
        else:
            yield

    app = FastAPI(
        title="EDF Evidence Collector",
        version="1",
        lifespan=lifespan,
        docs_url=None,
        redoc_url=None,
        openapi_url=None,
    )
    app.state.store = store

    @app.exception_handler(KeyError)
    async def missing(request, error):
        return JSONResponse({"detail": "Item not found"}, status_code=404)

    @app.exception_handler(Timeout)
    async def busy_upload(request, error):
        return JSONResponse(
            {
                "detail": "This upload is busy verifying or saving another request; wait and resume the same selection"
            },
            status_code=409,
            headers={"Retry-After": "1"},
        )

    @app.exception_handler(FileNotFoundError)
    async def unavailable(request, error):
        return JSONResponse(
            {"detail": "This run has no saved analysis or the requested file is unavailable"},
            status_code=404,
        )

    @app.exception_handler(ValueError)
    async def invalid(request, error):
        return JSONResponse(
            {"detail": str(error)}, status_code=409 if isinstance(error, Conflict) else 400
        )

    @app.middleware("http")
    async def authenticate(request: Request, call_next: Any) -> Any:
        received = 0
        original_receive = request.receive

        async def limited_receive():
            nonlocal received
            message = await original_receive()
            received += len(message.get("body", b""))
            if received > MAX_CHUNK:
                raise HTTPException(413, "Request exceeds the 8 MiB chunk limit")
            return message

        request._receive = limited_receive
        if (
            request.headers.get("content-length", "0").isdigit()
            and int(request.headers.get("content-length", "0")) > MAX_CHUNK
        ):
            return JSONResponse(
                {"detail": "Send uploads in chunks of at most 8 MiB"}, status_code=413
            )
        if request.url.path.startswith("/api/"):
            origin = request.headers.get("origin")
            if (
                request.method not in {"GET", "HEAD"}
                and origin
                and urlparse(origin).netloc != request.headers.get("host")
            ):
                return JSONResponse({"detail": "Cross-origin write denied"}, status_code=403)
            if request.url.path != "/api/login":
                try:
                    session = store.session(request.cookies.get("edf_session", ""))
                except KeyError:
                    return JSONResponse(
                        {"detail": "Sign in to access this workspace"}, status_code=401
                    )
                request.state.session = session
                if request.method not in {"GET", "HEAD"} and not secrets.compare_digest(
                    request.headers.get("x-csrf-token", ""), session["data"]["csrf"]
                ):
                    return JSONResponse(
                        {"detail": "Session verification failed; reload before saving"},
                        status_code=403,
                    )
        response = await call_next(request)
        if received > MAX_CHUNK:
            return JSONResponse(
                {"detail": "Request exceeds the 8 MiB chunk limit"}, status_code=413
            )
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["Referrer-Policy"] = "same-origin"
        response.headers["Cache-Control"] = "no-store"
        return response

    @app.post("/api/login")
    def login(credentials: Credentials, request: Request) -> Any:
        client = request.client.host if request.client else "local"
        with login_lock:
            recent = attempts[client]
            while recent and recent[0] < time.monotonic() - 60:
                recent.popleft()
            if len(recent) >= 10:
                raise HTTPException(429, "Too many sign-in attempts; wait a minute")
            recent.append(time.monotonic())
        try:
            token, csrf = store.login(credentials.username, credentials.password)
        except ValueError as error:
            raise HTTPException(401, str(error)) from error
        response = JSONResponse({"username": credentials.username, "csrf": csrf})
        response.set_cookie(
            "edf_session",
            token,
            httponly=True,
            samesite="strict",
            secure=secure_cookie,
            max_age=43200,
        )
        return response

    @app.get("/api/session")
    def session(request: Request) -> Any:
        return {
            "username": request.state.session["data"]["name"],
            "csrf": request.state.session["data"]["csrf"],
        }

    @app.post("/api/logout")
    def logout(request: Request) -> Any:
        entry = request.state.session
        store.replace(entry, {**entry["data"], "expires": 0})
        response = JSONResponse({"signed_out": True})
        response.delete_cookie("edf_session")
        return response

    @app.get("/api/schema")
    def schema():
        return {
            "review_fields": REVIEW_FIELDS,
            "commands": COMMANDS,
            "settings": {key: str(value) for key, value in ConfigDict.__annotations__.items()},
            "sections": [
                {"key": "cover", "title": "Cover page"},
                {"key": "toc", "title": "Contents"},
                *[asdict(section) for section in REPORT_SECTIONS],
            ],
            "defaults": EXTRACTION_DEFAULTS,
            "job_schema": Job.model_json_schema(),
        }

    @app.get("/api/openapi.json")
    def openapi():
        return app.openapi()

    @app.get("/api/diagnostics")
    def diagnostics():
        return diagnostic_report()

    @app.get("/api/roots")
    def root_list():
        return [{"id": key, "name": key} for key in roots]

    @app.get("/api/roots/{identity}/browse")
    def browse(identity: str, relative: str = "") -> Any:
        if identity not in roots:
            raise KeyError(identity)
        root = Path(roots[identity])
        path = contained(root, relative) if relative not in {"", ".", "./"} else root
        return [
            {
                "name": child.name,
                "relative": child.relative_to(root).as_posix(),
                "directory": child.is_dir(),
            }
            for child in sorted(
                path.iterdir(), key=lambda item: (not item.is_dir(), item.name.casefold())
            )
            if child.resolve().is_relative_to(root)
        ]

    @app.get("/api/assets")
    def assets():
        return [
            {
                **entry,
                "data": {key: value for key, value in entry["data"].items() if key != "files"},
            }
            for entry in store.list("asset")
        ]

    @app.post("/api/assets", status_code=201)
    def create_asset(asset: Asset) -> Any:
        return store.create(
            "asset", {**asset.model_dump(), "sealed": False, "files": [], "file_count": 0}
        )

    def upload_record(asset: dict, path: str) -> dict:
        directory = store.directory(asset)
        canonical = contained(directory, path).relative_to(directory).as_posix()
        identity = "file-" + asset["id"] + "-" + hashlib.sha256(canonical.encode()).hexdigest()
        try:
            return store.get(identity, "upload_file")
        except KeyError:
            count = asset["data"].get("file_count", 0)
            if count >= 100_000:
                raise HTTPException(413, "Folder exceeds the 100,000-file limit") from None
            entry = store.create("upload_file", {"path": canonical, "complete": False}, identity)
            store.replace(asset, {**asset["data"], "file_count": count + 1})
            return entry

    @app.get("/api/assets/{identity}/offset")
    def offset(identity: str, path: str) -> Any:
        with FileLock(str(contained(store.root, identity + ".upload.lock")), timeout=10):
            asset = store.get(identity, "asset")
            target = contained(store.directory(asset), path)
            return {
                "offset": target.stat().st_size if target.is_file() else 0,
                "sha256": artifact_record(str(target))["sha256"] if target.is_file() else None,
            }

    @app.put("/api/assets/{identity}/content")
    async def upload(identity: str, request: Request, path: str, offset: int = 0) -> Any:
        content = bytearray()
        async for chunk in request.stream():
            content.extend(chunk)
            if len(content) > MAX_CHUNK:
                raise HTTPException(413, "Upload chunk exceeds 8 MiB")

        def commit_chunk():
            with FileLock(str(contained(store.root, identity + ".upload.lock")), timeout=10):
                asset = store.get(identity, "asset")
                if asset["data"]["sealed"]:
                    raise Conflict("This source is finalized; create a new upload to change it")
                target = contained(store.directory(asset), path)
                actual = target.stat().st_size if target.exists() else 0
                if offset != actual:
                    raise Conflict(f"Upload offset changed; resume at {actual}")
                if actual + len(content) > 64 * 1024**3:
                    raise HTTPException(413, "File exceeds the 64 GiB upload limit")
                if len(content) > shutil.disk_usage(store.root).free:
                    raise HTTPException(507, "Not enough free space to continue this upload")
                file_record = upload_record(asset, path)
                target.parent.mkdir(parents=True, exist_ok=True)
                with target.open("ab") as output:
                    output.write(content)
                    output.flush()
                store.replace(
                    file_record,
                    {**file_record["data"], "size": actual + len(content), "complete": False},
                )
            return {"offset": actual + len(content)}

        return await asyncio.to_thread(commit_chunk)

    @app.post("/api/assets/{identity}/file-complete")
    def complete_file(identity: str, expected: CompleteFile, path: str) -> Any:
        with FileLock(str(contained(store.root, identity + ".upload.lock")), timeout=10):
            asset = store.get(identity, "asset")
            if asset["data"]["sealed"]:
                raise Conflict("This source is already finalized")
            actual = artifact_record(str(contained(store.directory(asset), path)))
            if actual["bytes"] != expected.size or actual["sha256"] != expected.sha256:
                raise Conflict("Uploaded bytes do not match the selected original")
            entry = upload_record(asset, path)
            return store.replace(
                entry, {**entry["data"], **expected.model_dump(), "complete": True}
            )

    @app.post("/api/assets/{identity}/seal")
    def seal(identity: str, expected: Seal) -> Any:
        with FileLock(str(contained(store.root, identity + ".upload.lock")), timeout=10):
            asset = store.get(identity, "asset")
            entries = store.upload_files(identity)
            files = [entry["data"]["path"] for entry in entries] or asset["data"]["files"]
            if not files or (
                asset["data"]["kind"] != "auto"
                and not asset["data"]["kind"].endswith("_dir")
                and len(files) != 1
            ):
                raise ValueError(
                    "Upload one file for this source, or at least one file for a folder"
                )
            if expected.files:
                if set(expected.files) != set(files):
                    raise ValueError("The uploaded filenames do not match the selected originals")
                for path, digest in expected.files.items():
                    actual = artifact_record(str(contained(store.directory(asset), path)))
                    if actual["bytes"] != digest.size or actual["sha256"] != digest.sha256:
                        raise ValueError(
                            "The uploaded bytes do not match the original; resume or create a fresh source"
                        )
            if expected.batch_id is not None:
                if (
                    len(entries) != expected.files_count
                    or any(
                        not entry["data"].get("complete")
                        or entry["data"].get("batch_id") != expected.batch_id
                        for entry in entries
                    )
                    or {entry["data"].get("order") for entry in entries} != set(range(len(entries)))
                ):
                    raise Conflict(
                        "The selected files do not match this upload batch; select all original files or start a fresh upload"
                    )
                files = [
                    entry["data"]["path"]
                    for entry in sorted(entries, key=lambda entry: entry["data"]["order"])
                ]
            return store.replace(asset, {**asset["data"], "files": files, "sealed": True})

    @app.get("/api/jobs")
    def jobs():
        return store.list("job")

    @app.post("/api/jobs", status_code=202)
    def submit(job: Job, request: Request) -> Any:
        for source in [*job.sources, *([job.input] if job.input else [])]:
            source_path(store, source, roots)
        if job.parent_id:
            store.get(job.parent_id, "job")
        if job.other_id:
            store.get(job.other_id, "job")
        return store.create(
            "job",
            {
                "request": job.model_dump(),
                "status": "queued",
                "cancel": False,
                "actor": request.state.session["data"]["name"],
                "created": time.time(),
                "error": None,
            },
        )

    @app.get("/api/jobs/{identity}")
    def job(identity: str) -> Any:
        return store.get(identity, "job")

    @app.post("/api/jobs/{identity}/cancel")
    def cancel(identity: str) -> Any:
        job = store.get(identity, "job")
        status = job["data"]["status"]
        return store.replace(
            job,
            {
                **job["data"],
                "cancel": True,
                "status": "cancelled" if status == "queued" else status,
            },
        )

    @app.get("/api/jobs/{identity}/events")
    async def events(identity: str, request: Request, after: int = 0) -> Any:
        store.get(identity, "job")
        try:
            cursor = int(request.headers.get("last-event-id", after))
        except ValueError as error:
            raise HTTPException(400, "Invalid event cursor") from error

        async def stream():
            nonlocal cursor
            while not await request.is_disconnected():
                batch = store.events(identity, cursor)
                for value in batch:
                    cursor = value["sequence"]
                    yield f"id: {cursor}\ndata: {json.dumps(value)}\n\n"
                if len(batch) == 1000:
                    continue
                if store.get(identity, "job")["data"]["status"] in TERMINAL:
                    yield "event: finished\ndata: {}\n\n"
                    break
                yield ": heartbeat\n\n"
                await asyncio.sleep(1)

        return StreamingResponse(
            stream(), media_type="text/event-stream", headers={"X-Accel-Buffering": "no"}
        )

    @app.get("/api/jobs/{identity}/artifacts")
    def artifacts(identity: str) -> Any:
        directory = store.directory(store.get(identity, "job"))
        return [
            {"path": path.relative_to(directory).as_posix(), "size": path.stat().st_size}
            for path in sorted(directory.rglob("*"))
            if path.is_file()
            and path.resolve().is_relative_to(directory)
            and public_artifact(path, directory)
        ]

    @app.get("/api/jobs/{identity}/intake")
    def intake(
        identity: str, offset: int = 0, status: str = "", limit: int = 200, search: str = ""
    ) -> Any:
        directory = store.directory(store.get(identity, "job"))
        rows = json.loads((directory / "intake.json").read_text(encoding="utf-8"))
        filtered = [
            row
            for row in rows
            if (not status or row["status"] == status)
            and search.casefold() in row["name"].casefold()
        ]
        return {
            "total": len(rows),
            "matched": len(filtered),
            "counts": {
                key: sum(row["status"] == key for row in rows)
                for key in {row["status"] for row in rows}
            },
            "items": [
                {key: value for key, value in row.items() if key != "path"}
                for row in filtered[max(0, offset) : max(0, offset) + min(1000, max(1, limit))]
            ],
        }

    @app.get("/api/jobs/{identity}/intake-source")
    def intake_source(identity: str, item_id: str) -> Any:
        directory = store.directory(store.get(identity, "job"))
        rows = json.loads((directory / "intake.json").read_text(encoding="utf-8"))
        item = next((row for row in rows if row["id"] == item_id), None)
        if (
            item is None
            or not allowed_path(store, roots, item["path"])
            or not Path(item["path"]).is_file()
        ):
            raise KeyError(item_id)
        target = Path(item["path"])
        if item.get("size") is not None and (
            target.stat().st_size != item["size"] or target.stat().st_mtime_ns != item["mtime_ns"]
        ):
            raise Conflict(
                "This original changed after inspection; inspect again before opening it"
            )
        return FileResponse(target, filename=target.name, content_disposition_type="attachment")

    @app.get("/api/jobs/{identity}/download")
    def download(identity: str, path: str) -> Any:
        directory = store.directory(store.get(identity, "job"))
        target = contained(directory, path)
        if not target.is_file() or not public_artifact(target, directory):
            raise KeyError(path)
        return FileResponse(target, filename=target.name, content_disposition_type="attachment")

    @app.get("/api/jobs/{identity}/pack")
    def pack(identity: str) -> Any:
        job = store.get(identity, "job")
        if job["data"]["status"] not in TERMINAL:
            raise Conflict("Wait for the run to stop before downloading its complete pack")
        directory = store.directory(job)
        temporary = TemporaryDirectory(prefix="edf-download-")
        archive = Path(temporary.name) / "Evidence_pack.zip"
        with ZipFile(archive, "w", ZIP_DEFLATED) as output:
            for path in sorted(directory.rglob("*")):
                if (
                    path.is_file()
                    and path.resolve().is_relative_to(directory)
                    and public_artifact(path, directory)
                ):
                    output.write(path, path.relative_to(directory).as_posix())
        from starlette.background import BackgroundTask

        return FileResponse(
            archive, filename="Evidence_pack.zip", background=BackgroundTask(temporary.cleanup)
        )

    @app.get("/api/jobs/{identity}/review")
    def review(identity: str) -> Any:
        entry, session = review_state(store, identity)
        return present_review(entry, session)

    @app.post("/api/jobs/{identity}/review")
    def edit(identity: str, command: Edit, request: Request) -> Any:
        entry, session = review_state(store, identity)
        if command.revision != entry["revision"]:
            raise Conflict(
                "This draft changed in another tab; reload before applying your pending edit"
            )
        journal = ReviewJournal(
            session.snapshot, entry["data"]["history"], entry["data"]["position"]
        )
        journal.apply(command.operation, command.arguments)
        session = journal.session
        saved = store.replace(
            entry,
            {
                "history": journal.history,
                "position": journal.position,
                "actor": request.state.session["data"]["name"],
            },
        )
        return present_review(saved, session)

    @app.get("/api/jobs/{identity}/sources")
    def sources(identity: str) -> Any:
        job = store.get(identity, "job")
        result = load_result(store.directory(job) / "result.json")
        view = build_evidence_review(
            result,
            str(store.directory(job)),
            allowed_source=lambda path: allowed_path(store, roots, path),
        )
        return [
            {
                "category": entry.category,
                "row_id": entry.row_id,
                "fields": object_value(entry.fields),
                "relation_note": entry.relation_note,
                "sources": [
                    {key: value for key, value in asdict(source).items() if key != "path"}
                    | {
                        "available": source.path is not None,
                        "media_type": "application/pdf"
                        if source.path and is_pdf_candidate(source.path)
                        else "application/octet-stream",
                    }
                    for source in entry.sources
                ],
            }
            for entry in view.entries
        ]

    @app.get("/api/jobs/{identity}/source")
    def source(identity: str, source_id: str) -> Any:
        directory = store.directory(store.get(identity, "job"))
        resolved = resolve_review_source(
            load_result(directory / "result.json"),
            str(directory),
            source_id,
            allowed_source=lambda path: allowed_path(store, roots, path),
        )
        if resolved is None:
            raise KeyError(source_id)
        target = Path(resolved)
        if is_pdf_candidate(str(target)):
            return FileResponse(target, media_type="application/pdf")
        return FileResponse(target, filename=target.name, content_disposition_type="attachment")

    static = Path(__file__).parent / "static"
    if static.is_dir():
        app.mount("/", StaticFiles(directory=static, html=True), name="frontend")
    else:

        @app.get("/", include_in_schema=False)
        def frontend_missing():
            return JSONResponse(
                {
                    "detail": "Build the browser interface with npm ci && npm run build in frontend/, or use Vite during development."
                },
                status_code=503,
            )

    return app
