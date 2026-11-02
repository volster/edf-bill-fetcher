"""Transactional metadata, scoped evidence storage and portable database migrations."""

from __future__ import annotations

import builtins
import os
import time
from pathlib import Path, PurePosixPath
from typing import Any
from uuid import uuid4

from filelock import FileLock
from sqlalchemy import create_engine, event, select, update

from edf_bill_fetcher.web import schema


class Conflict(ValueError):
    """A browser edited an obsolete revision or upload offset."""


def contained(root: Path, relative: str) -> Path:
    """Resolve portable relative names without escaping the configured directory."""
    path = PurePosixPath(relative)
    if (
        not relative
        or "\\" in relative
        or ":" in relative
        or path.is_absolute()
        or ".." in path.parts
    ):
        raise ValueError("Use a relative path within the selected folder")
    result = (root / relative).resolve()
    if result == root.resolve() or not result.is_relative_to(root.resolve()):
        raise ValueError("Path is outside the selected folder")
    return result


class Store:
    """PostgreSQL catalogue with a compatible local SQLite option; originals stay in files."""

    def __init__(self, directory: str | Path, database_url: str | None = None) -> None:
        """Upgrade metadata transactionally before exposing accounts or evidence."""
        self.root = Path(directory).resolve()
        self.root.mkdir(parents=True, exist_ok=True)
        self.database_url = (
            database_url
            or os.environ.get("EDF_DATABASE_URL")
            or "sqlite:///" + str(self.root / "catalogue.sqlite")
        )
        if self.database_url.startswith("postgres://"):
            self.database_url = "postgresql+psycopg://" + self.database_url[len("postgres://") :]
        elif self.database_url.startswith("postgresql://"):
            self.database_url = "postgresql+psycopg://" + self.database_url[len("postgresql://") :]
        sqlite = self.database_url.startswith("sqlite:")
        self.engine = create_engine(
            self.database_url,
            pool_pre_ping=True,
            **({"connect_args": {"timeout": 20}} if sqlite else {}),
        )
        if sqlite:

            @event.listens_for(self.engine, "connect")
            def configure(connection, record):
                connection.execute("PRAGMA journal_mode=WAL")
                connection.execute("PRAGMA foreign_keys=ON")

        self.entries = schema.entries
        from alembic import command
        from alembic.config import Config
        from sqlalchemy import inspect, text

        configuration = Config()
        configuration.set_main_option("script_location", str(Path(__file__).parent / "migrations"))
        with (
            FileLock(str(self.root / "migration.lock"), timeout=60),
            self.engine.begin() as connection,
        ):
            if sqlite:
                connection.exec_driver_sql("BEGIN IMMEDIATE")
            else:
                connection.execute(text("SELECT pg_advisory_xact_lock(713754821)"))
            configuration.attributes["connection"] = connection
            tables = inspect(connection).get_table_names()
            if "entries" in tables and "alembic_version" not in tables:
                columns = {column["name"] for column in inspect(connection).get_columns("entries")}
                if columns != {"id", "kind", "revision", "data"}:
                    raise ValueError("Unrecognised catalogue schema; restore a compatible backup")
                command.stamp(configuration, "0001_catalogue")
            command.upgrade(configuration, "head")
        from edf_bill_fetcher.web.accounts import Accounts

        self.accounts = Accounts(self)

    def create(
        self,
        kind: str,
        data: dict,
        identity: str | None = None,
        *,
        workspace_id: str | None = None,
        actor_id: str | None = None,
    ) -> dict:
        """Persist metadata with workspace ownership independent of browser payloads."""
        if workspace_id is None and kind in {"job", "asset", "draft", "upload_file"}:
            workspace_id = self.accounts.default_workspace()
        value: dict[str, Any] = {
            "id": identity or uuid4().hex,
            "kind": kind,
            "revision": 0,
            "data": data,
            "workspace_id": workspace_id,
            "created": time.time(),
        }
        with self.engine.begin() as connection:
            connection.execute(self.entries.insert().values(**value))
            if actor_id:
                self.accounts.audit(
                    connection, actor_id, kind + ".created", value["id"], workspace_id
                )
        return self.get(str(value["id"]), kind)

    def get(self, identity: str, kind: str | None = None) -> dict:
        """Look up an opaque identifier; HTTP callers additionally authorize its workspace."""
        with self.engine.connect() as connection:
            row = (
                connection.execute(select(self.entries).where(self.entries.c.id == identity))
                .mappings()
                .first()
            )
        if row is None or (kind is not None and row["kind"] != kind):
            raise KeyError("Object not found")
        return dict(row)

    def list(self, kind: str, workspace_id: str | None = None) -> list[dict]:
        """Return creation order; transports supply a workspace to enforce isolation."""
        query = select(self.entries).where(self.entries.c.kind == kind)
        if workspace_id is not None:
            query = query.where(self.entries.c.workspace_id == workspace_id)
        with self.engine.connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    query.order_by(self.entries.c.created, self.entries.c.id)
                ).mappings()
            ]

    def upload_files(self, asset: str) -> builtins.list[dict]:
        """Fetch a folder manifest without loading it for every uploaded chunk."""
        prefix = "file-" + asset + "-"
        with self.engine.connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    select(self.entries).where(
                        self.entries.c.id >= prefix, self.entries.c.id < prefix + "~"
                    )
                ).mappings()
            ]

    def register_upload(self, asset: dict, identity: str, path: str) -> dict:
        """Commit a folder member and its collection count together, including on failure."""
        with self.engine.begin() as connection:
            connection.execute(
                self.entries.insert().values(
                    id=identity,
                    kind="upload_file",
                    revision=0,
                    data={"path": path, "complete": False, "write_pending": True},
                    workspace_id=asset["workspace_id"],
                    created=time.time(),
                )
            )
            result = connection.execute(
                update(self.entries)
                .where(
                    self.entries.c.id == asset["id"],
                    self.entries.c.revision == asset["revision"],
                )
                .values(
                    revision=asset["revision"] + 1,
                    data={**asset["data"], "file_count": asset["data"].get("file_count", 0) + 1},
                )
            )
            if result.rowcount != 1:
                raise Conflict("This source changed while adding a file; resume the same selection")
        return self.get(identity, "upload_file")

    def replace(
        self,
        entry: dict,
        data: dict,
        *,
        actor_id: str | None = None,
        action: str = "review.updated",
    ) -> dict:
        """Reject stale writes rather than losing another user's changes."""
        with self.engine.begin() as connection:
            result = connection.execute(
                update(self.entries)
                .where(
                    self.entries.c.id == entry["id"], self.entries.c.revision == entry["revision"]
                )
                .values(revision=entry["revision"] + 1, data=data)
            )
            if result.rowcount != 1:
                raise Conflict(
                    "This item changed in another tab. Reload before saving; your pending edits have not been applied."
                )
            if actor_id:
                self.accounts.audit(
                    connection,
                    actor_id,
                    action,
                    entry["id"],
                    entry.get("workspace_id"),
                    revision=entry["revision"] + 1,
                )
        return self.get(entry["id"])

    def patch(self, identity: str, **fields: object) -> dict:
        """Retry conditional worker status updates after bounded contention."""
        for attempt in range(8):
            current = self.get(identity)
            try:
                return self.replace(current, {**current["data"], **fields})
            except Conflict:
                if attempt == 7:
                    raise
        raise AssertionError("Unreachable retry state")

    def directory(self, entry: dict) -> Path:
        """Keep new files under their workspace while retaining existing relative links."""
        relative = entry["kind"] + "/" + entry["id"]
        legacy = contained(self.root, relative)
        workspace = entry.get("workspace_id")
        path = (
            legacy
            if legacy.exists() or not workspace
            else contained(self.root, f"workspaces/{workspace}/{relative}")
        )
        path.mkdir(parents=True, exist_ok=True)
        return path

    def path_allowed(self, workspace_id: str | None, path: str, roots: dict[str, str]) -> bool:
        """Never resolve an imported manifest through another workspace's managed files."""
        resolved = Path(path).resolve()
        if workspace_id is None:
            managed = [self.root]
        else:
            managed = [contained(self.root, "workspaces/" + workspace_id)]
            if resolved.is_relative_to(self.root):
                parts = resolved.relative_to(self.root).parts
                if len(parts) >= 2 and parts[0] in {"job", "asset"}:
                    try:
                        legacy = self.get(parts[1], parts[0])
                        if legacy["workspace_id"] == workspace_id:
                            managed.append(contained(self.root, parts[0] + "/" + parts[1]))
                    except KeyError:
                        pass
        return any(
            resolved == root or resolved.is_relative_to(root)
            for root in [*managed, *(Path(value).resolve() for value in roots.values())]
        )

    def user(self, name: str, password: str, **options: Any) -> dict:
        """Create an account without an arbitrary minimum password length."""
        return self.accounts.create_user(name, password, **options)

    def login(self, name: str, password: str) -> tuple[str, str]:
        """Verify credentials and issue a relational session."""
        return self.accounts.login(name, password)

    def session(self, token: str) -> dict:
        """Recheck account activation on every request."""
        return self.accounts.session(token)

    def log(self, identity: str, message: str) -> None:
        """Append progress using a real sequence on either database backend."""
        with self.engine.begin() as connection:
            connection.execute(
                schema.events.insert().values(
                    job_id=identity, message=str(message)[-4000:], at=time.time()
                )
            )

    def events(self, identity: str, after: int = 0) -> builtins.list[dict]:
        """Replay progress without SQLite-specific rowid or JSON functions."""
        with self.engine.connect() as connection:
            rows = connection.execute(
                select(schema.events)
                .where(schema.events.c.job_id == identity, schema.events.c.sequence > after)
                .order_by(schema.events.c.sequence)
                .limit(1000)
            ).mappings()
            return [
                {
                    "sequence": row["sequence"],
                    "job": row["job_id"],
                    "message": row["message"],
                    "at": row["at"],
                }
                for row in rows
            ]

    def import_sqlite(self) -> bool:
        """Copy an existing local catalogue into an empty PostgreSQL database atomically."""
        if (
            self.engine.dialect.name != "postgresql"
            or not (self.root / "catalogue.sqlite").is_file()
        ):
            return False
        with self.engine.connect() as connection:
            if (
                connection.execute(select(schema.users.c.id).limit(1)).first()
                or connection.execute(select(schema.entries.c.id).limit(1)).first()
            ):
                return False
        source = Store(self.root, "sqlite:///" + str(self.root / "catalogue.sqlite"))
        try:
            with (
                FileLock(str(self.root / "database-import.lock"), timeout=60),
                self.engine.begin() as destination,
                source.engine.connect() as original,
            ):
                from sqlalchemy import text

                destination.execute(text("SELECT pg_advisory_xact_lock(713754821)"))
                if (
                    destination.execute(select(schema.users.c.id).limit(1)).first()
                    or destination.execute(select(schema.entries.c.id).limit(1)).first()
                ):
                    return False
                for table in [
                    schema.users,
                    schema.workspaces,
                    schema.memberships,
                    schema.workspace_roots,
                    schema.entries,
                    schema.events,
                    schema.audit,
                ]:
                    rows = [dict(row) for row in original.execute(select(table)).mappings()]
                    for row in rows:
                        if table in (schema.events, schema.audit):
                            row.pop("sequence")
                    if rows:
                        destination.execute(table.insert(), rows)
            return True
        finally:
            source.engine.dispose()
