"""Transactional server catalogue and containment boundaries for managed evidence."""

from __future__ import annotations

import builtins
import hashlib
import json
import secrets
import time
from pathlib import Path, PurePosixPath
from uuid import uuid4

from sqlalchemy import (
    JSON,
    Column,
    Integer,
    MetaData,
    String,
    Table,
    create_engine,
    event,
    select,
    update,
)


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
    """Small shared-workspace catalogue; evidence/results remain ordinary portable files."""

    def __init__(self, directory: str | Path):
        """Open a local SQLite catalogue with bounded writer contention."""
        self.root = Path(directory).resolve()
        self.root.mkdir(parents=True, exist_ok=True)
        self.engine = create_engine(
            "sqlite:///" + str(self.root / "catalogue.sqlite"), connect_args={"timeout": 20}
        )

        @event.listens_for(self.engine, "connect")
        def configure(connection, record):
            connection.execute("PRAGMA journal_mode=WAL")
            connection.execute("PRAGMA foreign_keys=ON")

        metadata = MetaData()
        self.entries = Table(
            "entries",
            metadata,
            Column("id", String, primary_key=True),
            Column("kind", String, nullable=False, index=True),
            Column("revision", Integer, nullable=False),
            Column("data", JSON, nullable=False),
        )
        # Prototype catalogues already created with this identical schema can
        # be adopted once, without dropping their users or accepted runs.
        from alembic import command
        from alembic.config import Config
        from sqlalchemy import inspect

        configuration = Config()
        configuration.set_main_option("script_location", str(Path(__file__).parent / "migrations"))
        with self.engine.begin() as connection:
            configuration.attributes["connection"] = connection
            tables = inspect(connection).get_table_names()
            if "entries" in tables and "alembic_version" not in tables:
                columns = {column["name"] for column in inspect(connection).get_columns("entries")}
                if columns != {"id", "kind", "revision", "data"}:
                    raise ValueError("Unrecognised catalogue schema; restore a compatible backup")
                command.stamp(configuration, "0001_catalogue")
            command.upgrade(configuration, "head")

    def create(self, kind: str, data: dict, identity: str | None = None) -> dict:
        """Create a detached object under an opaque identifier."""
        value = {"id": identity or uuid4().hex, "kind": kind, "revision": 0, "data": data}
        with self.engine.begin() as connection:
            connection.execute(self.entries.insert().values(**value))
        return self.get(str(value["id"]), kind)

    def get(self, identity: str, kind: str | None = None) -> dict:
        """Read one catalogue object without accepting filesystem paths as identifiers."""
        with self.engine.connect() as connection:
            row = (
                connection.execute(select(self.entries).where(self.entries.c.id == identity))
                .mappings()
                .first()
            )
        if row is None or (kind is not None and row["kind"] != kind):
            raise KeyError("Object not found")
        return dict(row)

    def list(self, kind: str) -> list[dict]:
        """Return stable creation order for a small household catalogue."""
        with self.engine.connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    select(self.entries).where(self.entries.c.kind == kind)
                ).mappings()
            ]

    def upload_files(self, asset: str) -> builtins.list[dict]:
        """Fetch one folder manifest only when needed, rather than on every chunk."""
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

    def replace(self, entry: dict, data: dict) -> dict:
        """Commit a conditional replacement; concurrent callers cannot silently overwrite."""
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
        return self.get(entry["id"])

    def patch(self, identity: str, **fields: object) -> dict:
        """Update worker-owned status without sharing a live database connection."""
        attempt = 0
        while True:
            current = self.get(identity)
            try:
                return self.replace(current, {**current["data"], **fields})
            except Conflict:
                attempt += 1
                if attempt == 8:
                    raise

    def directory(self, entry: dict) -> Path:
        """Locate a managed object directory using its server-generated identity."""
        path = contained(self.root, entry["kind"] + "/" + entry["id"])
        path.mkdir(parents=True, exist_ok=True)
        return path

    def user(self, name: str, password: str) -> None:
        """Create a named account with a salted password hash, never a default password."""
        if len(password) < 12 or not name.strip():
            raise ValueError("Choose a username and a password of at least 12 characters")
        salt = secrets.token_hex(16)
        digest = hashlib.pbkdf2_hmac(
            "sha256", password.encode(), bytes.fromhex(salt), 600_000
        ).hex()
        self.create(
            "user",
            {"name": name, "salt": salt, "digest": digest},
            "user-" + hashlib.sha256(name.encode()).hexdigest(),
        )

    def login(self, name: str, password: str) -> tuple[str, str]:
        """Issue a short-lived opaque session after constant-time password comparison."""
        try:
            user = self.get("user-" + hashlib.sha256(name.encode()).hexdigest(), "user")["data"]
        except KeyError:
            user = {"salt": "00" * 16, "digest": "00" * 32}
        digest = hashlib.pbkdf2_hmac(
            "sha256", password.encode(), bytes.fromhex(user["salt"]), 600_000
        ).hex()
        if not secrets.compare_digest(user["digest"], digest):
            raise ValueError("Invalid username or password")
        token, csrf = secrets.token_urlsafe(32), secrets.token_urlsafe(32)
        self.create(
            "session",
            {"name": name, "csrf": csrf, "expires": time.time() + 43200},
            hashlib.sha256(token.encode()).hexdigest(),
        )
        return token, csrf

    def session(self, token: str) -> dict:
        """Validate session expiry; raw bearer tokens are never stored."""
        entry = self.get(hashlib.sha256(token.encode()).hexdigest(), "session")
        if entry["data"]["expires"] <= time.time():
            raise KeyError("Session expired")
        return entry

    def log(self, identity: str, message: str) -> None:
        """Append bounded persistent progress without changing the job's edit revision."""
        self.create("event", {"job": identity, "message": str(message)[-4000:], "at": time.time()})

    def events(self, identity: str, after: int = 0) -> builtins.list[dict]:
        """Replay durable events after a client's last sequence number."""
        # Rowids provide insertion ordering even when event timestamps are equal.
        from sqlalchemy import text

        with self.engine.connect() as connection:
            rows = connection.execute(
                text(
                    "SELECT rowid, data FROM entries WHERE kind='event' AND json_extract(data, '$.job') = :job AND rowid > :after ORDER BY rowid LIMIT 1000"
                ),
                {"after": after, "job": identity},
            )
            return [{"sequence": row[0], **json.loads(row[1])} for row in rows]
