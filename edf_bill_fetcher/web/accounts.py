"""Accounts, session revocation and workspace access enforced by database transactions."""

from __future__ import annotations

import hashlib
import secrets
import time
import unicodedata
from contextlib import contextmanager, nullcontext
from typing import Any
from uuid import uuid4

from filelock import FileLock
from sqlalchemy import case, delete, func, select, text

from edf_bill_fetcher.web import schema as s
from edf_bill_fetcher.web.storage import Conflict

ROLES = {"viewer": 0, "editor": 1, "owner": 2}


class Forbidden(ValueError):
    """The current account does not have the required permission."""


def username(value: str) -> str:
    """Normalize usernames consistently at account creation and sign-in."""
    result = unicodedata.normalize("NFKC", value).strip().casefold()
    if not result or len(result) > 100 or any(ord(c) < 32 for c in result):
        raise ValueError("Choose a username between 1 and 100 characters")
    return result


def password_hash(password: str, salt: str | None = None) -> tuple[str, str]:
    """Accept user-chosen lengths while bounding password hashing input."""
    if not password or len(password) > 1024:
        raise ValueError("Enter a password between 1 and 1024 characters")
    salt = salt or secrets.token_hex(16)
    return salt, hashlib.pbkdf2_hmac(
        "sha256", password.encode(), bytes.fromhex(salt), 600_000
    ).hex()


def public_user(row: Any) -> dict:
    """Keep password material out of all transport responses and audit details."""
    return {
        key: row[key]
        for key in ("id", "username", "active", "is_admin", "setup_required", "created", "revision")
    }


class Accounts:
    """Relational accounts and memberships; administrator status grants no evidence access."""

    def __init__(self, store):
        """Use the catalogue connection pool for identity and access metadata."""
        self.store = store
        self.engine = store.engine

    @contextmanager
    def transaction(self):
        """Serialize membership invariants, including concurrent last-owner/admin changes."""
        lock = (
            FileLock(str(self.store.root / "accounts.lock"), timeout=20)
            if self.engine.dialect.name == "sqlite"
            else nullcontext()
        )
        with lock, self.engine.begin() as connection:
            if self.engine.dialect.name == "postgresql":
                connection.execute(text("SELECT pg_advisory_xact_lock(713754822)"))
            yield connection

    def audit(
        self,
        connection: Any,
        actor: str,
        action: str,
        target: str,
        workspace_id: str | None = None,
        **details: Any,
    ) -> None:
        """Append a secret-free event within the caller’s transaction."""
        connection.execute(
            s.audit.insert().values(
                actor_id=actor,
                action=action,
                target=target,
                workspace_id=workspace_id,
                details=details,
                at=time.time(),
            )
        )

    def record(
        self, actor: str, action: str, target: str, workspace_id: str | None = None, **details: Any
    ) -> None:
        """Append a standalone access event without exposing credential material."""
        with self.engine.begin() as connection:
            self.audit(connection, actor, action, target, workspace_id, **details)

    def default_workspace(self) -> str | None:
        """Find the initial workspace for legacy and internal callers."""
        with self.engine.connect() as connection:
            result = connection.execute(
                select(s.workspaces.c.id)
                .order_by(s.workspaces.c.created, s.workspaces.c.id)
                .limit(1)
            ).scalar()
            return str(result) if result is not None else None

    def _workspace(self, connection: Any, user_id: str, name: str) -> dict:
        name = name.strip()
        if not name or len(name) > 120:
            raise ValueError("Choose a workspace name between 1 and 120 characters")
        value: dict[str, Any] = {
            "id": uuid4().hex,
            "name": name,
            "created_by": user_id,
            "created": time.time(),
            "revision": 0,
        }
        connection.execute(s.workspaces.insert().values(**value))
        connection.execute(
            s.memberships.insert().values(workspace_id=value["id"], user_id=user_id, role="owner")
        )
        self.audit(connection, user_id, "workspace.created", value["id"], value["id"], name=name)
        return {**value, "role": "owner"}

    def _user(
        self, connection: Any, name: str, password: str, *, is_admin: bool, setup_required: bool
    ) -> dict:
        salt, digest = password_hash(password)
        value: dict[str, Any] = {
            "id": uuid4().hex,
            "username": username(name),
            "salt": salt,
            "digest": digest,
            "active": True,
            "is_admin": is_admin,
            "setup_required": setup_required,
            "created": time.time(),
            "revision": 0,
        }
        if connection.execute(
            select(s.users.c.id).where(s.users.c.username == value["username"])
        ).first():
            raise Conflict("That username is already in use")
        connection.execute(s.users.insert().values(**value))
        first_workspace = connection.execute(select(s.workspaces.c.id).limit(1)).first() is None
        workspace = self._workspace(connection, value["id"], value["username"] + " workspace")
        if first_workspace:
            connection.execute(
                s.entries.update()
                .where(s.entries.c.workspace_id.is_(None))
                .values(workspace_id=workspace["id"])
            )
        return public_user(value)

    def create_user(
        self,
        name: str,
        password: str,
        *,
        is_admin: bool | None = None,
        setup_required: bool = False,
        actor: str | None = None,
    ) -> dict:
        """Create an account and its private workspace, preserving existing evidence."""
        with self.transaction() as connection:
            if actor:
                self.administrator(actor, connection)
            first = connection.execute(select(s.users.c.id).limit(1)).first() is None
            if first and is_admin is False:
                raise ValueError("The first account must be a server administrator")
            value = self._user(
                connection,
                name,
                password,
                is_admin=first if is_admin is None else is_admin,
                setup_required=setup_required,
            )
            self.audit(
                connection,
                actor or value["id"],
                "account.created",
                value["id"],
                username=value["username"],
            )
            return value

    def ensure_bootstrap(self, root_labels: Any = ()) -> bool:
        """Seed admin/admin exactly once; never add a default account to an existing install."""
        with self.transaction() as connection:
            created = connection.execute(select(s.users.c.id).limit(1)).first() is None
            if created:
                self._user(connection, "admin", "admin", is_admin=True, setup_required=True)
            initialized = connection.execute(
                select(s.entries.c.id).where(s.entries.c.id == "server-root-grants-initialized")
            ).first()
            if not initialized:
                workspace = connection.execute(
                    select(s.workspaces.c.id)
                    .order_by(s.workspaces.c.created, s.workspaces.c.id)
                    .limit(1)
                ).scalar_one()
                for label in root_labels:
                    connection.execute(
                        s.workspace_roots.insert().values(workspace_id=workspace, label=label)
                    )
                connection.execute(
                    s.entries.insert().values(
                        id="server-root-grants-initialized",
                        kind="configuration",
                        revision=0,
                        data={},
                        created=time.time(),
                    )
                )
            return created

    def users(self) -> list[dict]:
        """List account metadata without password salts, hashes or session tokens."""
        with self.engine.connect() as connection:
            return [
                public_user(row)
                for row in connection.execute(
                    select(s.users).order_by(s.users.c.username)
                ).mappings()
            ]

    def user(self, user_id: str, connection: Any = None) -> dict:
        """Load a stable account identity for internal authorization decisions."""
        if connection is None:
            with self.engine.connect() as connection:
                return self.user(user_id, connection)
        row = connection.execute(select(s.users).where(s.users.c.id == user_id)).mappings().first()
        if row is None:
            raise KeyError("Account not found")
        return dict(row)

    def login(self, name: str, password: str) -> tuple[str, str]:
        """Normalize the username, verify the password and issue a hashed session."""
        try:
            name = username(name)
        except ValueError:
            name = ""
        with self.engine.begin() as connection:
            row = (
                connection.execute(select(s.users).where(s.users.c.username == name))
                .mappings()
                .first()
            )
            salt = row["salt"] if row else "00" * 16
            _, digest = password_hash(password, salt)
            if not row or not secrets.compare_digest(row["digest"], digest) or not row["active"]:
                raise ValueError("Invalid username or password")
            token, csrf = secrets.token_urlsafe(32), secrets.token_urlsafe(32)
            connection.execute(delete(s.sessions).where(s.sessions.c.expires <= time.time()))
            connection.execute(
                s.sessions.insert().values(
                    id=hashlib.sha256(token.encode()).hexdigest(),
                    user_id=row["id"],
                    csrf=csrf,
                    expires=time.time() + 43200,
                )
            )
            return token, csrf

    def session(self, token: str) -> dict:
        """Reject expired sessions and disabled accounts on every request."""
        with self.engine.connect() as connection:
            row = (
                connection.execute(
                    select(s.sessions).where(
                        s.sessions.c.id == hashlib.sha256(token.encode()).hexdigest(),
                        s.sessions.c.expires > time.time(),
                    )
                )
                .mappings()
                .first()
            )
            if row is None:
                raise KeyError("Session expired")
            user = self.user(row["user_id"], connection)
            if not user["active"]:
                raise KeyError("Account disabled")
            return {
                "id": row["id"],
                "user": public_user(user),
                "data": {
                    "name": user["username"],
                    "user_id": user["id"],
                    "csrf": row["csrf"],
                    "expires": row["expires"],
                },
            }

    def logout(self, identity: str) -> None:
        """Revoke a session immediately rather than merely deleting its cookie."""
        with self.engine.begin() as connection:
            connection.execute(delete(s.sessions).where(s.sessions.c.id == identity))

    def credentials(
        self,
        user_id: str,
        name: str,
        password: str,
        *,
        current_password: str | None = None,
        reset: bool = False,
        actor: str | None = None,
    ) -> None:
        """Replace credentials atomically and invalidate all earlier sessions."""
        with self.transaction() as connection:
            user = self.user(user_id, connection)
            if reset and actor:
                self.administrator(actor, connection)
            if not reset and not user["setup_required"]:
                if not current_password or not secrets.compare_digest(
                    password_hash(current_password, user["salt"])[1], user["digest"]
                ):
                    raise Forbidden("The current password is incorrect")
            name = username(name)
            if not reset and user["setup_required"] and name == "admin" and password == "admin":
                raise ValueError("Choose your own account credentials to finish setup")
            if connection.execute(
                select(s.users.c.id).where(s.users.c.username == name, s.users.c.id != user_id)
            ).first():
                raise Conflict("That username is already in use")
            salt, digest = password_hash(password)
            if user["setup_required"] and name != user["username"]:
                connection.execute(
                    s.workspaces.update()
                    .where(
                        s.workspaces.c.created_by == user_id,
                        s.workspaces.c.name == user["username"] + " workspace",
                    )
                    .values(name=name + " workspace", revision=s.workspaces.c.revision + 1)
                )
            connection.execute(
                s.users.update()
                .where(s.users.c.id == user_id)
                .values(
                    username=name,
                    salt=salt,
                    digest=digest,
                    setup_required=reset,
                    revision=user["revision"] + 1,
                )
            )
            connection.execute(delete(s.sessions).where(s.sessions.c.user_id == user_id))
            self.audit(
                connection,
                actor or user_id,
                "account.password_reset" if reset else "account.updated",
                user_id,
            )

    def update_user(
        self, actor: str, user_id: str, *, revision: int, active: bool, is_admin: bool
    ) -> dict:
        """Apply account privileges while preserving an active administrator."""
        with self.transaction() as connection:
            self.administrator(actor, connection)
            user = self.user(user_id, connection)
            if user["revision"] != revision:
                raise Conflict("This account changed; reload before saving")
            if actor == user_id and not active:
                raise ValueError("Use another administrator account to disable your account")
            if user["active"] and user["is_admin"] and not (active and is_admin):
                others = connection.execute(
                    select(func.count())
                    .select_from(s.users)
                    .where(
                        s.users.c.id != user_id,
                        s.users.c.active.is_(True),
                        s.users.c.is_admin.is_(True),
                    )
                ).scalar_one()
                if not others:
                    raise Conflict("Keep at least one active server administrator")
            connection.execute(
                s.users.update()
                .where(s.users.c.id == user_id)
                .values(active=active, is_admin=is_admin, revision=revision + 1)
            )
            if not active:
                connection.execute(delete(s.sessions).where(s.sessions.c.user_id == user_id))
            self.audit(
                connection,
                actor,
                "account.permissions_updated",
                user_id,
                active=active,
                is_admin=is_admin,
            )
            return public_user(self.user(user_id, connection))

    def role(self, user_id: str, workspace_id: str, connection: Any = None) -> str | None:
        """Read current membership rather than trusting a browser or session claim."""
        if connection is None:
            with self.engine.connect() as connection:
                return self.role(user_id, workspace_id, connection)
        result = connection.execute(
            select(s.memberships.c.role)
            .join(s.users)
            .where(
                s.memberships.c.user_id == user_id,
                s.memberships.c.workspace_id == workspace_id,
                s.users.c.active.is_(True),
            )
        ).scalar()
        return str(result) if result is not None else None

    def require(
        self, user_id: str, workspace_id: str, minimum: str = "viewer", connection: Any = None
    ) -> str:
        """Require explicit membership and a sufficient workspace role."""
        role = self.role(user_id, workspace_id, connection)
        if role is None:
            raise KeyError("Workspace not found")
        if ROLES[role] < ROLES[minimum]:
            raise Forbidden(f"This action requires {minimum} access to the workspace")
        return role

    def workspaces(self, user_id: str) -> list[dict]:
        """List only workspaces to which the account currently belongs."""
        with self.engine.connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    select(s.workspaces, s.memberships.c.role)
                    .join(s.memberships)
                    .where(s.memberships.c.user_id == user_id)
                    .order_by(
                        case((s.workspaces.c.created_by == user_id, 0), else_=1),
                        s.workspaces.c.created,
                        s.workspaces.c.id,
                    )
                ).mappings()
            ]

    def create_workspace(self, user_id: str, name: str) -> dict:
        """Create a private workspace with the caller as its first owner."""
        with self.transaction() as connection:
            return self._workspace(connection, user_id, name)

    def workspace_details(self, user_id: str, identity: str, configured: dict[str, str]) -> dict:
        """Bind a sharing form's revision, members and grants to one consistent read."""
        with self.transaction() as connection:
            role = self.require(user_id, identity, connection=connection)
            workspace = (
                connection.execute(select(s.workspaces).where(s.workspaces.c.id == identity))
                .mappings()
                .one()
            )
            members = [
                dict(row)
                for row in connection.execute(
                    select(s.users.c.id, s.users.c.username, s.users.c.active, s.memberships.c.role)
                    .join(s.memberships)
                    .where(s.memberships.c.workspace_id == identity)
                    .order_by(s.users.c.username)
                ).mappings()
            ]
            roots = [
                label
                for label in connection.execute(
                    select(s.workspace_roots.c.label).where(
                        s.workspace_roots.c.workspace_id == identity
                    )
                ).scalars()
                if label in configured
            ]
            return {**dict(workspace), "role": role, "members": members, "roots": roots}

    def rename_workspace(self, actor: str, identity: str, name: str, revision: int) -> None:
        """Rename a workspace while detecting concurrent changes."""
        name = name.strip()
        if not name or len(name) > 120:
            raise ValueError("Choose a workspace name between 1 and 120 characters")
        with self.transaction() as connection:
            self.require(actor, identity, "owner", connection)
            result = connection.execute(
                s.workspaces.update()
                .where(s.workspaces.c.id == identity, s.workspaces.c.revision == revision)
                .values(name=name, revision=revision + 1)
            )
            if result.rowcount != 1:
                raise Conflict("This workspace changed; reload before saving")
            self.audit(connection, actor, "workspace.renamed", identity, identity, name=name)

    def members(self, user_id: str, identity: str) -> list[dict]:
        """Expose the membership list only to existing workspace members."""
        self.require(user_id, identity)
        with self.engine.connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    select(s.users.c.id, s.users.c.username, s.users.c.active, s.memberships.c.role)
                    .join(s.memberships)
                    .where(s.memberships.c.workspace_id == identity)
                    .order_by(s.users.c.username)
                ).mappings()
            ]

    def member(
        self, actor: str, identity: str, name: str, role: str | None, revision: int | None = None
    ) -> None:
        """Change sharing without allowing the last active owner to be removed."""
        if role is not None and role not in ROLES:
            raise ValueError("Choose owner, editor or viewer access")
        with self.transaction() as connection:
            self.require(actor, identity, "owner", connection)
            self.sharing_revision(connection, identity, revision)
            user = (
                connection.execute(select(s.users).where(s.users.c.username == username(name)))
                .mappings()
                .first()
            )
            if user is None or (role is not None and not user["active"]):
                raise ValueError("An active account with that username was not found")
            previous = self.role(user["id"], identity, connection)
            if previous == "owner" and role != "owner":
                others = connection.execute(
                    select(func.count())
                    .select_from(s.memberships.join(s.users))
                    .where(
                        s.memberships.c.workspace_id == identity,
                        s.memberships.c.role == "owner",
                        s.memberships.c.user_id != user["id"],
                        s.users.c.active.is_(True),
                    )
                ).scalar_one()
                if not others:
                    raise Conflict("Keep at least one active workspace owner")
            connection.execute(
                delete(s.memberships).where(
                    s.memberships.c.workspace_id == identity, s.memberships.c.user_id == user["id"]
                )
            )
            if role:
                connection.execute(
                    s.memberships.insert().values(
                        workspace_id=identity, user_id=user["id"], role=role
                    )
                )
            self.audit(
                connection,
                actor,
                "workspace.member_updated",
                user["id"],
                identity,
                username=user["username"],
                role=role,
            )

    def roots(self, workspace_id: str, configured: dict[str, str]) -> dict[str, str]:
        """Return configured server folders explicitly granted to this workspace."""
        with self.engine.connect() as connection:
            labels = connection.execute(
                select(s.workspace_roots.c.label).where(
                    s.workspace_roots.c.workspace_id == workspace_id
                )
            ).scalars()
            return {label: configured[label] for label in labels if label in configured}

    def grant_roots(
        self,
        actor: str,
        workspace_id: str,
        labels: list[str],
        configured: dict[str, str],
        revision: int | None = None,
    ) -> None:
        """Require both server administration and workspace ownership for folder grants."""
        if not set(labels) <= configured.keys():
            raise ValueError("Unknown configured source folder")
        with self.transaction() as connection:
            self.require(actor, workspace_id, "owner", connection)
            self.sharing_revision(connection, workspace_id, revision)
            self.administrator(actor, connection)
            connection.execute(
                delete(s.workspace_roots).where(s.workspace_roots.c.workspace_id == workspace_id)
            )
            for label in set(labels):
                connection.execute(
                    s.workspace_roots.insert().values(workspace_id=workspace_id, label=label)
                )
            self.audit(
                connection,
                actor,
                "workspace.source_folders_updated",
                workspace_id,
                workspace_id,
                labels=labels,
            )

    def administrator(self, identity: str, connection: Any) -> None:
        """Recheck administrator status inside the protected mutation transaction."""
        user = self.user(identity, connection)
        if not user["active"] or not user["is_admin"]:
            raise Forbidden("Server administrator access is required")

    def sharing_revision(self, connection: Any, identity: str, expected: int | None) -> None:
        """Reject stale access changes in the same transaction as the membership update."""
        query = s.workspaces.update().where(s.workspaces.c.id == identity)
        if expected is not None:
            query = query.where(s.workspaces.c.revision == expected)
        if connection.execute(query.values(revision=s.workspaces.c.revision + 1)).rowcount != 1:
            raise Conflict("Workspace sharing changed; reload before saving")

    def history(self, user_id: str, workspace_id: str, after: int = 0) -> list[dict]:
        """Return a bounded, ordered audit history to workspace owners."""
        self.require(user_id, workspace_id, "owner")
        with self.engine.connect() as connection:
            return [
                dict(row)
                for row in connection.execute(
                    select(s.audit, s.users.c.username)
                    .join(s.users, s.audit.c.actor_id == s.users.c.id)
                    .where(s.audit.c.workspace_id == workspace_id, s.audit.c.sequence > after)
                    .order_by(s.audit.c.sequence)
                    .limit(200)
                ).mappings()
            ]
