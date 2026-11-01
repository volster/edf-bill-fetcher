"""Install accounts, memberships, scoped evidence and portable event sequences."""

import time
import unicodedata
from uuid import uuid4

import sqlalchemy as sa
from alembic import op

revision = "0002_workspaces"
down_revision = "0001_catalogue"


def upgrade():
    """Preserve existing accounts and evidence in a shared imported workspace."""
    original = sa.table("entries", sa.column("kind"), sa.column("data", sa.JSON()))
    names = [
        unicodedata.normalize("NFKC", row[0]["name"]).strip().casefold()
        for row in op.get_bind().execute(
            sa.select(original.c.data).where(original.c.kind == "user")
        )
    ]
    if any(not name or len(name) > 100 or any(ord(c) < 32 for c in name) for name in names) or len(
        set(names)
    ) != len(names):
        raise ValueError(
            "Existing usernames are invalid or collide after normalization; resolve them in the legacy catalogue before upgrading"
        )
    op.create_table(
        "web_users",
        sa.Column("id", sa.String(), primary_key=True),
        sa.Column("username", sa.String(100), nullable=False, unique=True),
        sa.Column("salt", sa.String(32), nullable=False),
        sa.Column("digest", sa.String(64), nullable=False),
        sa.Column("active", sa.Boolean(), nullable=False),
        sa.Column("is_admin", sa.Boolean(), nullable=False),
        sa.Column("setup_required", sa.Boolean(), nullable=False),
        sa.Column("created", sa.Float(), nullable=False),
        sa.Column("revision", sa.Integer(), nullable=False),
    )
    op.create_table(
        "web_workspaces",
        sa.Column("id", sa.String(), primary_key=True),
        sa.Column("name", sa.String(120), nullable=False),
        sa.Column("created_by", sa.String(), sa.ForeignKey("web_users.id"), nullable=False),
        sa.Column("created", sa.Float(), nullable=False),
        sa.Column("revision", sa.Integer(), nullable=False),
    )
    op.create_table(
        "web_memberships",
        sa.Column(
            "workspace_id",
            sa.String(),
            sa.ForeignKey("web_workspaces.id", ondelete="CASCADE"),
            primary_key=True,
        ),
        sa.Column(
            "user_id",
            sa.String(),
            sa.ForeignKey("web_users.id", ondelete="CASCADE"),
            primary_key=True,
        ),
        sa.Column("role", sa.String(20), nullable=False),
        sa.CheckConstraint("role IN ('owner', 'editor', 'viewer')", name="valid_workspace_role"),
    )
    op.create_table(
        "web_workspace_roots",
        sa.Column(
            "workspace_id",
            sa.String(),
            sa.ForeignKey("web_workspaces.id", ondelete="CASCADE"),
            primary_key=True,
        ),
        sa.Column("label", sa.String(120), primary_key=True),
    )
    op.create_table(
        "web_sessions",
        sa.Column("id", sa.String(64), primary_key=True),
        sa.Column(
            "user_id",
            sa.String(),
            sa.ForeignKey("web_users.id", ondelete="CASCADE"),
            nullable=False,
        ),
        sa.Column("csrf", sa.String(64), nullable=False),
        sa.Column("expires", sa.Float(), nullable=False),
    )
    op.create_index("ix_web_sessions_user_id", "web_sessions", ["user_id"])
    with op.batch_alter_table("entries") as batch:
        batch.add_column(sa.Column("workspace_id", sa.String(), nullable=True))
        batch.add_column(sa.Column("created", sa.Float(), nullable=False, server_default="0"))
        batch.create_foreign_key("entries_workspace_fk", "web_workspaces", ["workspace_id"], ["id"])
        batch.create_index("entries_workspace_kind", ["workspace_id", "kind"])
    op.create_table(
        "web_events",
        sa.Column(
            "sequence",
            sa.BigInteger().with_variant(sa.Integer, "sqlite"),
            primary_key=True,
            autoincrement=True,
        ),
        sa.Column("job_id", sa.String(), nullable=False),
        sa.Column("message", sa.String(4000), nullable=False),
        sa.Column("at", sa.Float(), nullable=False),
    )
    op.create_index("ix_web_events_job_id", "web_events", ["job_id"])
    op.create_table(
        "web_audit",
        sa.Column(
            "sequence",
            sa.BigInteger().with_variant(sa.Integer, "sqlite"),
            primary_key=True,
            autoincrement=True,
        ),
        sa.Column("workspace_id", sa.String(), sa.ForeignKey("web_workspaces.id"), nullable=True),
        sa.Column("actor_id", sa.String(), sa.ForeignKey("web_users.id"), nullable=False),
        sa.Column("action", sa.String(100), nullable=False),
        sa.Column("target", sa.String(), nullable=False),
        sa.Column("details", sa.JSON(), nullable=False),
        sa.Column("at", sa.Float(), nullable=False),
    )
    op.create_index("ix_web_audit_workspace_id", "web_audit", ["workspace_id"])
    connection = op.get_bind()
    metadata = sa.MetaData()
    tables = {
        name: sa.Table(name, metadata, autoload_with=connection)
        for name in ["entries", "web_users", "web_workspaces", "web_memberships", "web_events"]
    }
    entries = tables["entries"]
    ordering = sa.literal_column("rowid") if connection.dialect.name == "sqlite" else entries.c.id
    legacy = list(
        connection.execute(
            sa.select(entries).where(entries.c.kind == "user").order_by(ordering)
        ).mappings()
    )
    workspace_id = uuid4().hex
    for index, row in enumerate(legacy):
        user = row["data"]
        username = unicodedata.normalize("NFKC", user["name"]).strip().casefold()
        connection.execute(
            tables["web_users"]
            .insert()
            .values(
                id=row["id"],
                username=username,
                salt=user["salt"],
                digest=user["digest"],
                active=True,
                is_admin=index == 0,
                setup_required=False,
                created=time.time(),
                revision=0,
            )
        )
        if index == 0:
            connection.execute(
                tables["web_workspaces"]
                .insert()
                .values(
                    id=workspace_id,
                    name="Imported workspace",
                    created_by=row["id"],
                    created=time.time(),
                    revision=0,
                )
            )
        connection.execute(
            tables["web_memberships"]
            .insert()
            .values(workspace_id=workspace_id, user_id=row["id"], role="owner")
        )
    if legacy:
        connection.execute(entries.update().values(workspace_id=workspace_id))
    for row in connection.execute(sa.select(entries).where(entries.c.kind == "job")).mappings():
        if row["data"].get("status") in {"queued", "running"}:
            connection.execute(
                entries.update()
                .where(entries.c.id == row["id"])
                .values(
                    data={
                        **row["data"],
                        "status": "interrupted",
                        "error": "Queued before workspace permissions were introduced. Create a fresh run under your account.",
                    }
                )
            )
    ordering = sa.literal_column("rowid") if connection.dialect.name == "sqlite" else entries.c.id
    for row in connection.execute(
        sa.select(entries).where(entries.c.kind == "event").order_by(ordering)
    ).mappings():
        data = row["data"]
        connection.execute(
            tables["web_events"]
            .insert()
            .values(job_id=data["job"], message=data["message"], at=data["at"])
        )
    # Existing browsers sign in again; password hashes and evidence are retained.
    connection.execute(entries.delete().where(entries.c.kind.in_(["user", "session", "event"])))


def downgrade():
    """Require restoring a backup rather than discarding account and access history."""
    raise RuntimeError("Workspace migrations require a database backup to downgrade")
