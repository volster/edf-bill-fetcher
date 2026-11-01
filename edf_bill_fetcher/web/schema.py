"""Relational catalogue tables shared by PostgreSQL and local SQLite installations."""

from sqlalchemy import (
    JSON,
    BigInteger,
    Boolean,
    CheckConstraint,
    Column,
    Float,
    ForeignKey,
    Index,
    Integer,
    MetaData,
    String,
    Table,
)

metadata = MetaData()
users = Table(
    "web_users",
    metadata,
    Column("id", String, primary_key=True),
    Column("username", String(100), nullable=False, unique=True),
    Column("salt", String(32), nullable=False),
    Column("digest", String(64), nullable=False),
    Column("active", Boolean, nullable=False),
    Column("is_admin", Boolean, nullable=False),
    Column("setup_required", Boolean, nullable=False),
    Column("created", Float, nullable=False),
    Column("revision", Integer, nullable=False, default=0),
)
workspaces = Table(
    "web_workspaces",
    metadata,
    Column("id", String, primary_key=True),
    Column("name", String(120), nullable=False),
    Column("created_by", String, ForeignKey("web_users.id"), nullable=False),
    Column("created", Float, nullable=False),
    Column("revision", Integer, nullable=False, default=0),
)
memberships = Table(
    "web_memberships",
    metadata,
    Column(
        "workspace_id",
        String,
        ForeignKey("web_workspaces.id", ondelete="CASCADE"),
        primary_key=True,
    ),
    Column("user_id", String, ForeignKey("web_users.id", ondelete="CASCADE"), primary_key=True),
    Column("role", String(20), nullable=False),
    CheckConstraint("role IN ('owner', 'editor', 'viewer')", name="valid_workspace_role"),
)
workspace_roots = Table(
    "web_workspace_roots",
    metadata,
    Column(
        "workspace_id",
        String,
        ForeignKey("web_workspaces.id", ondelete="CASCADE"),
        primary_key=True,
    ),
    Column("label", String(120), primary_key=True),
)
sessions = Table(
    "web_sessions",
    metadata,
    Column("id", String(64), primary_key=True),
    Column(
        "user_id",
        String,
        ForeignKey("web_users.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    ),
    Column("csrf", String(64), nullable=False),
    Column("expires", Float, nullable=False),
)
entries = Table(
    "entries",
    metadata,
    Column("id", String, primary_key=True),
    Column("kind", String, nullable=False, index=True),
    Column("revision", Integer, nullable=False),
    Column("data", JSON, nullable=False),
    Column("workspace_id", String, ForeignKey("web_workspaces.id"), nullable=True),
    Column("created", Float, nullable=False, default=0),
    Index("entries_workspace_kind", "workspace_id", "kind"),
)
events = Table(
    "web_events",
    metadata,
    Column(
        "sequence",
        BigInteger().with_variant(Integer, "sqlite"),
        primary_key=True,
        autoincrement=True,
    ),
    Column("job_id", String, nullable=False, index=True),
    Column("message", String(4000), nullable=False),
    Column("at", Float, nullable=False),
)
audit = Table(
    "web_audit",
    metadata,
    Column(
        "sequence",
        BigInteger().with_variant(Integer, "sqlite"),
        primary_key=True,
        autoincrement=True,
    ),
    Column("workspace_id", String, ForeignKey("web_workspaces.id"), nullable=True, index=True),
    Column("actor_id", String, ForeignKey("web_users.id"), nullable=False),
    Column("action", String(100), nullable=False),
    Column("target", String, nullable=False),
    Column("details", JSON, nullable=False),
    Column("at", Float, nullable=False),
)
