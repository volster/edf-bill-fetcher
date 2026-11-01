"""Create the durable server catalogue without changing evidence interchange files."""

from alembic import op
from sqlalchemy import JSON, Column, Integer, String

revision = "0001_catalogue"
down_revision = None


def upgrade():
    """Install the first supported catalogue schema."""
    op.create_table(
        "entries",
        Column("id", String, primary_key=True),
        Column("kind", String, nullable=False, index=True),
        Column("revision", Integer, nullable=False),
        Column("data", JSON, nullable=False),
    )


def downgrade():
    """Remove catalogue metadata only; application evidence directories remain untouched."""
    op.drop_table("entries")
