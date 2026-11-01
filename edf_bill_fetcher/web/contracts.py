"""Explicit HTTP job contracts; no web-specific financial calculations."""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator

from edf_bill_fetcher.application.configuration import validate_settings
from edf_bill_fetcher.models.report_sections import REPORT_SECTIONS

SourceKind = Literal[
    "auto", "pdf_dir", "eml_dir", "msg_dir", "pst", "ost", "htm", "result", "workbook", "bundle"
]


class Contract(BaseModel):
    """Reject misspelled fields rather than silently ignoring requested behaviour."""

    model_config = ConfigDict(extra="forbid", strict=True)


class Source(Contract):
    """An uploaded collection or a contained configured server location."""

    asset_id: str | None = None
    intake_id: str | None = None
    root_id: str | None = None
    relative: str | None = None
    kind: SourceKind

    @model_validator(mode="after")
    def location(self):
        """Require exactly one source location."""
        if sum(bool(value) for value in (self.asset_id, self.root_id, self.intake_id)) != 1 or (
            (self.root_id or self.intake_id) and not self.relative
        ):
            raise ValueError("Select an upload or a path under one configured source root")
        return self


class Job(Contract):
    """Captured immutable job request, independent of browser connection lifetime."""

    mode: Literal["intake", "extract", "render", "workbook", "review", "preview", "diff"]
    name: str = Field(default="Evidence", min_length=1, max_length=120)
    sources: list[Source] = Field(default_factory=list, max_length=100)
    parent_id: str | None = None
    other_id: str | None = None
    input: Source | None = None
    config: dict[str, Any] = Field(default_factory=dict)
    formats: list[Literal["xlsx", "pdf", "docx", "html"]] = Field(
        default=["xlsx", "pdf", "docx", "html"], min_length=1, max_length=4
    )
    sections: list[str] | None = None
    draft: dict | None = None

    @model_validator(mode="after")
    def valid(self):
        """Validate bounded settings before reserving or starting a job."""
        if self.mode in {"extract", "intake"} and (
            not self.sources
            or any(s.kind in {"result", "bundle", "workbook"} for s in self.sources)
        ):
            raise ValueError("Extraction requires billing sources")
        if self.mode not in {"extract", "intake"} and not (self.parent_id or self.input):
            raise ValueError("Select a saved run or an imported input")
        if self.mode == "diff" and not self.other_id:
            raise ValueError("Select a second run")
        if self.mode in {"review", "preview"} and self.draft is None:
            raise ValueError("Supply a captured review draft")
        if self.mode in {"review", "preview"} and not self.parent_id:
            raise ValueError("Import the analysis as a saved run before reviewing it")
        if len(set(self.formats)) != len(self.formats):
            raise ValueError("Report formats must be unique")
        if self.sections is None and "report_sections" in self.config:
            value = self.config["report_sections"]
            if not isinstance(value, list) or any(not isinstance(key, str) for key in value):
                raise ValueError("Report sections must be a list of section keys")
            self.sections = value
        if self.sections is not None and (
            len(set(self.sections)) != len(self.sections)
            or not set(self.sections) <= ({s.key for s in REPORT_SECTIONS} | {"cover", "toc"})
        ):
            raise ValueError("Unknown or duplicate report section")
        validate_settings(self.config)
        return self


class Credentials(Contract):
    """Bound password processing input independently of HTTP body size."""

    username: str = Field(min_length=1, max_length=100)
    password: str = Field(min_length=1, max_length=1024)


class Asset(Contract):
    """Named upload collection, finalized before extraction may read it."""

    name: str = Field(min_length=1, max_length=120)
    kind: SourceKind


class Edit(Contract):
    """Optimistic revision is required for every review mutation."""

    revision: int = Field(ge=0)
    operation: str
    arguments: dict = Field(default_factory=dict)


class UploadDigest(Contract):
    """Expected original bytes after streaming the client file."""

    size: int = Field(ge=0, le=64 * 1024**3)
    sha256: str = Field(pattern=r"^[0-9a-f]{64}$")


class Seal(Contract):
    """Browser clients verify original hashes before finalizing an upload."""

    files: dict[str, UploadDigest] = Field(default_factory=dict)
    batch_id: str | None = Field(default=None, max_length=64)
    files_count: int | None = Field(default=None, ge=1, le=100_000)


class CompleteFile(UploadDigest):
    """One bounded completion request, independent of the number of folder members."""

    batch_id: str = Field(min_length=1, max_length=64)
    order: int = Field(ge=0, lt=100_000)
