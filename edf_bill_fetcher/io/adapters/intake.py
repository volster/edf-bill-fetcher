"""Bounded mixed-evidence discovery with explicit unsupported-file and archive outcomes."""

from __future__ import annotations

import hashlib
import os
import shutil
from collections.abc import Callable, Iterator
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from zipfile import BadZipFile, ZipFile, ZipInfo

from edf_bill_fetcher.io.persistence.run_manifest import artifact_record, is_run_directory


@dataclass(frozen=True)
class IntakeLimits:
    """Bound aggregate archive expansion and discovery independently of file uploads."""

    files: int = 100_000
    expanded_bytes: int = 64 * 1024**3
    depth: int = 3


def detect(name: str, header: bytes) -> str:
    """Use signatures first; ambiguous or unsupported content is never guessed into a bill."""
    suffix = Path(name).suffix.lower()
    if header.lstrip().startswith(b"%PDF-"):
        return "pdf"
    if header.startswith(b"!BDN"):
        return "pst"
    if header.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1") and suffix == ".msg":
        return "msg"
    if header.startswith((b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")):
        return "workbook" if suffix == ".xlsx" else "zip"
    text = header.decode("utf-8-sig", errors="replace").lstrip().lower()
    if text.startswith(("<!doctype html", "<html")) or suffix in {".htm", ".html"} and "<" in text:
        return "htm"
    if suffix == ".eml" or ("\nfrom:" in "\n" + text and "\nsubject:" in "\n" + text):
        return "eml"
    if suffix == ".json" and text.startswith(("{", "[")):
        return "result"
    return "unsupported"


def _member(name: str) -> PurePosixPath:
    path = PurePosixPath(name)
    if (
        not name
        or path.is_absolute()
        or ".." in path.parts
        or "\\" in name
        or ":" in name
        or any(ord(char) < 32 for char in name)
    ):
        raise ValueError("Unsafe ZIP member path")
    return path


def validate_zip_members(members: list[ZipInfo]) -> set[str]:
    """Reject unsafe or conflicting archive paths before writing any original."""
    files: set[str] = set()
    directories: set[str] = set()
    for info in members:
        member = _member(info.orig_filename)
        portable = member.as_posix().casefold()
        if (info.external_attr >> 16) & 0o170000 == 0o120000 or info.flag_bits & 1:
            raise ValueError("Encrypted or symbolic-link ZIP members require manual handling")
        parents = {parent.as_posix().casefold() for parent in member.parents if str(parent) != "."}
        if (
            portable == "."
            or parents & files
            or portable in files
            or (not info.is_dir() and portable in directories)
        ):
            raise ValueError("ZIP contains conflicting member names")
        directories.update(parents)
        if info.is_dir():
            directories.add(portable)
        else:
            files.add(portable)
    return files


def is_pdf_candidate(path: str) -> bool:
    """Preserve PDF error handling and preview renamed PDFs by their content."""
    if Path(path).suffix.lower() == ".pdf":
        return True
    try:
        with Path(path).open("rb") as source:
            return detect(Path(path).name, source.read(8192)) == "pdf"
    except OSError:
        return False


def discover(
    path: Path,
    staging: Path,
    *,
    cancelled: Callable[[], bool] = lambda: False,
    limits: IntakeLimits = IntakeLimits(),
) -> Iterator[dict]:
    """Yield classified originals; keep expanded members alive until source publication."""
    count, expanded = 0, 0
    root = path if path.is_dir() else path.parent

    def inspect(
        current: Path, relative: str, depth: int = 0, container: dict | None = None
    ) -> Iterator[dict]:
        nonlocal count, expanded
        if cancelled():
            return
        count += 1
        if count > limits.files:
            raise ValueError("Mixed intake exceeds the file-count limit")
        value: dict = {
            "path": str(current),
            "name": relative,
            "kind": "unsupported",
            "status": "unsupported",
            "detail": "",
            "container": container,
        }
        try:
            with current.open("rb") as source:
                header = source.read(8192)
            stat = current.stat()
            value.update(size=stat.st_size, mtime_ns=stat.st_mtime_ns, origin=str(path.resolve()))
            kind = detect(current.name, header)
            value.update(
                kind=kind,
                status="ready"
                if kind in {"pdf", "pst", "msg", "eml", "htm"}
                else "needs_choice"
                if kind in {"result", "workbook"}
                else "unsupported",
            )
            if kind != "zip":
                if kind == "unsupported":
                    value["detail"] = (
                        "Unsupported file retained in the intake inventory; no billing values inferred"
                    )
                yield value
                return
            if depth >= limits.depth:
                raise ValueError("Nested ZIP depth limit exceeded")
            with ZipFile(current) as archive:
                members = archive.infolist()
                if any(PurePosixPath(info.filename).name == "result.json" for info in members):
                    value.update(
                        kind="bundle",
                        status="needs_choice",
                        detail="Saved analysis pack: import it separately to avoid counting its generated reports as new evidence",
                    )
                    yield value
                    return
                if len(members) + count > limits.files:
                    raise ValueError("ZIP exceeds the file-count limit")
                if expanded + sum(info.file_size for info in members) > limits.expanded_bytes:
                    raise ValueError("ZIP expansion exceeds the aggregate byte limit")
                # Validate the whole archive before writing any member.
                names = validate_zip_members(members)
                folder = staging / hashlib.sha256(relative.encode()).hexdigest()
                folder.mkdir(parents=True, exist_ok=True)
                if sum(info.file_size for info in members) > shutil.disk_usage(folder).free:
                    raise ValueError("Not enough free space to expand this ZIP")
                container_hash = artifact_record(str(current))["sha256"]
                value.update(
                    kind="zip",
                    status="container",
                    detail=f"{len(names)} members; contents classified individually",
                )
                yield value
                for info in members:
                    if cancelled():
                        return
                    if info.is_dir():
                        continue
                    target = folder / _member(info.filename)
                    target.parent.mkdir(parents=True, exist_ok=True)
                    with archive.open(info) as source, target.open("xb") as output:
                        for block in iter(lambda: source.read(1024 * 1024), b""):
                            if cancelled():
                                return
                            expanded += len(block)
                            if expanded > limits.expanded_bytes:
                                raise ValueError("ZIP expansion exceeds the aggregate byte limit")
                            output.write(block)
                    yield from inspect(
                        target,
                        relative + "!/" + info.filename,
                        depth + 1,
                        {
                            "Container": relative,
                            "Container SHA256": container_hash,
                            "Member": info.filename,
                            "Parent Container": container,
                        },
                    )
        except (OSError, ValueError, BadZipFile, RuntimeError) as error:
            value.update(status="failed", detail=str(error))
            yield value

    if path.is_file():
        yield from inspect(path, path.name)
        return
    errors: list[OSError] = []
    for base, directories, files in os.walk(path, onerror=errors.append, followlinks=False):
        if cancelled():
            return
        if is_run_directory(base):
            directories.clear()
            yield {
                "path": base,
                "name": Path(base).relative_to(root).as_posix(),
                "kind": "prior_output",
                "status": "skipped",
                "detail": "Earlier output pack; not re-extracted",
                "container": None,
            }
            continue
        directories[:] = sorted(
            directory for directory in directories if not Path(base, directory).is_symlink()
        )
        for name in sorted(files):
            if cancelled():
                return
            current = Path(base, name)
            if current.is_symlink():
                yield {
                    "path": str(current),
                    "name": current.relative_to(root).as_posix(),
                    "kind": "symlink",
                    "status": "skipped",
                    "detail": "Symbolic link not followed",
                    "container": None,
                }
            else:
                yield from inspect(current, current.relative_to(root).as_posix())
    for error in errors:
        yield {
            "path": str(error.filename),
            "name": str(error.filename),
            "kind": "folder",
            "status": "failed",
            "detail": str(error),
            "container": None,
        }
