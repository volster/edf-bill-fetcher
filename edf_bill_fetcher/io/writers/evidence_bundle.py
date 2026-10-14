"""Save referenced evidence files and build a themed DOCX bundle index.

Canonical home: ``edf_bill_fetcher.io.writers.evidence_bundle`` (absorbed
from the top-level ``evidence_bundle.py`` during the shim-removal wave).

This module implements Stream P5 of the SAP/Evidence Bundle design
(``scratch/superpowers/specs/2026-07-16-sap-dumps-and-evidence-bundle-design.md``):

- ``save_evidence_files(expr_df, source_paths, dest_dir)`` copies into a flat
  ``dest_dir`` every file the Evidence Report references via its
  ``Attachment Name`` column. Collisions are handled by appending ``-2``,
  ``-3`` etc.
- ``build_bundle_index(evidence_df, saved_files, output_path)`` writes a
  themed-section DOCX bundle index using python-docx. The section-to-row
  mapping follows a 3-layer inference (prefix / content fingerprint / default
  with the ``⚠  unprefixed`` flag).
"""

from __future__ import annotations

import hashlib
import os
import re
import shutil
from collections import defaultdict
from collections.abc import Callable
from pathlib import Path
from typing import Any
from urllib.parse import quote

import pandas as pd

from edf_bill_fetcher.helpers.artifacts import atomic_output, write_json
from edf_bill_fetcher.io.persistence.run_manifest import artifact_record

# ---------------------------------------------------------------------------
# Section inference (Stream P5, spec §7.3)
# ---------------------------------------------------------------------------

# Section labels -- order matters when we render the bundle.
SECTION_LABELS: dict[str, str] = {
    "A": "A — Ombudsman and Regulatory",
    "B": "B — Correspondence",
    "C": "C — PSR",
    "D": "D — Invoices",
    "E": "E — Meter Readings",
    "F": "F — Balance",
    "G": "G — Calls",
}

# Filename-prefix regex: ``A1 - foo.pdf`` or ``B12 - bar.pdf``. The captured
# group is the upper-case letter identifying the section.
_PREFIX_RE = re.compile(r"^([A-G])\d*\s*-\s*", re.IGNORECASE)

# Content-fingerprint regexes (Stream P5, spec §7.3 layer 2): recognisable
# EDF dump filenames without prefix.
_METER_HISTORY_RE = re.compile(r"Meter-?Read-?History", re.IGNORECASE)
_FINANCIAL_TX_RE = re.compile(r"Financial-?Transactions", re.IGNORECASE)
_CONTRACT_HISTORY_RE = re.compile(r"Contract-?and-?Product-?Change-?History", re.IGNORECASE)


def _infer_section(attachment_name: str, fallback: str = "D") -> tuple[str, bool]:
    """Return ``(section_letter, is_unprefixed)`` for a single row.

    Three-layer inference:
      1. Filename suffix-letter prefix (``A1 -``, ``B12 -``, ...).
      2. Content fingerprint (``[[Meter-Read-History`` → E, etc.).
      3. Default ``D — Invoices`` -- the unrouted-but-categorised catchall.
         Unprefixed files are flagged so the bundle can display a
         ``⚠  unprefixed`` marker.
    """
    name = str(attachment_name or "")
    base = os.path.basename(name)
    m = _PREFIX_RE.match(base)
    if m:
        letter = m.group(1).upper()
        if letter in SECTION_LABELS:
            return letter, False
    # Layer 2: fingerprint by filename + context.
    if _METER_HISTORY_RE.search(base):
        # Meter-Read-History has no £/kWh readings per-row -- route to E.
        return "E", False
    if _FINANCIAL_TX_RE.search(base):
        return "F", False
    if _CONTRACT_HISTORY_RE.search(base):
        return "D", False
    # Layer 3: default-to-D, flagged as unprefixed.
    return fallback, True


# ---------------------------------------------------------------------------
# Save referenced files into a flat dest dir
# ---------------------------------------------------------------------------

_WINDOWS_ILLEGAL_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_WS_RUN_RE = re.compile(r"\s+")


def _text(value: object) -> str:
    text = str(value)
    return "" if text.strip() in {"", "N/A", "None", "nan", "NaT", "<NA>"} else text


@atomic_output()
def _copy_source(source: str, output_path: str, expected_hash: str = "") -> None:
    shutil.copy2(source, output_path)
    if expected_hash:
        digest = hashlib.sha256()
        with open(output_path, "rb") as handle:
            for block in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(block)
        if digest.hexdigest() != expected_hash:
            raise ValueError("source contents changed after extraction")


def sanitise_filename(name: str) -> str:
    """Return a filesystem-safe filename (illegal chars and directory separators neutralised)."""
    base = str(name or "").strip()
    base = _WINDOWS_ILLEGAL_RE.sub("_", base)
    base = _WS_RUN_RE.sub("_", base).strip(" .")
    if re.match(r"^(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)", base, re.I):
        base = "_" + base
    stem, ext = os.path.splitext(base)
    # POSIX filenames are capped in bytes; character limits alone still allow
    # valid Unicode labels to exceed NAME_MAX and silently lose a bundle copy.
    base = stem.encode("utf-8")[:180].decode("utf-8", errors="ignore") + ext.encode("utf-8")[
        :20
    ].decode("utf-8", errors="ignore")
    return base or "attachment"


def save_evidence_files(
    evidence_df: pd.DataFrame,
    source_paths: dict[str, str],
    dest_dir: str,
    log: Callable[[str], Any] | None = None,
    stats: dict[str, int] | None = None,
) -> dict[str, str]:
    """Copy every referred file into ``dest_dir``.

    ``source_paths[attachment_name] -> absolute_path`` is the reverse lookup
    the GUI / CLI builds while walking the local PDF / PST / HTM corpus. Any
    missing path is logged and skipped. Collisions (``Attachment_N.pdf`` in
    multiple PST emails, for example) are deduped by suffixing ``-2``, ``-3``
    etc. on the destination basename.

    Returns a dict ``{attachment_name: destination_path}`` for every file
    successfully copied. Multiple invoice slices can point to the same whole
    source file. Source IDs take precedence over display names and recorded
    content hashes are verified before publication. ``stats`` counts physical
    saved files, missing attachment references and ambiguous names.
    Missing or changed sources are reported without claiming a successful copy.
    """
    if log is None:

        def log(_msg: str) -> None:
            pass

    os.makedirs(dest_dir, exist_ok=True)
    saved: dict[str, str] = {}
    used_names: set[str] = set()
    attempted: set[str] = set()

    if evidence_df is None or "Attachment Name" not in evidence_df.columns:
        return saved

    rows = evidence_df.to_dict("records")
    paths: dict[str, set[str]] = defaultdict(set)
    unresolved_identities: set[str] = set()
    hashes: dict[str, set[str]] = defaultdict(set)
    source_invoices: dict[str, set[str]] = defaultdict(set)
    invoice_sources: dict[str, set[str]] = defaultdict(set)
    for row in rows:
        att = _text(row.get("Attachment Name"))
        source_id = _text(row.get("Source ID"))
        # A recorded identity is authoritative: a missing ID must never pick a
        # different occurrence merely because its display filename matches.
        src = source_paths.get(source_id) if source_id else source_paths.get(att)
        if att and source_id and not src:
            unresolved_identities.add(att)
        if att and src:
            src = os.path.realpath(src)
            paths[att].add(src)
            digest = _text(row.get("Content Hash"))
            if digest:
                hashes[src].add(digest)
            invoice = _text(row.get("Invoice #"))
            if invoice:
                source_invoices[src].add(invoice)
                invoice_sources[invoice].add(src)

    copied_sources: dict[str, str] = {}
    multi_att_invoices: set[str] = set()
    att_invoices: dict[str, set[str]] = {}
    if "Invoice #" in evidence_df.columns:
        for _, r in evidence_df.iterrows():
            inv_num = str(r.get("Invoice #", "")).strip()
            att = str(r.get("Attachment Name", "N/A"))
            if not inv_num or inv_num in ("N/A", "None", "nan"):
                continue
            if att in ("N/A", "", "None", "nan"):
                continue
            att_invoices.setdefault(att, set()).add(inv_num)
            if len(invoice_sources[inv_num]) > 1:
                multi_att_invoices.add(inv_num)

    for _, r in evidence_df.iterrows():
        att = str(r.get("Attachment Name", "N/A"))
        if att in ("N/A", "", "None", "nan"):
            continue
        if att in attempted:
            continue
        attempted.add(att)
        if att in unresolved_identities:
            log(f"evidence_files: missing source identity for {att!r}; no file copied")
            continue
        choices = paths.get(att, set())
        if len(choices) > 1:
            log(f"evidence_files: ambiguous source identity for {att!r}; no file copied")
            continue
        src = next(iter(choices), "")
        if not src or not os.path.isfile(src):
            log(f"evidence_files: missing source for {att!r}")
            continue
        if src in copied_sources:
            saved[att] = copied_sources[src]
            continue
        if len(hashes[src]) > 1:
            log(f"evidence_files: conflicting content hashes for {att!r}; no file copied")
            continue
        # Invoice slices still reference a whole parent PDF, never separate documents.
        display_name = _text(r.get("Original Filename")) or re.sub(r" #[0-9]+$", "", att)
        dest_base = sanitise_filename(display_name)
        inv_num = next(iter(source_invoices[src]), "") if len(source_invoices[src]) == 1 else ""
        if inv_num and inv_num not in multi_att_invoices:
            dest_base = sanitise_filename(inv_num + os.path.splitext(display_name)[1])
        stem, ext = os.path.splitext(dest_base)
        n = 2
        while dest_base.casefold() in used_names or os.path.exists(
            os.path.join(dest_dir, dest_base)
        ):
            dest_base = f"{stem}-{n}{ext}"
            n += 1
        used_names.add(dest_base.casefold())
        dest_path = os.path.join(dest_dir, dest_base)
        try:
            _copy_source(src, dest_path, next(iter(hashes[src]), ""))
        except (OSError, ValueError) as error:
            log(f"evidence_files: could not copy {att!r}: {error}")
            continue
        saved[att] = dest_path
        copied_sources[src] = dest_path

    ambiguous = {att for att, invs in att_invoices.items() if len(invs) > 1} | {
        att for att, choices in paths.items() if len(choices) > 1
    }
    for att in sorted(ambiguous):
        log(f"evidence_files: ambiguous attachment {att!r} referenced by multiple invoices")
    if stats is not None:
        missing_attrs = {
            str(r.get("Attachment Name", "N/A"))
            for _, r in evidence_df.iterrows()
            if str(r.get("Attachment Name", "N/A")) not in ("N/A", "", "None", "nan")
            and str(r.get("Attachment Name", "N/A")) not in saved
        }
        stats["saved"] = len(copied_sources)
        stats["missing"] = len(missing_attrs)
        stats["ambiguous"] = len(ambiguous)
    return saved


def build_bundle_manifest(
    evidence_df: pd.DataFrame, saved_files: dict[str, str], bundle_dir: str
) -> dict:
    """Publish portable filename, source identity, invoice and page mappings."""
    entries: dict[tuple[str, str], dict] = {}
    artifacts: dict[str, dict] = {}
    for row in evidence_df.to_dict("records"):
        att = _text(row.get("Attachment Name"))
        if not att:
            continue
        source_id = _text(row.get("Source ID"))
        saved = saved_files.get(att)
        if saved and saved not in artifacts:
            artifacts[saved] = artifact_record(saved)
        entry = entries.setdefault(
            (source_id, att),
            {
                "source_id": source_id or None,
                "attachment_name": att,
                "original_filename": _text(row.get("Original Filename")) or att,
                "parent_message": _text(row.get("Parent Message")) or None,
                "saved_path": os.path.relpath(saved, bundle_dir).replace(os.sep, "/")
                if saved
                else None,
                "sha256": artifacts[saved]["sha256"] if saved else None,
                "bytes": artifacts[saved]["bytes"] if saved else None,
                "status": "saved" if saved else "unavailable",
                "invoices": [],
                "page_ranges": [],
            },
        )
        invoice = _text(row.get("Invoice #"))
        if invoice and invoice not in entry["invoices"]:
            entry["invoices"].append(invoice)
        pages = [row.get("Page From"), row.get("Page To")]
        if (
            all(
                isinstance(page, (int, float))
                and not isinstance(page, bool)
                and 1 <= page <= 1_000_000
                and int(page) == page
                for page in pages
            )
            and pages[0] <= pages[1]
        ):
            page_range = [int(page) for page in pages]
            if page_range not in entry["page_ranges"]:
                entry["page_ranges"].append(page_range)
    manifest = {"bundle_version": 1, "files": list(entries.values())}
    write_json(manifest, os.path.join(bundle_dir, "evidence_manifest.json"))
    return manifest


def resolve_bundle_paths(state: dict, bundle_dir: str) -> tuple[dict[str, str], tuple[str, ...]]:
    """Resolve relocated originals by saved relative paths without changing analysis."""
    root = Path(bundle_dir).resolve()
    paths: dict[str, str] = {}
    notes: list[str] = []
    manifest = state.get("source_manifest", {})
    if not isinstance(manifest, dict):
        return {}, ("Stored evidence mapping is unavailable.",)
    aliases: dict[str, set[str]] = defaultdict(set)
    for source_id, metadata in manifest.items():
        if isinstance(metadata, dict):
            name = metadata.get("Attachment Name")
            if isinstance(name, str) and name:
                # Include unavailable occurrences too: their names do not become
                # unambiguous simply because only one copy can currently open.
                aliases[name].add(source_id)
    for source_id, metadata in manifest.items():
        if not isinstance(metadata, dict):
            notes.append(f"Stored evidence mapping is unavailable: {source_id}")
            continue
        relative = metadata.get("Saved Path")
        if not relative:
            continue
        if not isinstance(relative, str) or re.match(r"^(?:[A-Za-z]:|[/\\])", relative):
            notes.append(f"Stored evidence path is invalid: {source_id}")
            continue
        try:
            source = (root / relative.replace("\\", "/")).resolve()
            if not source.is_relative_to(root) or not source.is_file():
                notes.append(
                    f"Original evidence unavailable: {metadata.get('Original Filename', source_id)}"
                )
                continue
        except (OSError, ValueError, RuntimeError):
            notes.append(f"Original evidence unavailable: {source_id}")
            continue
        expected = metadata.get("Content Hash")
        try:
            if expected and artifact_record(str(source))["sha256"] != expected:
                notes.append(
                    f"Original evidence changed: {metadata.get('Original Filename', source_id)}"
                )
                continue
        except OSError as error:
            notes.append(f"Original evidence unreadable: {source.name}: {error}")
            continue
        paths[source_id] = str(source)
    for name, identities in aliases.items():
        if len(identities) > 1:
            notes.append(f"Stored evidence attachment name is ambiguous: {name}")
        elif name not in manifest:
            identity = next(iter(identities))
            if identity in paths:
                paths[name] = paths[identity]
    # Older compatibility results may have absolute mappings without a manifest.
    if not manifest:
        legacy_paths = state.get("saved_evidence_paths", {})
        if not isinstance(legacy_paths, dict):
            return {}, ("Stored evidence mapping is unavailable.",)
        for key, path in legacy_paths.items():
            try:
                if not isinstance(path, str) or not path:
                    raise ValueError("Missing legacy evidence path")
                legacy_source = Path(path.replace("\\", "/"))
                relative = not legacy_source.is_absolute()
                if relative:
                    legacy_source = root / legacy_source
                legacy_source = legacy_source.resolve()
                if (
                    relative and not legacy_source.is_relative_to(root)
                ) or not legacy_source.is_file():
                    raise ValueError("Unavailable legacy evidence path")
                paths[key] = str(legacy_source)
            except (OSError, ValueError, RuntimeError):
                notes.append(f"Original evidence unavailable: {key}")
    return paths, tuple(notes)


# ---------------------------------------------------------------------------
# Build the themed-section DOCX bundle index
# ---------------------------------------------------------------------------


@atomic_output()
def build_bundle_index(
    evidence_df: pd.DataFrame,
    saved_files: dict[str, str],
    output_path: str,
    account: str = "",
) -> None:
    """Generate the themed-section DOCX bundle index.

    The doc has one ``<h2>`` per section (A..G, in the order the spec
    defines) and one bullet per saved file. Rows are routed via
    :func:`_infer_section`. Unprefixed rows are tagged with the
    ``⚠  unprefixed`` marker on the bullet line.
    """
    import docx  # local import so unit tests can mock it out cheaply.

    doc = docx.Document()
    h = doc.add_heading("EDF Evidence Bundle Index", level=1)
    if account:
        h.add_run(f"  |  Account {account}")
    doc.add_paragraph(
        "Successfully copied source files are reproduced in the "
        "`evidence_files/` subfolder. This index lists them grouped by the "
        "section letter assigned during inference."
    )
    doc.add_paragraph(
        "Original filenames and invoice references are mapped to the linked saved files. "
        "Merged PDFs are preserved whole; evidence_manifest.json records invoice page ranges."
    )

    # Bucket rows by section letter.
    by_section: dict[str, list[tuple[str, str, bool]]] = defaultdict(list)
    seen_paths: set[str] = set()
    original_names: dict[str, set[str]] = defaultdict(set)
    invoices: dict[str, set[str]] = defaultdict(set)
    for row in evidence_df.to_dict("records") if evidence_df is not None else []:
        att = _text(row.get("Attachment Name"))
        if att in saved_files:
            path = saved_files[att]
            original_names[path].add(_text(row.get("Original Filename")) or att)
            invoice = _text(row.get("Invoice #"))
            if invoice:
                invoices[path].add(invoice)
    if evidence_df is not None and not evidence_df.empty:
        for _, r in evidence_df.iterrows():
            att = str(r.get("Attachment Name", "N/A"))
            if att in ("N/A", "", "None", "nan"):
                continue
            if att not in saved_files:
                continue
            saved_path = saved_files[att]
            if saved_path in seen_paths:
                continue
            seen_paths.add(saved_path)
            section_letter, is_unprefixed = _infer_section(att)
            date = _text(r.get("Date", ""))
            by_section[section_letter].append((att, date, is_unprefixed))

    # Emit one section heading + bullet list per letter present.
    for letter, label in SECTION_LABELS.items():
        rows = by_section.get(letter, [])
        if not rows:
            continue
        doc.add_heading(label, level=2)
        # Render one bullet per saved file.
        for att, date, is_unprefixed in rows:
            saved_path = saved_files.get(att, att)
            bullet_text = "Original: " + ", ".join(sorted(original_names[saved_path]))
            if invoices[saved_path]:
                bullet_text += "  |  Invoices: " + ", ".join(sorted(invoices[saved_path]))
            if date:
                bullet_text += f"  —  date {date}"
            if is_unprefixed:
                bullet_text += "  ⚠  unprefixed"
            paragraph = doc.add_paragraph(bullet_text + "  |  Saved: ", style="List Bullet")
            from docx.opc.constants import RELATIONSHIP_TYPE
            from docx.oxml import OxmlElement
            from docx.oxml.ns import qn

            target = os.path.relpath(saved_path, os.path.dirname(os.path.abspath(output_path)))
            relationship = paragraph.part.relate_to(
                quote(target.replace(os.sep, "/"), safe="/"),
                RELATIONSHIP_TYPE.HYPERLINK,
                is_external=True,
            )
            hyperlink = OxmlElement("w:hyperlink")
            hyperlink.set(qn("r:id"), relationship)
            run = OxmlElement("w:r")
            text = OxmlElement("w:t")
            text.text = os.path.basename(saved_path)
            run.append(text)
            hyperlink.append(run)
            paragraph._p.append(hyperlink)

    doc.save(output_path)
