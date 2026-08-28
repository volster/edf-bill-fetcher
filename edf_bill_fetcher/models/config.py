"""Typed configuration contract shared across all consumer layers.

The evidence pipeline passes one configuration dictionary from the GUI and
CLI entry points down to the collectors, Excel writer, and PDF/DOCX
reporters. Historically that dictionary was untyped (``dict``), so a
misspelled key such as ``use_acc_filt`` instead of ``use_acc_filter`` was
only caught at runtime -- and worse, the GUI persisted a *different* set of
names (``use_reading_class``, ``use_acc_filt``) than the consumers read
(``use_reading_classification``, ``use_acc_filter``), silently dropping the
GUI values unless the ``_run`` mapping kept them in sync.

``ConfigDict`` is the single documented source of truth for that
configuration surface. Every consumer takes ``ConfigDict`` and every key is
optional (``total=False``) because consumers read through
``config.get(key, default)`` and the GUI/CLI may legitimately omit keys.

Naming rules enforced by the type checker:
    - All keys use their canonical full name (e.g. ``use_acc_filter``).
    - Legacy short names from the old GUI persistence schema
      (``use_acc_filt``, ``use_reading_class``) are NOT valid keys.
    - ``models.ConfigDict`` (not ``dict``) is the parameter type for
      engine, exporter, and reporter entry points.
"""

from __future__ import annotations

from typing import TypedDict


class ConfigDict(TypedDict, total=False):
    """Typed contract for the engine/export/report configuration dictionary.

    Keys are optional so that partial configuration (``{}`` from the CLI
    with no ``--config`` file) type-checks; consumers supply defaults via
    ``.get()``. Keeping every key optional means the type checker still
    validates the *names* and *value types* of the keys that ARE present,
    which is the entire point: a misspelled key becomes a mypy error at the
    call site instead of a silent runtime no-op.
    """

    # --- Account / identity -------------------------------------------------
    acc_num: str
    report_account_ref: str

    # --- Amount thresholds --------------------------------------------------
    min_amount: float
    analysis_min: float

    # --- Filtering ----------------------------------------------------------
    filter_below: bool
    save_filtered: bool
    save_dups: bool
    use_domain_filter: bool
    domain_filter: str

    # --- Deduplication / amalgamation ---------------------------------------
    use_dedup: bool
    amalgamate_duplicates: bool

    # --- Extraction heuristics ----------------------------------------------
    use_anchors: bool
    use_large: bool
    use_reading_classification: bool
    use_pdf_fields: bool
    use_acc_filter: bool

    # --- SAP / reconciliation gates -----------------------------------------
    scan_sap_dumps: bool
    generate_reconciliation_sheet: bool

    # --- Report section selection (PDF/DOCX) --------------------------------
    report_sections: list[str]

    # --- Compensation estimator (Wave 6d) -----------------------------------
    as_of: str
    credit_hold_days: int
    credit_interest_rate: float

    # --- Evidence-file sidecar ----------------------------------------------
    save_evidence_files: bool


__all__ = ["ConfigDict"]
