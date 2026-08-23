"""EvidenceEngine — orchestrator for PDF/PST/text extraction.

Extracted from ``edf_collector.py`` as part of the modularization refactor
(Task 7 — Phase 6). The class plus its 7 tightly-coupled helper functions
live here; backward-compat re-exports remain in ``edf_collector.py`` until
Task 8 strips them.
"""

from __future__ import annotations

import hashlib
import os
import re
import sys
import tempfile
import threading

try:
    import pypff  # type: ignore[import-untyped]  # noqa: F401

    HAS_PYPFF = True
except ImportError:
    HAS_PYPFF = False
from collections.abc import Callable
from datetime import datetime
from typing import Any

import pdfplumber
from bs4 import BeautifulSoup
from openpyxl.styles import Side as _Side

# Canonical homes (per Tasks 1-6 organization)
from edf_bill_fetcher.helpers.date_utils import (
    parse_to_display_date,
)
from edf_bill_fetcher.helpers.domain_filter import matches_domain_filter
from edf_bill_fetcher.helpers.fallback_extractors import (
    fallback_amount,
    fallback_inv_num,
    fallback_period_from,
    fallback_period_to,
)
from edf_bill_fetcher.helpers.formatting import account_number_matches
from edf_bill_fetcher.helpers.pst_resources import (
    EMAIL_ADDR_RE,
    extract_sender_email,
    pst_attachment_filename,
)
from edf_bill_fetcher.io.adapters.eml import parse_eml_message
from edf_bill_fetcher.io.adapters.html import parse_htm_account_history
from edf_bill_fetcher.io.adapters.msg import HAS_EXTRACT_MSG, parse_msg_message
from edf_bill_fetcher.io.adapters.pdf import extract_admit_phrase, slice_pdf_pages
from edf_bill_fetcher.models.config import ConfigDict
from edf_bill_fetcher.models.records import BillingRecord
from edf_bill_fetcher.processors.detection import detect_pdf_format
from edf_bill_fetcher.processors.patterns import (
    _AMOUNT_PATTERN_NEW_BILL,
    _AMOUNT_PATTERN_ONGOING_BALANCE,
    _POUND_AMOUNT_FALLBACK_RE,
    AMOUNT_PATTERNS,
    PERIOD_RE,
    READING_PATTERNS,
    extract_sub_periods,
)
from edf_bill_fetcher.processors.sap_parsers import (
    detect_reconciliation_statement,
    detect_sap_dump,
    extract_new_credit_fields,
    extract_new_invoice_fields,
    extract_reconciliation_statement_rows,
    parse_sap_contract_history,
    parse_sap_financial_transactions,
    parse_sap_meter_read_history,
)

# --- Fallback extractors ---
# ``_fallback_inv_num`` / ``_fallback_period_from`` / ``_fallback_period_to``
# / ``_fallback_amount`` live in ``edf_bill_fetcher.helpers.fallback_extractors``
# (single source of truth); the underscore aliases preserve the names the
# engine's internal callers use.
_fallback_inv_num = fallback_inv_num
_fallback_period_from = fallback_period_from
_fallback_period_to = fallback_period_to
_fallback_amount = fallback_amount


# detect_sap_dump moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); see re-export block at top of file.

# _sap_to_iso_date helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.

# _parse_sap_csv helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.

# parse_sap_contract_history moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); see re-export block at top of file.

# parse_sap_meter_read_history moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); see re-export block at top of file.

# parse_sap_financial_transactions moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); see re-export block at top of file.

_SAP_DEBT_MGMT_FLAG_VALUE = "Installment Plan Item"
_SAP_MIN_CLUSTER_SIZE = 4
_SAP_MATCH_DAY_BANDS = ((0, 50), (3, 25), (14, 5))
_SAP_MATCH_AMOUNT_BANDS = ((0.05, 40), (0.25, 20), (0.50, 5))
_SAP_CONFIDENCE_BANDS = (("High", 75), ("Medium", 40), ("Low", 10))

_OLD_PDF_DATE_RE = re.compile(
    r"(?:Bill date|Date issued):\s*[\",]*\s*(\d{1,2}\s+\w+\s+\d{4})",
    re.IGNORECASE,
)
_OLD_PDF_KWH_RE = re.compile(r"([\d,]+)\s*kWh", re.IGNORECASE)
_OLD_PDF_STANDING_RE = re.compile(r"(\d+\.\d{2})p\s*per day", re.IGNORECASE)
_OLD_PDF_INV_RE = re.compile(r"Invoice number[\s:,\"\'\n]*([A-Z0-9\-]+)", re.IGNORECASE)
_OLD_PDF_PERIOD_CHARGE_RE = re.compile(
    r"total charges for this (?:period|bill|invoice)\s+£\s?([\d,]+(?:\.\d{2})?)",
    re.IGNORECASE,
)

_BILL_MARKERS_RE = re.compile(
    r"(?:bill date|date issued|invoice number|total charges|your charges)"
)
_ACCOUNT_BALANCE_LANG_RE = re.compile(
    r"(?:account balance|running balance|balance brought forward)"
)
_BILL_INDICATORS_RE = re.compile(r"(?:kwh|standing charge|tariff)")


# _parse_amount_for_event helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.


# --- PST/OST helpers ---
# ``_pst_attachment_filename`` / ``_extract_sender_email`` live in
# ``edf_bill_fetcher.helpers.pst_resources`` (single source of truth); the
# underscore aliases preserve the names the engine's internal callers use.
_pst_attachment_filename = pst_attachment_filename
_extract_sender_email = extract_sender_email


_matches_domain_filter = matches_domain_filter


# --- EvidenceEngine ---
class EvidenceEngine:
    """Orchestrate extraction of EDF billing records from PDF, HTM, and PST sources."""

    def __init__(
        self,
        config: ConfigDict,
        update_ui_cb: Callable[[str], None],
        progress_cb: Callable[[int, int, str], None] | None = None,
        cancel_event: threading.Event | None = None,
    ):
        """Initialize the engine with config, a UI callback, and cancellation hooks."""
        self.config = config
        self.records: list[dict[str, Any]] = []
        self.filtered_records: list[dict[str, Any]] = []
        self.update_ui = update_ui_cb
        self.update_progress = progress_cb
        self.cancel_event = cancel_event or threading.Event()
        self.pdf_count = 0
        self.email_count = 0
        self.error_log: list[str] = []
        self.seen_pdf_hashes: set[str] = set()
        self.lock = threading.Lock()
        # Stream P1: SAP CSV-in-PDF data dumps are detected in
        # ``process_pdf_file`` and routed to three row accumulators
        # (contract / meter_read / financial). ``export_to_excel``
        # reads these through ``sap_rows={...}`` to render the
        # dedicated SAP sheets + the cross-source Reconciliation sheet.
        self.sap_contract_rows: list[dict] = []
        self.sap_meter_rows: list[dict] = []
        self.sap_financial_rows: list[dict] = []
        # Stream P5: reverse-lookup from attachment name to absolute source
        # path, populated by process_pdf_file so save_evidence_files can copy
        # the originals into evidence_files/. Pre-fix this attribute was never
        # set — getattr(engine, "source_paths", {}) always returned {} and
        # every attachment was skipped with "missing source for X". Spec §3.9.
        self.source_paths: dict[str, str] = {}

    # ------------------------------------------------------------------
    # Pickle support — Phase 1.4
    # ------------------------------------------------------------------
    def __getstate__(self) -> dict:
        """Return a picklable snapshot of the engine data.

        ``EvidenceEngine`` carries three non-picklable runtime
        primitives — ``threading.Lock``, ``threading.Event``, and the
        two callbacks — which can't survive a naive ``pickle.dump`` of
        the instance (``TypeError: cannot pickle '_thread.lock' object``).
        We round-trip the *data* the engine holds, and rebuild the
        threading primitives fresh in ``__setstate__``.

        This means a loaded engine is fully usable again — just with
        fresh ``Lock``/``Event`` instances and no cancellation state
        from the persisting session — which is the right semantic for
        a CLI report-on-engine-data flow that resumes a saved snapshot.

        Concretely we strip:
          * ``self.lock``             (``threading.Lock`` — not picklable)
          * ``self.cancel_event``     (``threading.Event`` — not picklable)
          * ``self.update_ui``        (a GUI callback; serialising
                                      Tkinter closures would leak the GUI
                                      context across the CLI↔GUI boundary)
          * ``self.update_progress``  (same reason)

        and rebuild them in ``__setstate__``.
        """
        return {
            "config": self.config,
            "records": self.records,
            "filtered_records": self.filtered_records,
            "pdf_count": self.pdf_count,
            "email_count": self.email_count,
            "error_log": self.error_log,
            "seen_pdf_hashes": self.seen_pdf_hashes,
            "sap_contract_rows": self.sap_contract_rows,
            "sap_meter_rows": self.sap_meter_rows,
            "sap_financial_rows": self.sap_financial_rows,
            "source_paths": self.source_paths,
        }

    def __setstate__(self, state: dict) -> None:
        """Restore a pickled snapshot — rebuild non-picklable fields fresh.

        See ``__getstate__`` for why each of these is set this way.
        """
        self.config = state["config"]
        self.records = state["records"]
        self.filtered_records = state["filtered_records"]
        self.pdf_count = state["pdf_count"]
        self.email_count = state["email_count"]
        self.error_log = state["error_log"]
        self.seen_pdf_hashes = state["seen_pdf_hashes"]
        # Newer data fields (added after the original pickle contract):
        # restore with backwards-compatible defaults so pre-existing
        # pickles (which lack these keys) still load.
        self.sap_contract_rows = state.get("sap_contract_rows", [])
        self.sap_meter_rows = state.get("sap_meter_rows", [])
        self.sap_financial_rows = state.get("sap_financial_rows", [])
        self.source_paths = state.get("source_paths", {})
        # Rebuild runtime primitives fresh — the persisted snapshot
        # does not carry cancel state forward.
        self.cancel_event = threading.Event()
        self.lock = threading.Lock()
        # GUI callbacks don't survive a CLI↔CLI round-trip; a GUI
        # consumer can install its own after loading the snapshot via
        # ``engine.update_ui = my_gui_callback``.
        self.update_ui = lambda *_a, **_kw: None
        self.update_progress = lambda *_a, **_kw: None

    def is_cancelled(self):
        """Return True if the cancel event has been set by the GUI."""
        return self.cancel_event.is_set()

    def log_error(self, context, err):
        """Append a timestamped error entry to the engine's error log."""
        self.error_log.append(f"[{datetime.now().strftime('%H:%M:%S')}] {context} — {err}")

    def find_billing_period(self, text):
        """Extract the billing period (Period From, Period To) from text."""
        m = PERIOD_RE.search(text)
        if m:
            return (
                parse_to_display_date(m.group(1).strip()),
                parse_to_display_date(m.group(2).strip()),
            )
        return "N/A", "N/A"

    def _add_record(self, rec):
        """Thread-safe record append after optional magnitude-based filter check.

        Filter compares ``abs(amount)`` to ``min_amount`` so high-magnitude
        refunds (e.g. ``-£1000``) are KEPT in the main records — only
        records whose absolute amount is below the threshold are shelved
        to ``filtered_records``.

        The amount field accepts both numeric and string values: strings
        are coerced to float where possible, and an uncoercible sentinel
        (e.g. ``"N/A"``) becomes ``None`` so the filter check is skipped
        and the record is kept rather than crashing the run.
        """
        raw_amt = rec.get("Amount (£)", 0) or 0
        try:
            amt = float(raw_amt)
        except (TypeError, ValueError):
            amt = None
        if (
            amt is not None
            and self.config.get("filter_below", True)
            and abs(amt) < self.config.get("min_amount", 50.0)
        ):
            with self.lock:
                self.filtered_records.append(
                    {
                        "Source": rec.get("Source", ""),
                        "Date": rec.get("Date", ""),
                        "Amount (£)": amt,
                        "Details": rec.get("Details", "")[:60],
                        "Logic Used": rec.get("Logic Used", ""),
                        "Reason": f"Amount magnitude below £{self.config.get('min_amount', 50.0):,.2f} threshold",
                    }
                )
            return
        with self.lock:
            self.records.append(rec)

    # ------------------------------------------------------------------
    # New-format PDF processing
    # ------------------------------------------------------------------

    def _process_new_invoice(
        self, text, source_label, detail_label, fallback_date, sender="", attachment_name=""
    ):
        fields = extract_new_invoice_fields(text)
        if "amount" not in fields:
            return False  # didn't match

        sub_periods = extract_sub_periods(text)

        # Account filter
        if self.config.get("use_acc_filter"):
            acc = self.config.get("acc_num", "")
            if acc and not account_number_matches(acc, text):
                return False

        r_type = "Unknown"
        for label, pat in READING_PATTERNS.items():
            if pat.search(text):
                r_type = label
                break

        # Multi-regex fallback chain (Stream P3): when the canonical
        # extractors miss, try the cover-block / loose-token fallbacks so
        # the analyser tabs see fewer N/A entries. Each fallback records
        # the regex that produced the value into ``_regex_trace`` so the
        # Source Excerpt column (Stream P3 / Task 6) can show a parse trace.
        regex_trace: list[str] = []
        inv_num = fields.get("inv_num")
        if not inv_num or inv_num == "N/A":
            val, label = _fallback_inv_num(text)
            if val:
                fields["inv_num"] = val
                regex_trace.append(f"inv_num via {label}")
        else:
            regex_trace.append("inv_num via _INV_NUMBER_RE")

        period_from = fields.get("period_from")
        if not period_from or period_from == "N/A":
            val, label = _fallback_period_from(text)
            if val:
                fields["period_from"] = val
                regex_trace.append(f"period_from via {label}")
        else:
            regex_trace.append("period_from via _BILLING_PERIOD_RE")

        period_to = fields.get("period_to")
        if not period_to or period_to == "N/A":
            val, label = _fallback_period_to(text)
            if val:
                fields["period_to"] = val
                regex_trace.append(f"period_to via {label}")
        else:
            regex_trace.append("period_to via _BILLING_PERIOD_RE")

        if "period_charge" not in fields:
            amt_val, label = _fallback_amount(text)
            if amt_val is not None:
                fields["period_charge"] = amt_val
                regex_trace.append(f"period_charge via {label}")

        # Classify entry type: New Bill if it has period charges, else Ongoing Balance
        entry_type = (
            "New Bill"
            if fields.get("period_charge") or fields.get("period_from")
            else "Ongoing Balance"
        )

        self._add_record(
            BillingRecord(
                source=source_label,
                sender=sender,
                date=fields.get("date", fallback_date),
                period_from=fields.get("period_from", "N/A"),
                period_to=fields.get("period_to", "N/A"),
                invoice_num=fields.get("inv_num", "N/A"),
                amount=fields["amount"],
                period_charge=fields.get("period_charge", "N/A"),
                entry_type=entry_type,
                reading=r_type,
                units_kwh=fields.get("units_used", "N/A"),
                standing_charge=fields.get("standing_charge", "N/A"),
                tariff=fields.get("tariff", "N/A"),
                attachment_name=attachment_name or "N/A",
                details=(detail_label or "New invoice")[:60],
                logic_used="New Invoice Format",
                source_pdf_text=text[:4000],
                regex_trace="; ".join(regex_trace) if regex_trace else "",
                cancel_rebill_admitted=bool(extract_admit_phrase(text)),
                sub_periods=sub_periods,
            ).to_dict()
        )
        return True

    def _process_new_credit(
        self, text, source_label, detail_label, fallback_date, sender="", attachment_name=""
    ):
        fields = extract_new_credit_fields(text)
        if "amount" not in fields:
            return False

        if self.config.get("use_acc_filter"):
            acc = self.config.get("acc_num", "")
            if acc and not account_number_matches(acc, text):
                return False

        # Multi-regex fallback chain (Stream P3): borrow the same patterns
        # used in ``_process_new_invoice`` to recover fields the canonical
        # extractor missed.
        regex_trace: list[str] = []
        inv_num = fields.get("inv_num")
        if not inv_num or inv_num == "N/A":
            val, label = _fallback_inv_num(text)
            if val:
                fields["inv_num"] = val
                regex_trace.append(f"inv_num via {label}")
        else:
            regex_trace.append("inv_num via _CREDIT_NUMBER_RE")

        pf_val, pf_label = _fallback_period_from(text)
        pt_val, pt_label = _fallback_period_to(text)
        period_from = pf_val or "N/A"
        period_to = pt_val or "N/A"
        if pf_val:
            regex_trace.append(f"period_from via {pf_label}")
        if pt_val:
            regex_trace.append(f"period_to via {pt_label}")

        self._add_record(
            BillingRecord(
                source=source_label,
                sender=sender,
                date=fields.get("date", fallback_date),
                period_from=period_from,
                period_to=period_to,
                invoice_num=fields.get("inv_num", "N/A"),
                amount=fields["amount"],
                period_charge="N/A",
                entry_type="Credit",
                reading="N/A",
                units_kwh="N/A",
                standing_charge="N/A",
                tariff="N/A",
                attachment_name=attachment_name or "N/A",
                details=(detail_label or "Credit note")[:60],
                logic_used="New Credit Note Format",
                source_pdf_text=text[:4000],
                regex_trace="; ".join(regex_trace) if regex_trace else "",
                cancel_rebill_admitted=bool(extract_admit_phrase(text)),
            ).to_dict()
        )
        return True

    # ------------------------------------------------------------------
    # Generic text processing (old format + email bodies)
    # ------------------------------------------------------------------

    def process_text(self, text, source_type, detail, fallback_date, sender="", attachment_name=""):
        """Extract a record from a generic text body (old-format PDF or email)."""
        if not text:
            return

        clean_text = re.sub(r"\s+", " ", text)

        # Account filter
        if self.config.get("use_acc_filter"):
            acc = self.config.get("acc_num", "")
            if acc and not account_number_matches(acc, clean_text):
                return

        found_amt, strategy = None, ""
        matched_pattern_name: str | None = None

        if self.config.get("use_anchors", True):
            for name, p in AMOUNT_PATTERNS:
                # Patterns are pre-compiled at module load with
                # `re.IGNORECASE` baked in, so search() takes no flags.
                m = p.search(clean_text)
                if m:
                    try:
                        found_amt = float(m.group(1).replace(",", ""))
                        strategy = "Smart Context"
                        matched_pattern_name = name
                        break
                    except Exception:
                        continue

        if not found_amt and self.config.get("use_large", True):
            matches = _POUND_AMOUNT_FALLBACK_RE.findall(clean_text)
            if matches:
                floats = [float(x.replace(",", "")) for x in matches]
                highs = [x for x in floats if x >= self.config.get("min_amount", 50.0)]
                if highs:
                    found_amt = max(highs)
                    strategy = "Large Amount Fallback"

        if not found_amt:
            return

        # Date extraction
        date_to_use = fallback_date
        if "PDF" in source_type or "old" in source_type.lower():
            date_m = _OLD_PDF_DATE_RE.search(clean_text)
            if date_m:
                date_to_use = parse_to_display_date(date_m.group(1))

        r_type = "Unknown"
        if self.config.get("use_reading_classification", True):
            for label, pat in READING_PATTERNS.items():
                if pat.search(clean_text):
                    r_type = label
                    break

        units_used = standing_charge = inv_num = "N/A"
        if self.config.get("use_pdf_fields", True):
            u_m = _OLD_PDF_KWH_RE.search(clean_text)
            sc_m = _OLD_PDF_STANDING_RE.search(clean_text)
            in_m = _OLD_PDF_INV_RE.search(clean_text)
            if u_m:
                units_used = u_m.group(1)
            if sc_m:
                standing_charge = sc_m.group(1)
            if in_m:
                inv_num = in_m.group(1)

        period_from, period_to = self.find_billing_period(clean_text)

        # Attempt to extract period charge separately from cumulative balance
        period_charge: str | float = "N/A"
        pc_m = _OLD_PDF_PERIOD_CHARGE_RE.search(clean_text)
        if pc_m:
            try:
                period_charge = float(pc_m.group(1).replace(",", ""))
            except (ValueError, AttributeError):
                pass

        # Mine per-sub-period rows ("About your charges") from old-format
        # cancel-and-rebill invoices too.  The T-series bills are OLD-format
        # and arrive via this generic path (never ``_process_new_invoice``),
        # so without this the ``Sub Periods`` column stays empty and the
        # Option C sub-period unlawful-charge computation never activates.
        sub_periods = extract_sub_periods(clean_text)

        # Classify Entry Type based on content
        entry_type = self._classify_entry_type(
            clean_text, matched_pattern_name, period_from, period_to, strategy
        )

        self._add_record(
            BillingRecord(
                source=source_type,
                sender=sender,
                date=date_to_use,
                period_from=period_from,
                period_to=period_to,
                invoice_num=inv_num,
                amount=found_amt,
                period_charge=period_charge,
                entry_type=entry_type,
                reading=r_type,
                units_kwh=units_used,
                standing_charge=standing_charge,
                tariff="N/A",
                attachment_name=attachment_name or "N/A",
                details=detail[:60],
                logic_used=strategy,
                source_pdf_text=clean_text[:4000],
                regex_trace="",
                cancel_rebill_admitted=bool(extract_admit_phrase(clean_text)),
                sub_periods=sub_periods,
            ).to_dict()
        )

    def _classify_entry_type(
        self,
        text: str,
        pattern_name: str | None,
        period_from: str,
        period_to: str,
        strategy: str,
    ) -> str:
        """Classify a record as New Bill, Ongoing Balance, or Other based on content.

        Args:
            text: the cleaned bill body text.
            pattern_name: the name of the regex from
                :data:`AMOUNT_PATTERNS` that matched, or ``None`` if no
                anchored match was found.
            period_from: ``"N/A"`` or a parsed date string for the billing period start.
            period_to: ``"N/A"`` or a parsed date string for the billing period end.
            strategy: either ``"Smart Context"`` (anchored pattern
                matched) or ``"Large Amount Fallback"`` (anchored missed,
                number extracted by fallback).

        The classifier explicitly maps ``pattern_name`` to ``New Bill`` or
        ``Ongoing Balance`` via
        :data:`_AMOUNT_PATTERN_NEW_BILL` /
        :data:`_AMOUNT_PATTERN_ONGOING_BALANCE`. Unknown / unset names
        fall through to heuristic checks against the bill body text.

        """
        text_lower = text.lower()

        # If it has billing period dates AND charges/invoice details → New Bill
        has_period = period_from != "N/A" and period_to != "N/A"
        has_bill_markers = bool(_BILL_MARKERS_RE.search(text_lower))

        if has_period and has_bill_markers:
            return "New Bill"

        # Pattern-name driven classification. The integer-index lookup
        # used previously was brittle: reordering or inserting a pattern
        # silently changed classification. Names are stable.
        if pattern_name is not None:
            if pattern_name in _AMOUNT_PATTERN_NEW_BILL:
                return "New Bill"
            if pattern_name in _AMOUNT_PATTERN_ONGOING_BALANCE:
                return "Ongoing Balance"

        # If matched via "balance" pattern or has "account balance" language → Ongoing Balance
        if _ACCOUNT_BALANCE_LANG_RE.search(text_lower):
            return "Ongoing Balance"

        # If matched via total/amount to pay with period info → New Bill
        if has_period:
            return "New Bill"

        # Fallback strategy check
        if strategy == "Large Amount Fallback":
            return "Other"

        # Default: if it looks like a bill (has kWh, standing charge) → New Bill
        if _BILL_INDICATORS_RE.search(text_lower):
            return "New Bill"

        return "Ongoing Balance"

    # ------------------------------------------------------------------
    # PDF file processing — detects format automatically
    # ------------------------------------------------------------------

    def process_pdf_file(
        self, path, source_label, detail_label, fallback_date, sender="", attachment_name=""
    ):
        """Read a PDF file, detect its format, and extract any records it contains."""
        if self.is_cancelled():
            return
        try:
            import io

            with open(path, "rb") as fh:
                raw = fh.read()
            pdf_hash = hashlib.sha256(raw).hexdigest()
            with self.lock:
                if pdf_hash in self.seen_pdf_hashes:
                    return
                self.seen_pdf_hashes.add(pdf_hash)

            with pdfplumber.open(io.BytesIO(raw)) as pdf:
                # Handle empty or corrupt PDFs gracefully
                if not pdf.pages:
                    self.log_error(f"PDF: {detail_label}", "PDF has no pages")
                    return
                pdf_text_parts: list[str] = []
                for p in pdf.pages:
                    try:
                        page_text = p.extract_text()
                        if page_text:
                            pdf_text_parts.append(page_text)
                    except (
                        pdfplumber.utils.exceptions.PdfminerException,
                        ValueError,
                        TypeError,
                    ) as page_err:
                        # Narrowly catch PDF-syntax / text-coercion errors so
                        # a single bad page does not skip the whole file.
                        # ``BaseException`` (e.g. ``KeyboardInterrupt``) and
                        # unexpected runtime errors propagate so the caller
                        # can still cancel or surface real bugs.
                        self.log_error(
                            f"PDF page {detail_label}", f"Page extraction failed: {page_err}"
                        )
            del raw

            # Use original filename as attachment_name if not already set
            if not attachment_name:
                attachment_name = detail_label or ""

            # Stream P5: record the absolute source path so save_evidence_files
            # can copy the original into evidence_files/. Multi-slice merged
            # PDFs all point at the same path under per-slice attachment names
            # — the contract is "open the parent PDF", true regardless of
            # which slice a reviewer clicked. Spec §3.9 (issue 8b root cause).
            self.source_paths[attachment_name] = path

            # Multi-invoice PDF slicer. Merged PDFs (e.g. evidence-2026
            # ``D2 - T-series invoices (Sep 2023 - May 2024, merged).pdf``)
            # contain multiple invoices end-to-end. ``slice_pdf_pages``
            # partitions per-page text on ``Invoice number:`` or
            # ``Page 1 of N`` boundaries. Single-invoice PDFs return a
            # one-chunk list -- semantically identical to the legacy
            # whole-document concat.
            slices = slice_pdf_pages(pdf_text_parts)
            multi = len(slices) > 1

            for i, slice_pages in enumerate(slices, start=1):
                slice_text = " ".join(slice_pages)
                if multi:
                    slice_detail = f"{detail_label} #{i}"
                    slice_attachment = f"{attachment_name} #{i}"
                else:
                    slice_detail = detail_label
                    slice_attachment = attachment_name

                try:
                    # Stream P1: SAP CSV-in-PDF data dumps (Contract /
                    # Meter-Read / Financial-Transactions). Detected
                    # via the header-row marker; routed to dedicated
                    # parsers and stored on the engine's SAP-row
                    # accumulators (not engine.records -- they get
                    # their own dedicated Excel sheets via
                    # ``export_to_excel(..., sap_rows=...)``).
                    sap_kind = detect_sap_dump(slice_text)
                    if sap_kind is not None:
                        source_file = slice_attachment or slice_detail or ""
                        if sap_kind == "contract":
                            self.sap_contract_rows.extend(
                                parse_sap_contract_history(
                                    slice_text,
                                    source_file=source_file,
                                    report=lambda msg: self.log_error("SAP dump", msg),
                                )
                            )
                        elif sap_kind == "meter_read":
                            self.sap_meter_rows.extend(
                                parse_sap_meter_read_history(
                                    slice_text,
                                    source_file=source_file,
                                    report=lambda msg: self.log_error("SAP dump", msg),
                                )
                            )
                        elif sap_kind == "financial":
                            self.sap_financial_rows.extend(
                                parse_sap_financial_transactions(
                                    slice_text,
                                    source_file=source_file,
                                    report=lambda msg: self.log_error("SAP dump", msg),
                                )
                            )
                        continue
                    # Reconciliation statement PDFs (e.g. EDF's
                    # consolidated "Bill reference: … / Account number:
                    # A-… / Balance on your last bill" statements) carry
                    # many individual charge/reversal/payment rows under
                    # a single statement header. Detect them early and
                    # bypass the regular invoice-format dispatch so they
                    # emit one record per underlying row.
                    if detect_reconciliation_statement(slice_text):
                        rows = extract_reconciliation_statement_rows(slice_text, slice_attachment)
                        for row in rows:
                            self._add_record(row)
                        continue
                    fmt = detect_pdf_format(slice_text)
                    if fmt == "new_invoice":
                        self._process_new_invoice(
                            slice_text,
                            source_label,
                            slice_detail,
                            fallback_date,
                            sender=sender,
                            attachment_name=slice_attachment,
                        )
                    elif fmt == "new_credit":
                        self._process_new_credit(
                            slice_text,
                            source_label,
                            slice_detail,
                            fallback_date,
                            sender=sender,
                            attachment_name=slice_attachment,
                        )
                    else:
                        self.process_text(
                            slice_text,
                            source_label,
                            slice_detail,
                            fallback_date,
                            sender=sender,
                            attachment_name=slice_attachment,
                        )
                except Exception as slice_err:
                    # One bad slice must not lose the rest of the
                    # file's invoices. Swallow + log so the user still
                    # gets rows for any invoices that did parse.
                    self.log_error(f"PDF slice {i} {slice_detail}", str(slice_err))
                    continue

        except Exception as e:
            self.log_error(f"PDF: {detail_label}", str(e))

    # ------------------------------------------------------------------
    # HTM account history
    # ------------------------------------------------------------------

    def process_htm_file(self, path):
        """Read an HTM account-history export and extract every record it contains."""
        try:
            # Read with strict UTF-8 first — evidence data must not be
            # silently corrupted by mojibake replacement.  Fall back to
            # "replace" only if strict fails, and log a warning so the
            # user knows data may be imperfect.
            try:
                with open(path, encoding="utf-8", errors="strict") as f:
                    content = f.read()
            except UnicodeDecodeError:
                self.log_error(f"HTM: {path}", "UTF-8 decode error — some characters replaced")
                with open(path, encoding="utf-8", errors="replace") as f:
                    content = f.read()
            soup = BeautifulSoup(content, "html.parser")
            text = soup.get_text(separator=" ", strip=True)
            recs = parse_htm_account_history(text)
            for rec in recs:
                self._add_record(rec)
            self.update_ui(f"HTM: extracted {len(recs)} account history entries")
        except Exception as e:
            self.log_error(f"HTM: {path}", str(e))
            # Pre-fix the bare except swallowed failures silently; surface
            # them via update_ui, falling back to stderr if that's unusable.
            try:
                self.update_ui(f"Warning: failed to process HTM file {path}: {e}")
            except Exception:
                sys.stderr.write(f"Warning: failed to process HTM file {path}: {e}\n")

    def process_pst_file(self, path):
        """Open a PST file at ``path`` and crawl its root folder.

        Wrapper around :meth:`crawl_pst` so the public per-file API
        is symmetric with :meth:`process_pdf_file` and
        :meth:`process_htm_file`. Returns nothing; outcomes are
        surfaced through ``update_ui`` / ``error_log``.
        """
        if not HAS_PYPFF:
            self.log_error(
                "PST",
                f"pypff not installed — cannot open PST file {path}",
            )
            return
        try:
            pst = pypff.file()
            pst.open(path)
            try:
                root = pst.get_root_folder()
                self.crawl_pst(root)
            finally:
                try:
                    pst.close()
                except Exception:
                    pass
        except Exception as e:
            self.log_error(f"PST: {path}", str(e))

    # `process_ost_file` is the same code path: ``libpff-python`` accepts
    # both PST and OST archives. Exposed as an explicit alias so
    # callers picking from the per-file API do not have to know that.
    def process_ost_file(self, path):
        """Process an OST archive using the same code path as PST."""
        self.process_pst_file(path)

    # ------------------------------------------------------------------
    # PST / OST crawl
    # ------------------------------------------------------------------

    def crawl_pst(self, folder):
        """Recursively walk a PST/OST folder and process every EDF email found."""
        if not HAS_PYPFF:
            self.log_error("PST", "pypff not installed — skipping PST processing")
            return
        if self.is_cancelled():
            return

        msg_total = folder.get_number_of_sub_messages()
        for i in range(msg_total):
            if self.is_cancelled():
                return
            try:
                msg = folder.get_sub_message(i)
                subj = str(msg.get_subject() or "")
                d_time = msg.get_delivery_time()
                date_str = parse_to_display_date(d_time.strftime("%Y-%m-%d")) if d_time else "N/A"

                if self.update_progress and i % 100 == 0:
                    self.update_progress(
                        i + 1, msg_total, f"Scanning PST/OST folder: {i + 1}/{msg_total}"
                    )

                # Extract sender email for domain filtering and spreadsheet
                sender_email = _extract_sender_email(msg)

                # Determine if this email should be processed
                use_domain = self.config.get("use_domain_filter", False)
                domain_str = self.config.get("domain_filter", "")
                should_process = False
                if use_domain and domain_str:
                    if _matches_domain_filter(sender_email, domain_str):
                        should_process = True
                else:
                    if any(
                        k in subj.upper()
                        for k in ["EDF", "BILL", "STATEMENT", "ACCOUNT", "INVOICE"]
                    ):
                        should_process = True

                if should_process:
                    with self.lock:
                        self.email_count += 1
                    html = msg.get_html_body()
                    plain = msg.get_plain_text_body()

                    if html:
                        body_text = BeautifulSoup(html, "html.parser").get_text(separator=" ")
                        self.process_text(
                            body_text, "Email Body", subj, date_str, sender=sender_email
                        )
                    elif plain:
                        self.process_text(
                            plain.decode("utf-8", errors="ignore"),
                            "Email Body",
                            subj,
                            date_str,
                            sender=sender_email,
                        )
                    else:
                        rtf_body = None
                        try:
                            rtf_body = msg.get_rtf_body()
                        except Exception:
                            pass
                        if rtf_body:
                            try:
                                rtf_str = rtf_body.decode("utf-8", errors="replace")
                                rtf_text = re.sub(r"\\[a-z]+[-\d]*\s?", " ", rtf_str)
                                rtf_text = re.sub(r"[{}\\]", " ", rtf_text)
                                self.process_text(
                                    rtf_text,
                                    "Email Body (RTF)",
                                    subj,
                                    date_str,
                                    sender=sender_email,
                                )
                            except Exception as e:
                                self.log_error(f"Email: {subj}", f"RTF decode: {e}")
                        else:
                            self.log_error(f"Email: {subj} ({date_str})", "No readable body")

                    for a_idx in range(msg.get_number_of_attachments()):
                        if self.is_cancelled():
                            return
                        try:
                            att = msg.get_attachment(a_idx)
                            size = att.get_size()
                            if size > 4:
                                buf = att.read_buffer(size)
                                if buf and buf.startswith(b"%PDF"):
                                    with self.lock:
                                        self.pdf_count += 1
                                    att_name = _pst_attachment_filename(att)
                                    if not att_name:
                                        att_name = f"Attachment_{self.pdf_count}.pdf"
                                    with tempfile.NamedTemporaryFile(
                                        delete=False, suffix=".pdf"
                                    ) as tmp:
                                        tmp.write(buf)
                                        tmp_path = tmp.name
                                    try:
                                        self.process_pdf_file(
                                            tmp_path,
                                            "PST PDF Attachment",
                                            att_name,
                                            date_str,
                                            sender=sender_email,
                                            attachment_name=att_name,
                                        )
                                    finally:
                                        if os.path.exists(tmp_path):
                                            os.remove(tmp_path)
                        except Exception as e:
                            self.log_error(f'Attachment in "{subj}"', str(e))

            except Exception as e:
                self.log_error(f"PST message index {i}", str(e))

        self.update_ui(f"Scanned {self.email_count} emails, {self.pdf_count} attached PDFs…")
        for j in range(folder.get_number_of_sub_folders()):
            if self.is_cancelled():
                return
            self.crawl_pst(folder.get_sub_folder(j))

    # ------------------------------------------------------------------
    # Local PDF folder
    # ------------------------------------------------------------------

    def crawl_local_pdfs(self, path):
        """Recursively walk a local folder and process every PDF bill found."""
        if not path or not os.path.exists(path):
            return
        # Recursive walk: PDF bills are commonly organised into
        # sub-folders by year or account reference (e.g.
        # ``pdfs/2023/2023-01.pdf``).  The legacy implementation
        # only scanned the top-level directory and silently
        # dropped any bills in nested folders — a real EDF
        # dispute case with year-organised PDFs would have
        # silently undercounted, so this matters for ombudsman
        # submissions where a missing bill undoes the entire
        # argument.
        pdf_files: list[tuple[str, str]] = []
        for root, _dirs, files in os.walk(path):
            for f in files:
                if f.lower().endswith(".pdf"):
                    pdf_files.append((root, f))
        # Sort by relative path so the progress narrative is
        # deterministic across runs (otherwise os.walk's
        # filesystem-order output varies by platform).
        pdf_files.sort(
            key=lambda pair: os.path.relpath(os.path.join(pair[0], pair[1]), path).lower()
        )
        total = len(pdf_files)

        def _process_one(i_file):
            idx, (root, fname) = i_file
            if self.is_cancelled():
                return
            file_path = os.path.join(root, fname)
            fallback_date = parse_to_display_date(
                datetime.fromtimestamp(os.path.getmtime(file_path)).strftime("%Y-%m-%d")
            )
            with self.lock:
                self.pdf_count += 1
            self.process_pdf_file(
                file_path, "Local PDF Folder", fname, fallback_date, attachment_name=fname
            )
            if self.update_progress:
                relative = os.path.relpath(file_path, path)
                self.update_progress(idx, total, f"Scanning local PDFs: {idx}/{total} ({relative})")

        # Sequential pass.  The ``_process_one`` closure comment
        # above used to imply a thread-pool dispatch that's no
        # longer present (see also ``EvidenceEngine.lock`` which
        # is in fact exercised by ``process_pdf_file``'s own
        # write paths).  Keeping the indirection for now so
        # transition to ``ThreadPoolExecutor`` later stays a
        # one-line change.
        for item in enumerate(pdf_files, start=1):
            _process_one(item)

        self.update_ui(f"PDF folder: {self.pdf_count} PDFs processed")

    # ------------------------------------------------------------------
    # Local EML / MSG folders
    # ------------------------------------------------------------------

    def _process_message_folder(self, path, extension, parser, source_name):
        """Process every ``<extension>`` message in a local folder.

        Shared implementation for :meth:`process_eml_folder` and
        :meth:`process_msg_folder`.  Each file is parsed by ``parser``
        into the adapter's 5-key dict and then run through the same
        per-message pipeline ``crawl_pst`` uses: sender extraction →
        domain-filter / subject-keyword gate → ``process_text``
        preferring the html body over the plain one.  A malformed file
        logs to ``error_log`` and never aborts the folder.
        """
        if not path or not os.path.exists(path):
            return
        try:
            files = [f for f in os.listdir(path) if f.lower().endswith(extension)]
        except OSError as e:
            self.log_error(f"{source_name} folder: {path}", str(e))
            return
        # Sort for a deterministic progress narrative across runs
        # (os.listdir's filesystem order varies by platform).
        files.sort(key=lambda f: f.lower())
        total = len(files)

        for idx, fname in enumerate(files):
            if self.is_cancelled():
                return
            file_path = os.path.join(path, fname)
            try:
                data = parser(file_path)
            except Exception as e:
                self.log_error(f"{source_name}: {fname}", str(e))
                continue
            try:
                raw_sender = str(data.get("sender") or "")
                sender_match = EMAIL_ADDR_RE.search(raw_sender)
                sender_email = (
                    sender_match.group(1).lower() if sender_match else raw_sender.lower().strip()
                )
                subj = str(data.get("subject") or "")
                date_str = str(data.get("date_str") or "") or "N/A"

                # Same gate as crawl_pst: domain filter when configured,
                # otherwise the EDF subject-keyword heuristic.
                use_domain = self.config.get("use_domain_filter", False)
                domain_str = self.config.get("domain_filter", "")
                should_process = False
                if use_domain and domain_str:
                    if _matches_domain_filter(sender_email, domain_str):
                        should_process = True
                else:
                    if any(
                        k in subj.upper()
                        for k in ["EDF", "BILL", "STATEMENT", "ACCOUNT", "INVOICE"]
                    ):
                        should_process = True

                if should_process:
                    with self.lock:
                        self.email_count += 1
                    body_html = data.get("body_html")
                    body_text = data.get("body_text")
                    if body_html:
                        clean = BeautifulSoup(body_html, "html.parser").get_text(separator=" ")
                        self.process_text(clean, "Email Body", subj, date_str, sender=sender_email)
                    elif body_text:
                        self.process_text(
                            body_text, "Email Body", subj, date_str, sender=sender_email
                        )
                    else:
                        self.log_error(f"Email: {subj} ({date_str})", "No readable body")
            except Exception as e:
                self.log_error(f"Email: {subj}", str(e))

            if self.update_progress and idx % 100 == 0:
                self.update_progress(
                    idx + 1, total, f"Scanning {source_name} folder: {idx + 1}/{total}"
                )

        self.update_ui(f"{source_name} folder: {total} message(s) scanned")

    def process_eml_folder(self, path):
        """Process every ``.eml`` message in a local folder.

        ``.eml`` files are parsed with the stdlib-backed
        :mod:`edf_bill_fetcher.io.adapters.eml` adapter, so this works
        in any environment — no optional dependency involved.
        """
        self._process_message_folder(path, ".eml", parse_eml_message, "EML")

    def process_msg_folder(self, path):
        """Process every ``.msg`` message in a local folder.

        ``.msg`` parsing needs the optional ``extract-msg`` library
        (adapter flag ``HAS_EXTRACT_MSG``).  When it is missing the
        folder is skipped with a single error entry — the same
        loud-but-continue behaviour as ``process_pst_file``'s missing
        ``pypff`` guard.
        """
        if not HAS_EXTRACT_MSG:
            self.log_error(
                "MSG",
                "extract-msg not installed — cannot parse .msg files; "
                "install it with `pip install 'edf-bill-fetcher[msg]'`",
            )
            return
        self._process_message_folder(path, ".msg", parse_msg_message, "MSG")


# ---------------------------------------------------------------------------
# Excel helpers
# ---------------------------------------------------------------------------

THIN = _Side(style="thin", color="DDDDDD")
