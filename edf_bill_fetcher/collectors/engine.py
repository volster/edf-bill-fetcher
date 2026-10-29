"""EvidenceEngine — orchestrator for PDF/PST/text extraction.

Extracted from ``edf_collector.py`` as part of the modularization refactor
(Task 7 — Phase 6). The class plus its 7 tightly-coupled helper functions
live here; backward-compat re-exports remain in ``edf_collector.py`` until
Task 8 strips them.
"""

from __future__ import annotations

import base64
import hashlib
import os
import re
import sys
import tempfile
import threading
from pathlib import Path

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
from bs4.dammit import EncodingDetector, UnicodeDammit
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
from edf_bill_fetcher.helpers.identity import account_matches, document_identity
from edf_bill_fetcher.helpers.pst_resources import (
    EMAIL_ADDR_RE,
    extract_sender_email,
    pst_attachment_filename,
)
from edf_bill_fetcher.io.adapters.eml import parse_eml_message
from edf_bill_fetcher.io.adapters.html import (
    is_pdf_account_history,
    parse_htm_account_history,
    parse_pdf_account_history,
)
from edf_bill_fetcher.io.adapters.msg import HAS_EXTRACT_MSG, parse_msg_message
from edf_bill_fetcher.io.adapters.ocr import OcrCancelled, PdfOcrReader, requires_ocr
from edf_bill_fetcher.io.adapters.pdf import (
    extract_admit_phrase,
    join_sap_pdf_pages,
    slice_pdf_pages,
)
from edf_bill_fetcher.io.persistence.run_manifest import is_run_directory
from edf_bill_fetcher.models.config import ConfigDict
from edf_bill_fetcher.models.records import BillingRecord
from edf_bill_fetcher.processors.bill_facts import bill_facts, payment_history
from edf_bill_fetcher.processors.detection import detect_pdf_format
from edf_bill_fetcher.processors.message_amounts import message_amount_context
from edf_bill_fetcher.processors.patterns import (
    _AMOUNT_PATTERN_NEW_BILL,
    _AMOUNT_PATTERN_ONGOING_BALANCE,
    _POUND_AMOUNT_FALLBACK_RE,
    AMOUNT_PATTERNS,
    PERIOD_RE,
    READING_PATTERNS,
    extract_sub_periods,
)
from edf_bill_fetcher.processors.preparation import normalize_record
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
    r"total charges for this (?:period|bill|invoice)\s+£\s?([\d,]+(?:\.\d{2})?)(?:\s+(?:in\s+)?(credit|debit))?",
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
        self.source_manifest: dict[str, dict[str, Any]] = {}
        self._source_context: dict[str, Any] = {}
        self.attachment_count = 0
        self.skipped_attachment_count = 0
        self._attachment_dir: tempfile.TemporaryDirectory[str] | None = None
        self._ocr_reader: PdfOcrReader | None = None
        self._archive_origin = ""

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
            "source_manifest": self.source_manifest,
            "attachment_count": self.attachment_count,
            "skipped_attachment_count": self.skipped_attachment_count,
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
        self.source_manifest = state.get("source_manifest", {})
        self.attachment_count = state.get("attachment_count", 0)
        self.skipped_attachment_count = state.get("skipped_attachment_count", 0)
        self._source_context = {}
        self._attachment_dir = None
        self._ocr_reader = None
        # Rebuild runtime primitives fresh — the persisted snapshot
        # does not carry cancel state forward.
        self.cancel_event = threading.Event()
        self.lock = threading.Lock()
        # GUI callbacks don't survive a CLI↔CLI round-trip; a GUI
        # consumer can install its own after loading the snapshot via
        # ``engine.update_ui = my_gui_callback``.
        self.update_ui = lambda *_a, **_kw: None
        self.update_progress = lambda *_a, **_kw: None

    def cleanup(self) -> None:
        """Release extracted attachment files after all exports have finished."""
        if self._attachment_dir is not None:
            self._attachment_dir.cleanup()
            self._attachment_dir = None
        self._ocr_reader = None

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
        """Preserve full observations and consistently enforce account selection."""
        rec = normalize_record({**self._source_context, **rec})
        if self.config.get("use_acc_filter") and self.config.get("acc_num"):
            if not account_matches(self.config["acc_num"], rec):
                self.log_error(
                    "Account filter",
                    f"Unmatched or unidentified account in {rec.get('Attachment Name')}",
                )
                return
        amount = rec["Transaction Amount (£)"]
        if amount is None:
            amount = rec["Balance (£)"]
        if (
            amount is not None
            and self.config.get("filter_below", True)
            and abs(amount) < self.config.get("min_amount", 50.0)
        ):
            rec["Reason"] = (
                f"Amount magnitude below £{self.config.get('min_amount', 50.0):,.2f} threshold"
            )
            with self.lock:
                self.filtered_records.append(rec)
        else:
            with self.lock:
                self.records.append(rec)

    # ------------------------------------------------------------------
    # New-format PDF processing
    # ------------------------------------------------------------------

    def _process_new_invoice(
        self, text, source_label, detail_label, fallback_date, sender="", attachment_name=""
    ):
        self._source_context.update(document_identity(text))
        fields = extract_new_invoice_fields(text)
        if "amount" not in fields and "period_charge" not in fields:
            return False  # didn't match

        sub_periods = extract_sub_periods(text)
        facts = bill_facts(text, sub_periods)
        facts["Document Role"] = "Invoice"
        facts["Document Date"] = fields.get("date", fallback_date)

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
                amount=fields.get("amount"),
                amount_kind="Balance",
                energy_unit_rate=self._energy_rate(text),
                period_charge=fields.get("period_charge", "N/A"),
                entry_type=entry_type,
                reading=facts.get("Reading", r_type),
                units_kwh=facts.get("Units (kWh)", fields.get("units_used", "N/A")),
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
            | facts
        )
        return True

    def _process_new_credit(
        self, text, source_label, detail_label, fallback_date, sender="", attachment_name=""
    ):
        self._source_context.update(document_identity(text))
        fields = extract_new_credit_fields(text)
        if "amount" not in fields and "period_charge" not in fields:
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
                amount=-abs(fields["amount"]),
                amount_kind="Transaction",
                transaction_amount=-abs(fields["amount"]),
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
        self._source_context.update(document_identity(text))

        # Terms and conditions are supporting evidence, not a bill whose
        # liability limit should be promoted to a financial observation.
        if re.search(
            r"(?:^|\n)\s*(?:terms (?:and|&) conditions|THE PRINCIPAL TERMS|Contract for the supply of energy|Conditions of supplying energy)",
            text[:1000],
            re.I,
        ) and not _BILL_MARKERS_RE.search(clean_text.lower()):
            self._add_record(
                {
                    "Source": source_type,
                    "Sender": sender,
                    "Date": "N/A",
                    "Attachment Name": attachment_name or "N/A",
                    "Entry Type": "Supporting Document",
                    "Document Role": "Terms and conditions",
                    "Amount Kind": "Not monetary",
                    "Details": detail[:60],
                    "Logic Used": "Supporting document classification",
                    "Source PDF Text": text[:4000],
                }
            )
            return

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
                        if name in _AMOUNT_PATTERN_ONGOING_BALANCE and re.match(
                            r"\s+(?:in\s+)?credit\b", clean_text[m.end(1) :], re.I
                        ):
                            found_amt = -found_amt
                        strategy = "Smart Context"
                        matched_pattern_name = name
                        break
                    except Exception:
                        continue

            closing = re.search(
                r"Your new account balance\s+£\s*([\d,]+(?:\.\d{2})?)(?:\s+(?:in\s+)?(credit|debit)\b)?",
                clean_text,
                re.I,
            )
            if closing:
                found_amt = float(closing[1].replace(",", ""))
                if (closing[2] or "").lower() == "credit":
                    found_amt = -found_amt
                matched_pattern_name, strategy = "your_new_account_balance", "Smart Context"

        message_context = None
        if (
            source_type.startswith("Email Body")
            and self.config.get("use_anchors", True)
            and matched_pattern_name in {None, "balance_within"}
            and not re.search(r"\bcurrent balance\s*(?:is|:)?\s*£", clean_text, re.I)
        ):
            message_context = message_amount_context(clean_text, detail)
            if message_context is not None:
                if message_context["Amount Kind"] == "Not monetary":
                    self._add_record(
                        {
                            "Source": source_type,
                            "Sender": sender,
                            "Date": fallback_date,
                            "Attachment Name": attachment_name or "N/A",
                            "Details": detail[:60],
                            "Logic Used": "Explicit supporting email amount",
                            "Source PDF Text": clean_text[:4000],
                            **message_context,
                        }
                    )
                    return
                found_amt = message_context["Amount (£)"]
                matched_pattern_name, strategy = "payment_reminder_demand", "Explicit email demand"

        if found_amt is None and self.config.get("use_large", True):
            matches = _POUND_AMOUNT_FALLBACK_RE.findall(clean_text)
            if matches:
                floats = [float(x.replace(",", "")) for x in matches]
                highs = [x for x in floats if x >= self.config.get("min_amount", 50.0)]
                if highs:
                    found_amt = max(highs)
                    strategy = "Large Amount Fallback"

        if found_amt is None:
            return

        renewal = bool(
            re.search(
                r"Important news about your energy contract|Your business has transferred on to our variable",
                text,
                re.I,
            )
        )
        if renewal:
            balance_match = re.search(
                r"(?:outstanding balance of|currently)\s+£([\d,]+\.\d{2})", text, re.I
            )
            if not balance_match:
                self._add_record(
                    {
                        "Source": source_type,
                        "Sender": sender,
                        "Attachment Name": attachment_name or "N/A",
                        "Date": "N/A",
                        "Entry Type": "Supporting Document",
                        "Document Role": "Contract notice",
                        "Amount Kind": "Not monetary",
                        "Details": detail[:60],
                        "Logic Used": "Contract notice: no stated balance extracted",
                        "Source PDF Text": text[:4000],
                    }
                )
                return
            found_amt = float(balance_match[1].replace(",", ""))
            matched_pattern_name, strategy = (
                "your_new_account_balance",
                "Stated balance in contract notice",
            )

        # Date extraction
        date_to_use = fallback_date
        if "PDF" in source_type or "old" in source_type.lower():
            date_m = _OLD_PDF_DATE_RE.search(clean_text)
            if date_m is None:
                date_m = re.match(r"\s*(\d{1,2}\s+[A-Za-z]+\s+\d{4})\b", text)
            if date_m:
                date_to_use = parse_to_display_date(date_m.group(1))

        r_type = "Unknown"
        if self.config.get("use_reading_classification", True):
            for label, pat in READING_PATTERNS.items():
                if pat.search(clean_text):
                    r_type = label
                    break

        if source_type.startswith("Email Body"):
            # Advice about what to do IF a bill is estimated, or advertising a
            # smart meter, does not classify the reading used for this bill.
            r_type = "Unknown"
            if self.config.get("use_reading_classification", True):
                for label, phrase in (
                    (
                        "Estimated",
                        r"(?:this|your) bill (?:is |was )?based on (?:an? |our )?estimat(?:e|ed readings?)",
                    ),
                    (
                        "Smart",
                        r"(?:this|your) bill (?:uses|is based on) (?:your )?smart (?:meter )?readings?",
                    ),
                    (
                        "Actual",
                        r"(?:this|your) bill (?:uses|is based on) (?:an? |your )?actual (?:meter )?readings?",
                    ),
                ):
                    for match in re.finditer(phrase, clean_text, re.I):
                        prefix = re.split(
                            r"[.!?;:]", clean_text[max(0, match.start() - 100) : match.start()]
                        )[-1]
                        if not re.search(r"\b(?:if|whether|when)\b", prefix, re.I):
                            r_type = label
                            break
                    if r_type != "Unknown":
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
                if (pc_m.group(2) or "").lower() == "credit":
                    period_charge = -period_charge
            except (ValueError, AttributeError):
                pass

        # Mine per-sub-period rows ("About your charges") from old-format
        # cancel-and-rebill invoices too.  The T-series bills are OLD-format
        # and arrive via this generic path (never ``_process_new_invoice``),
        # so without this the ``Sub Periods`` column stays empty and the
        # Option C sub-period unlawful-charge computation never activates.
        sub_periods = extract_sub_periods(text)
        facts = bill_facts(text, sub_periods)
        units_used = facts.pop("Units (kWh)", units_used)
        r_type = facts.pop("Reading", r_type)

        # Classify Entry Type based on content
        entry_type = self._classify_entry_type(
            clean_text, matched_pattern_name, period_from, period_to, strategy
        )
        if renewal:
            entry_type = "Ongoing Balance"
            period_from = period_to = "N/A"
            units_used = standing_charge = "N/A"
            period_charge = "N/A"
            facts = {"Document Role": "Contract notice", "Document Date": date_to_use}
            sub_periods = []
        if matched_pattern_name == "payment_reminder_demand":
            entry_type = "Ongoing Balance"
            period_charge = "N/A"
            facts.update(
                {
                    "Document Role": "Payment reminder",
                    "Demand Amount (£)": found_amt,
                    "Balance Basis": "Demand repeated in a reminder; not a new charge",
                }
            )
            if message_context is not None:
                facts["Demand Context"] = message_context["Demand Context"]
        if entry_type == "New Bill" and "PDF" in source_type:
            facts["Document Role"] = "Invoice"
            facts["Document Date"] = date_to_use

        self._add_record(
            BillingRecord(
                source=source_type,
                sender=sender,
                date=date_to_use,
                period_from=period_from,
                period_to=period_to,
                invoice_num=inv_num,
                amount=found_amt,
                amount_kind="Balance"
                if matched_pattern_name in _AMOUNT_PATTERN_ONGOING_BALANCE
                or matched_pattern_name == "current_balance_debit"
                else "Unknown",
                energy_unit_rate=None if renewal else self._energy_rate(clean_text),
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
            | facts
        )

        if inv_num != "N/A" and period_charge != "N/A":
            for payment in payment_history(text):
                self._add_record(
                    {
                        "Source": source_type,
                        "Sender": sender,
                        "Attachment Name": attachment_name or "N/A",
                        "Source Invoice #": inv_num,
                        "Document Date": date_to_use,
                        **payment,
                    }
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

    @staticmethod
    def _energy_rate(text: str) -> float | None:
        """Extract a stated energy unit rate; total-charge/kWh is not a rate."""
        rates = re.findall(
            r"(?:Electricity used|Energy used|Unit rate).*?(?:@|at)\s*([\d.]+)\s*p(?:/|\s*per\s*)kWh",
            text,
            re.I,
        )
        values = {float(value) for value in rates}
        return next(iter(values)) if len(values) == 1 else None

    def _selected_sap_rows(self, rows: list[dict]) -> list[dict]:
        """Apply the selected account to each row in a possibly mixed SAP dump."""
        result = []
        for row in rows:
            if (
                self.config.get("use_acc_filter")
                and self.config.get("acc_num")
                and not account_matches(self.config["acc_num"], row)
            ):
                self.log_error("SAP account filter", f"Excluded account {row.get('Account')!r}")
                continue
            result.append({**self._source_context, **row})
        return result

    def _register_source(self, path: str, name: str, content: bytes) -> dict[str, Any]:
        """Record a source occurrence using content and origin identity."""
        digest = hashlib.sha256(content).hexdigest()
        origin = self._source_context.get("Parent Message") or os.path.abspath(path)
        source_id = hashlib.sha256((origin + "\0" + name + "\0" + digest).encode()).hexdigest()
        original_name = name or os.path.basename(path)
        unique_name = original_name
        if unique_name in self.source_paths and self.source_paths[unique_name] != path:
            stem, ext = os.path.splitext(original_name)
            unique_name = f"{stem}-{source_id[:12]}{ext}"
        self.source_paths[unique_name] = path
        self.source_paths[source_id] = path
        metadata = {
            "Source ID": source_id,
            "Content Hash": digest,
            "Original Path": os.path.abspath(path),
            "Original Filename": original_name,
            "Attachment Name": unique_name,
            "Parent Message": self._source_context.get("Parent Message", ""),
        }
        self.source_manifest[source_id] = metadata.copy()
        return metadata

    def _process_attachment(
        self, data: bytes, name: str, source: str, date_str: str, sender: str, parent: str = ""
    ) -> None:
        """Keep supported attachments alive until all bundle writers finish."""
        self.attachment_count += 1
        if not data.startswith(b"%PDF"):
            self.skipped_attachment_count += 1
            if name.lower().endswith(".pdf"):
                self.log_error(
                    "Attachment",
                    f"Attachment {name!r} in {parent!r} is empty or has no PDF header; PDF not extracted",
                )
            return
        if self._attachment_dir is None:
            self._attachment_dir = tempfile.TemporaryDirectory(prefix="edf-attachments-")
        with tempfile.NamedTemporaryFile(
            delete=False, suffix=".pdf", dir=self._attachment_dir.name
        ) as handle:
            handle.write(data)
            path = handle.name
        self._source_context = {"Parent Message": parent}
        self.pdf_count += 1
        self.process_pdf_file(path, source, name, date_str, sender=sender, attachment_name=name)

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
            parent_message = self._source_context.get("Parent Message", "")
            self._source_context = self._register_source(path, attachment_name or detail_label, raw)
            self._source_context["Parent Message"] = parent_message
            attachment_name = self._source_context["Attachment Name"]
            pdf_hash = self._source_context["Content Hash"]
            with self.lock:
                if pdf_hash in self.seen_pdf_hashes:
                    for record in self.records + self.filtered_records:
                        if record.get("Content Hash") == pdf_hash:
                            record.setdefault("Source IDs", []).append(
                                self._source_context["Source ID"]
                            )
                            record.setdefault("Source Occurrences", []).append(
                                self._source_context.copy()
                            )
                    return
                self.seen_pdf_hashes.add(pdf_hash)

            with pdfplumber.open(io.BytesIO(raw)) as pdf:
                # Handle empty or corrupt PDFs gracefully
                if not pdf.pages:
                    self.log_error(f"PDF: {detail_label}", "PDF has no pages")
                    return
                pdf_text_parts: list[str] = []
                ocr_pages: dict[int, dict[str, Any]] = {}
                history_pdf = False
                for page_index, p in enumerate(pdf.pages):
                    if self.is_cancelled():
                        return
                    try:
                        page_text = p.extract_text()
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
                        page_text = ""
                        self.log_error(
                            f"PDF page {detail_label}", f"Page extraction failed: {page_err}"
                        )
                    history_pdf = history_pdf or is_pdf_account_history(page_text or "")
                    # OCR only empty pages or image pages with a sparse text
                    # header. A short native amount-bearing page stays native.
                    if requires_ocr(page_text, p):
                        page_number = page_index + 1
                        context = f"OCR {detail_label}, page {page_number}"
                        self.update_ui(f"{context}: recognizing scanned text…")
                        if self._ocr_reader is None:
                            self._ocr_reader = PdfOcrReader()
                        try:
                            recognized = self._ocr_reader.read_page(
                                raw, page_index, self.is_cancelled
                            )
                            if self.is_cancelled():
                                return
                            ocr_pages[page_number] = {
                                "Status": "recognized" if recognized.text.strip() else "no text",
                                "Text": recognized.text,
                                "Minimum Confidence (%)": recognized.minimum_confidence,
                                "Engine": "Tesseract (English)",
                            }
                            if recognized.text.strip():
                                native_identity = document_identity(page_text or "")
                                native_history = is_pdf_account_history(page_text or "")
                                page_text = recognized.text
                                # Preserve a readable native customer URL:
                                # OCR can insert spaces inside its account ID.
                                if native_identity["Account"] and native_history:
                                    page_text += f"\nAccount number: {native_identity['Account']}"
                                self.log_error(
                                    context,
                                    "OCR-generated evidence: verify amounts, dates and invoice references against the original. "
                                    f"Lowest word confidence: {recognized.minimum_confidence}% (not proof of accuracy).",
                                )
                            else:
                                self.log_error(
                                    context,
                                    "No readable text found; page may be blank or require manual review",
                                )
                        except OcrCancelled:
                            return
                        except Exception as error:
                            ocr_pages[page_number] = {"Status": "failed", "Error": str(error)}
                            self.log_error(context, str(error))
                    pdf_text_parts.append(page_text or "")
                if ocr_pages:
                    self.source_manifest[self._source_context["Source ID"]]["OCR Page Results"] = {
                        str(number): result for number, result in ocr_pages.items()
                    }
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

            first_page = 1
            for i, slice_pages in enumerate(slices, start=1):
                if self.is_cancelled():
                    return
                slice_text = " ".join(slice_pages)
                self._source_context.update(document_identity(slice_text))
                self._source_context["Page From"] = first_page
                self._source_context["Page To"] = first_page + len(slice_pages) - 1
                covered_ocr = {
                    number: result
                    for number, result in ocr_pages.items()
                    if first_page <= number < first_page + len(slice_pages)
                }
                for key in ("OCR Pages", "OCR Minimum Confidence (%)", "OCR Review Required"):
                    self._source_context.pop(key, None)
                self._source_context["Text Extraction"] = "PDF text"
                if covered_ocr:
                    self._source_context["Text Extraction"] = (
                        "OCR" if len(covered_ocr) == len(slice_pages) else "PDF text + OCR"
                    )
                    self._source_context["OCR Pages"] = list(covered_ocr)
                    self._source_context["OCR Review Required"] = True
                    scores = [
                        result["Minimum Confidence (%)"]
                        for result in covered_ocr.values()
                        if result.get("Minimum Confidence (%)") is not None
                    ]
                    self._source_context["OCR Minimum Confidence (%)"] = (
                        min(scores) if scores else None
                    )
                first_page += len(slice_pages)
                if multi:
                    slice_detail = f"{detail_label} #{i}"
                    slice_attachment = f"{attachment_name} #{i}"
                else:
                    slice_detail = detail_label
                    slice_attachment = attachment_name
                self.source_paths[slice_attachment] = path

                try:
                    if history_pdf or is_pdf_account_history(slice_text):
                        extracted = 0
                        for page_offset, page_text in enumerate(slice_pages):
                            page_number = self._source_context["Page From"] + page_offset
                            for source_row, record in enumerate(
                                parse_pdf_account_history(page_text), 1
                            ):
                                self._add_record(
                                    {
                                        **record,
                                        "Attachment Name": slice_attachment,
                                        "Source Row": f"page:{page_number}:event:{source_row}",
                                        "Page From": page_number,
                                        "Page To": page_number,
                                        "OCR Pages": [page_number]
                                        if page_number in ocr_pages
                                        else [],
                                        "Text Extraction": "OCR"
                                        if page_number in ocr_pages
                                        else "PDF text",
                                        "OCR Review Required": page_number in ocr_pages,
                                        "OCR Minimum Confidence (%)": ocr_pages.get(
                                            page_number, {}
                                        ).get("Minimum Confidence (%)"),
                                    }
                                )
                                extracted += 1
                        if not extracted:
                            self.log_error(
                                f"PDF: {slice_detail}",
                                "Account history dates or transaction headings are unreadable; manual review required",
                            )
                            self._add_record(
                                {
                                    "Source": source_label,
                                    "Attachment Name": slice_attachment,
                                    "Entry Type": "Supporting Document",
                                    "Document Role": "Account history: incomplete text",
                                    "Amount Kind": "Not monetary",
                                    "Details": "Account history requires manual review; no transactions extracted",
                                }
                            )
                        continue
                    # Stream P1: SAP CSV-in-PDF data dumps (Contract /
                    # Meter-Read / Financial-Transactions). Detected
                    # via the header-row marker; routed to dedicated
                    # parsers and stored on the engine's SAP-row
                    # accumulators (not engine.records -- they get
                    # their own dedicated Excel sheets via
                    # ``export_to_excel(..., sap_rows=...)``).
                    sap_kind = detect_sap_dump(slice_text)
                    if sap_kind is not None:
                        slice_text, repaired_pages = join_sap_pdf_pages(slice_pages)
                        if repaired_pages:
                            self.source_manifest[self._source_context["Source ID"]][
                                "SAP PDF Boundary Repairs"
                            ] = [
                                self._source_context["Page From"] + page - 1
                                for page in repaired_pages
                            ]
                        source_file = slice_attachment or slice_detail or ""

                        def report_sap(message: str, label: str = source_file) -> None:
                            self.log_error(f"SAP dump {label}", message)

                        if sap_kind == "contract":
                            self.sap_contract_rows.extend(
                                self._selected_sap_rows(
                                    parse_sap_contract_history(
                                        slice_text,
                                        source_file=source_file,
                                        report=report_sap,
                                    )
                                )
                            )
                        elif sap_kind == "meter_read":
                            self.sap_meter_rows.extend(
                                self._selected_sap_rows(
                                    parse_sap_meter_read_history(
                                        slice_text,
                                        source_file=source_file,
                                        report=report_sap,
                                    )
                                )
                            )
                        elif sap_kind == "financial":
                            self.sap_financial_rows.extend(
                                self._selected_sap_rows(
                                    parse_sap_financial_transactions(
                                        slice_text,
                                        source_file=source_file,
                                        report=report_sap,
                                    )
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
                        for source_row, row in enumerate(rows, 1):
                            row["Source Row"] = source_row
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
            self._source_context = {}
            raw = Path(path).read_bytes()
            self._source_context = self._register_source(path, os.path.basename(path), raw)
            # Browser exports can declare Windows-1252 or UTF-16. Honour
            # their encoding and preserve the pound sign used by the parser.
            _, bom_encoding = EncodingDetector.strip_byte_order_mark(raw)
            declared = EncodingDetector.find_declared_encoding(raw, is_html=True)
            encodings = [bom_encoding] if bom_encoding else []
            if declared:
                encodings.append(declared)
            encodings.extend(["utf-8", "windows-1252"])
            decoded = UnicodeDammit(raw, known_definite_encodings=encodings, is_html=True)
            try:
                raw.decode("utf-8")
            except UnicodeDecodeError:
                if not declared and not raw.startswith((b"\xff\xfe", b"\xfe\xff")):
                    self.log_error(
                        f"HTM: {path}",
                        f"UTF-8 decode error — read as {decoded.original_encoding}",
                    )
            if decoded.contains_replacement_characters:
                self.log_error(f"HTM: {path}", "Some undecodable characters replaced")
            soup = BeautifulSoup(decoded.unicode_markup or "", "html.parser")
            text = soup.get_text(separator=" ", strip=True)
            account_links = " ".join(
                str(a.get("href", ""))
                for a in soup.select("a[href]")
                if re.match(
                    r"https://edfenergy\.com/myaccount/accounts/A-\d+\b",
                    str(a.get("href", "")),
                    re.I,
                )
            )
            self._source_context.update(document_identity(text + " " + account_links))
            recs = parse_htm_account_history(text)
            for source_row, rec in enumerate(recs, 1):
                rec["Attachment Name"] = self._source_context["Attachment Name"]
                rec["Source Row"] = source_row
                rec["Document Role"] = "Historical account statement"
                rec["Historical Activity"] = True
                rec["Posting Date"] = rec["Date"]
                rec["First Billed Date"] = "N/A"
                rec["Billing Date Status"] = (
                    "Unknown: account history posting date is not an invoice issue date"
                )
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
            try:
                pst.open(path)
                self._archive_origin = os.path.abspath(path)
                root = pst.get_root_folder()
                self.crawl_pst(root)
            finally:
                self._archive_origin = ""
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

    def _process_pst_body(
        self, msg: Any, subj: str, date_str: str, sender: str, context: str
    ) -> None:
        """Try body representations independently without blocking attachment recovery."""
        for label, getter, source in (
            ("HTML", "get_html_body", "Email Body"),
            ("plain text", "get_plain_text_body", "Email Body"),
            ("RTF", "get_rtf_body", "Email Body (RTF)"),
        ):
            try:
                body = getattr(msg, getter)()
                if not body:
                    continue
                if label == "HTML":
                    body_text = BeautifulSoup(body, "html.parser").get_text(separator=" ")
                else:
                    if isinstance(body, bytes):
                        decoded = UnicodeDammit(
                            body, known_definite_encodings=["utf-8", "windows-1252"]
                        )
                        body_text = decoded.unicode_markup or ""
                        if (
                            decoded.original_encoding not in {"utf-8", "ascii"}
                            or decoded.contains_replacement_characters
                        ):
                            self.log_error(
                                f"PST {label} body ({context})",
                                f"Body decoded as {decoded.original_encoding}; review extracted text",
                            )
                    else:
                        body_text = str(body)
                    if label == "RTF":
                        body_text = re.sub(r"\\[a-z]+[-\d]*\s?", " ", body_text)
                        body_text = re.sub(r"[{}\\]", " ", body_text)
                if not body_text.strip():
                    continue
                if self._attachment_dir is None:
                    self._attachment_dir = tempfile.TemporaryDirectory(prefix="edf-attachments-")
                archive = getattr(self, "_archive_origin", "")
                origin = f"{archive}\0{context}"
                raw_body = body if isinstance(body, bytes) else str(body).encode("utf-8")
                exported = (
                    "Generated PST/OST message body export; not an original .eml file\n"
                    f"Archive: {archive}\nOccurrence: {context}\n"
                    f"From: {sender}\nSubject: {subj}\nDate: {date_str}\n"
                    f"Body representation: {label}\n"
                    f"Captured body base64 (UTF-8 if pypff returned text): {base64.b64encode(raw_body).decode('ascii')}\n\n{body_text}"
                ).encode()
                name = f"message-{hashlib.sha256(origin.encode()).hexdigest()[:16]}.txt"
                path = os.path.join(self._attachment_dir.name, name)
                with open(path, "wb") as handle:
                    handle.write(exported)
                self._source_context = {"Parent Message": origin}
                self._source_context = self._register_source(path, name, exported)
                self._source_context["Source Representation"] = "Generated PST/OST body text export"
                self.source_manifest[self._source_context["Source ID"]]["Source Representation"] = (
                    self._source_context["Source Representation"]
                )
                self.process_text(
                    body_text, source, subj, date_str, sender=sender, attachment_name=name
                )
                return
            except Exception as error:
                self.log_error(f"PST {label} body ({context})", str(error))
        self.log_error(
            f"Email: {subj} ({date_str})",
            "No readable body; attachment extraction will still be attempted",
        )

    def crawl_pst(self, folder: Any, *, folder_path: str = "") -> None:
        """Recursively walk a PST/OST folder and process every EDF email found."""
        if not HAS_PYPFF:
            self.log_error("PST", "pypff not installed — skipping PST processing")
            return
        if self.is_cancelled():
            return

        try:
            name = folder.get_name()
        except Exception:
            name = None
        name = name if isinstance(name, str) and name else "Unnamed folder"
        folder_label = f"{folder_path}/{name}" if folder_path else name

        msg_total = folder.get_number_of_sub_messages()
        for i in range(msg_total):
            if self.is_cancelled():
                return
            try:
                msg = folder.get_sub_message(i)
                self._source_context = {}
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
                    context = f'{folder_label}, message index {i}, "{subj}" ({date_str})'
                    self._process_pst_body(msg, subj, date_str, sender_email, context)
                    try:
                        attachment_total = msg.get_number_of_attachments()
                    except Exception as error:
                        self.log_error(
                            f"PST attachments ({context})",
                            "Unable to enumerate attachments; attachments may be missing. Any readable body was processed separately. "
                            + str(error),
                        )
                        continue
                    for a_idx in range(attachment_total):
                        if self.is_cancelled():
                            return
                        try:
                            att = msg.get_attachment(a_idx)
                            size = att.get_size()
                            buf = att.read_buffer(size) if size else b""
                            buf = buf or b""
                            att_name = _pst_attachment_filename(att) or (
                                f"Attachment_{a_idx + 1}.pdf"
                                if buf.startswith(b"%PDF")
                                else f"Attachment_{a_idx + 1}"
                            )
                            self._process_attachment(
                                buf,
                                att_name,
                                "PST PDF Attachment",
                                date_str,
                                sender_email,
                                f"{getattr(self, '_archive_origin', '')} | {context}",
                            )
                        except Exception as e:
                            self.log_error(
                                f'Attachment in "{subj}" ({folder_label}, message index {i}, attachment index {a_idx})',
                                str(e),
                            )

            except Exception as e:
                self.log_error(f"PST message index {i} ({folder_label})", str(e))

        self.update_ui(
            f"Scanned {self.email_count} emails, {self.pdf_count} attached PDFs; skipped {self.skipped_attachment_count} attachments…"
        )
        for j in range(folder.get_number_of_sub_folders()):
            if self.is_cancelled():
                return
            self.crawl_pst(folder.get_sub_folder(j), folder_path=folder_label)

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
        for root, dirs, files in os.walk(path):
            if is_run_directory(root):
                dirs.clear()
                self.update_ui("Skipping earlier evidence pack: " + root)
                continue
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
            self._source_context = {}
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
            files: list[str] = []
            for base, dirs, names in os.walk(path):
                if is_run_directory(base):
                    dirs.clear()
                    self.update_ui("Skipping earlier evidence pack: " + base)
                    continue
                files.extend(
                    os.path.relpath(os.path.join(base, name), path)
                    for name in names
                    if name.lower().endswith(extension)
                )
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
                self._source_context = {}
                with open(file_path, "rb") as message_file:
                    self._source_context = self._register_source(
                        file_path, fname, message_file.read()
                    )
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
                        self.process_text(
                            clean,
                            "Email Body",
                            subj,
                            date_str,
                            sender=sender_email,
                            attachment_name=self._source_context.get("Attachment Name", ""),
                        )
                    elif body_text:
                        self.process_text(
                            body_text,
                            "Email Body",
                            subj,
                            date_str,
                            sender=sender_email,
                            attachment_name=self._source_context.get("Attachment Name", ""),
                        )
                    else:
                        self.log_error(f"Email: {subj} ({date_str})", "No readable body")
                    for attachment in data.get("attachments", []):
                        if self.is_cancelled():
                            break
                        self._process_attachment(
                            attachment["data"],
                            attachment["name"],
                            f"{source_name} PDF Attachment",
                            date_str,
                            sender_email,
                            file_path,
                        )
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
