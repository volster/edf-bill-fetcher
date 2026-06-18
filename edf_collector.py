#!/usr/bin/env python3
"""
EDF Master Evidence Collector
Collects billing data from PST/OST files, local PDF folders, and HTM account exports.
Fixed version: correct Excel date serials, dynamic range references, new PDF format support.
"""

import gc
import hashlib
import os
import re
import tempfile
import threading
import tkinter as tk
import traceback
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

import numpy as np
import openpyxl
import pandas as pd
import pdfplumber
from bs4 import BeautifulSoup
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# Optional imports — gracefully degrade if missing
try:
    import pypff

    HAS_PYPFF = True
except ImportError:
    HAS_PYPFF = False

# Try to import scipy for advanced stats (graceful fallback)
try:
    import importlib.util

    HAS_SCIPY = importlib.util.find_spec("scipy") is not None
except ImportError:
    HAS_SCIPY = False

# Try to import statsmodels for forecasting (graceful fallback)
try:
    from statsmodels.tsa.holtwinters import ExponentialSmoothing

    HAS_STATSMODELS = True
except ImportError:
    HAS_STATSMODELS = False

# Optional import for PDF report generation
try:
    from edf_report import generate_pdf_from_gui

    HAS_PDF_REPORT = True
except ImportError:
    HAS_PDF_REPORT = False


# ---------------------------------------------------------------------------
# Branding
# ---------------------------------------------------------------------------
EDF_ORANGE = "#FE5716"
EDF_NAVY = "#10367A"
EDF_OFFWHITE = "#F5F5F5"
EST_YELLOW = "FFFF99"
JUMP_RED = "FF9999"
DUP_GREY = "E0E0E0"

# ---------------------------------------------------------------------------
# Extraction patterns
# ---------------------------------------------------------------------------
AMOUNT_PATTERNS = [
    # New-style KI / KCR invoices — "Current balance £X debit"
    r"current balance\s+£\s?([\d,]+(?:\.\d{2})?)\s*(?:in\s*)?debit",
    # New-style KI — "Total charges for this period £X debit"
    r"total charges for this period\s+£\s?([\d,]+(?:\.\d{2})?)\s*(?:in\s*)?debit",
    # New-style KCR — "Total credits for this bill £X"
    r"total credits for this bill\s+£\s?([\d,]+(?:\.\d{2})?)",
    # Old-style cumulative balance
    r"your new account balance\s+£\s?([\d,]+(?:\.\d{2})?)",
    # Generic anchors
    r"balance[\s\S]{0,30}?£\s?([\d,]+(?:\.\d{2})?)",
    r"total charges[\s\S]{0,30}?£\s?([\d,]+(?:\.\d{2})?)",
    r"total amount due[\s\S]{0,30}?£\s?([\d,]+(?:\.\d{2})?)",
    r"amount to pay[\s\S]{0,30}?£\s?([\d,]+(?:\.\d{2})?)",
    r"£\s?([\d,]+(?:\.\d{2})?)\s*(?:in\s*)?debit",
    r"current balance[\s\S]{0,30}?£\s?([\d,]+(?:\.\d{2})?)",
]

READING_PATTERNS = {
    "Estimated": re.compile(r"estimated|est\.|estimate", re.IGNORECASE),
    "Actual": re.compile(r"actual|customer reading|your reading", re.IGNORECASE),
    "Smart": re.compile(r"smart meter|automated reading|smart reading", re.IGNORECASE),
}

PERIOD_RE = re.compile(
    r"(\d{1,2}(?:\s+\w+\s+\d{4}|\s*/\s*\d{2}\s*/\s*\d{4}|\s*-\s*\d{2}\s*-\s*\d{4}))"
    r"\s*(?:to|to:|–|-)\s*"
    r"(\d{1,2}(?:\s+\w+\s+\d{4}|\s*/\s*\d{2}\s*/\s*\d{4}|\s*-\s*\d{2}\s*-\s*\d{4}))",
    re.IGNORECASE,
)

_ISO_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


# ---------------------------------------------------------------------------
# Date helpers
# ---------------------------------------------------------------------------


def parse_to_sort_date(date_input):
    s = str(date_input).strip() if date_input else ""
    if not s or s in ("Unknown", "N/A", ""):
        return pd.NaT
    try:
        if _ISO_DATE_RE.match(s):
            return pd.to_datetime(s, format="%Y-%m-%d", errors="coerce")
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.isna(dt):
            dt = pd.to_datetime(s, dayfirst=False, errors="coerce")
        return dt
    except Exception:
        return pd.NaT


def parse_to_display_date(date_input):
    dt = parse_to_sort_date(date_input)
    return dt.strftime("%d/%m/%Y") if not pd.isna(dt) else str(date_input)


def to_excel_date(date_input):
    """Return a Python datetime for openpyxl to write as a true Excel date serial."""
    dt = parse_to_sort_date(date_input)
    if pd.isna(dt):
        return None
    return dt.to_pydatetime()


# ---------------------------------------------------------------------------
# Detect which EDF bill format we're looking at
# ---------------------------------------------------------------------------


def detect_pdf_format(text):
    """Return 'new_invoice', 'new_credit', or 'old' based on document markers."""
    if re.search(r"invoice number:\s*KI-", text, re.IGNORECASE):
        return "new_invoice"
    if re.search(r"credit note number:\s*KCR-", text, re.IGNORECASE):
        return "new_credit"
    return "old"


def extract_new_invoice_fields(text):
    """Extract key fields from new-style KI-XXXXXXXX invoices."""
    fields = {}

    # Invoice number
    m = re.search(r"Invoice number:\s*(KI-[\w-]+)", text, re.IGNORECASE)
    if m:
        fields["inv_num"] = m.group(1).strip()

    # Account number (A-XXXXXXXX format)
    m = re.search(r"Account number:\s*(A-[\d]+)", text, re.IGNORECASE)
    if m:
        fields["acc_num"] = m.group(1).strip()

    # Date issued
    m = re.search(r"Date issued:\s*(\d{1,2}\s+\w+\s+\d{4})", text, re.IGNORECASE)
    if m:
        fields["date"] = parse_to_display_date(m.group(1).strip())

    # Billing period from "Your charges: DD Mon YYYY - DD Mon YYYY"
    m = re.search(
        r"Your charges:\s*(\d{1,2}\s+\w+\s+\d{4})\s*[-–]\s*(\d{1,2}\s+\w+\s+\d{4})",
        text,
        re.IGNORECASE,
    )
    if m:
        fields["period_from"] = parse_to_display_date(m.group(1).strip())
        fields["period_to"] = parse_to_display_date(m.group(2).strip())

    # Current balance (the running account total — used as primary Amount)
    m = re.search(r"Current balance\s+£([\d,]+\.\d{2})\s+debit", text, re.IGNORECASE)
    if m:
        fields["amount"] = float(m.group(1).replace(",", ""))

    # Period charge (total for this invoice)
    m = re.search(r"Total charges for this period\s+£([\d,]+\.\d{2})\s+debit", text, re.IGNORECASE)
    if m:
        fields["period_charge"] = float(m.group(1).replace(",", ""))

    # kWh used
    m = re.search(r"Electricity used\s+([\d,]+\.?\d*)\s+kWh", text, re.IGNORECASE)
    if m:
        fields["units_used"] = m.group(1)

    # Standing charge
    m = re.search(r"Standing charge\s+\d+\s+days\s+@\s+([\d.]+)p/day", text, re.IGNORECASE)
    if m:
        fields["standing_charge"] = m.group(1)

    # Tariff name
    m = re.search(r"Tariff name\s+(\w[\w\s]+?)(?:Payment type|$)", text, re.IGNORECASE)
    if m:
        fields["tariff"] = m.group(1).strip()

    return fields


def extract_new_credit_fields(text):
    """Extract key fields from new-style KCR-XXXXXXXX credit notes."""
    fields = {}

    m = re.search(r"Credit note number:\s*(KCR-[\w-]+)", text, re.IGNORECASE)
    if m:
        fields["inv_num"] = m.group(1).strip()

    m = re.search(r"Account number:\s*(A-[\d]+)", text, re.IGNORECASE)
    if m:
        fields["acc_num"] = m.group(1).strip()

    m = re.search(r"Date issued:\s*(\d{1,2}\s+\w+\s+\d{4})", text, re.IGNORECASE)
    if m:
        fields["date"] = parse_to_display_date(m.group(1).strip())

    # Total credits for this bill
    m = re.search(r"Total credits for this bill\s+£([\d,]+\.\d{2})", text, re.IGNORECASE)
    if m:
        fields["amount"] = float(m.group(1).replace(",", ""))

    return fields


# ---------------------------------------------------------------------------
# HTM account-history parser
# ---------------------------------------------------------------------------


def parse_htm_account_history(text):
    """
    Parse the EDF MyAccount 'Payments and Invoices' HTM export.
    Returns a list of record dicts ready for process_text bypass.
    """
    records = []

    # We look for the recurring pattern:
    # "DD Mon YYYY We charged your account £X.XX For Y kWh … between D Mon YYYY and D Mon YYYY Balance £X.XX in debit"
    # "DD Mon YYYY You paid us £X.XX … Balance £X.XX in debit"

    # Normalise whitespace
    text = re.sub(r"\s+", " ", text)

    # Find all "charged" entries
    charge_re = re.compile(
        r"(\d{1,2}\s+\w+\s+\d{4})\s+We charged your account\s+£([\d,]+\.\d{2})"
        r"(?:\s+For\s+([\d,]+)\s+kWh\s+of\s+electricity\s+used\s+between\s+"
        r"(\d{1,2}\s+\w+\s+\d{4})\s+and\s+(\d{1,2}\s+\w+\s+\d{4}))?"
        r".*?Balance\s+£([\d,]+\.\d{2})\s+in\s+debit",
        re.IGNORECASE,
    )
    for m in charge_re.finditer(text):
        date_str = parse_to_display_date(m.group(1))
        period_from = parse_to_display_date(m.group(4)) if m.group(4) else "N/A"
        period_to = parse_to_display_date(m.group(5)) if m.group(5) else "N/A"
        units = m.group(3) if m.group(3) else "N/A"
        charge_amt = float(m.group(2).replace(",", ""))
        balance = float(m.group(6).replace(",", ""))
        records.append(
            {
                "Source": "HTM Account History",
                "Sender": "",
                "Date": date_str,
                "Period From": period_from,
                "Period To": period_to,
                "Invoice #": "N/A",
                "Amount (£)": balance,
                "Period Charge (£)": charge_amt,
                "Entry Type": "Ongoing Balance",
                "Reading": "Unknown",
                "Units (kWh)": units,
                "Standing Chg (p/day)": "N/A",
                "Attachment Name": "N/A",
                "Details": "HTM: charged account",
                "Logic Used": "HTM Charge",
            }
        )

    # Find all "You paid us" entries
    pay_re = re.compile(
        r"(\d{1,2}\s+\w+\s+\d{4})\s+You paid us\s+£([\d,]+\.\d{2})"
        r".*?Balance\s+£([\d,]+\.\d{2})\s+in\s+debit",
        re.IGNORECASE,
    )
    for m in pay_re.finditer(text):
        date_str = parse_to_display_date(m.group(1))
        balance = float(m.group(3).replace(",", ""))
        records.append(
            {
                "Source": "HTM Account History",
                "Sender": "",
                "Date": date_str,
                "Period From": "N/A",
                "Period To": "N/A",
                "Invoice #": "N/A",
                "Amount (£)": balance,
                "Period Charge (£)": "N/A",
                "Entry Type": "Payment",
                "Reading": "Unknown",
                "Units (kWh)": "N/A",
                "Standing Chg (p/day)": "N/A",
                "Attachment Name": "N/A",
                "Details": "HTM: payment received",
                "Logic Used": "HTM Payment",
            }
        )

    # Find all "reversed account charge" entries (credits applied)
    rev_re = re.compile(
        r"(\d{1,2}\s+\w+\s+\d{4})\s+Reversed account charge\s+£([\d,]+\.\d{2})"
        r".*?Balance\s+£([\d,]+\.\d{2})\s+in\s+debit",
        re.IGNORECASE,
    )
    for m in rev_re.finditer(text):
        date_str = parse_to_display_date(m.group(1))
        balance = float(m.group(3).replace(",", ""))
        records.append(
            {
                "Source": "HTM Account History",
                "Sender": "",
                "Date": date_str,
                "Period From": "N/A",
                "Period To": "N/A",
                "Invoice #": "N/A",
                "Amount (£)": balance,
                "Period Charge (£)": "N/A",
                "Entry Type": "Credit",
                "Reading": "Unknown",
                "Units (kWh)": "N/A",
                "Standing Chg (p/day)": "N/A",
                "Attachment Name": "N/A",
                "Details": "HTM: reversed account charge",
                "Logic Used": "HTM Reversal",
            }
        )

    return records


# ---------------------------------------------------------------------------
# Evidence Engine
# ---------------------------------------------------------------------------


def _extract_sender_email(msg):
    """Extract sender email address from a pypff message, trying multiple methods."""
    sender = None
    # Try transport headers first (most reliable for SMTP email address)
    try:
        headers = msg.get_transport_headers()
        if headers:
            headers_str = (
                headers if isinstance(headers, str) else headers.decode("utf-8", errors="replace")
            )
            m = re.search(
                r"^From:\s*.*?([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})",
                headers_str,
                re.MULTILINE | re.IGNORECASE,
            )
            if m:
                sender = m.group(1).lower()
    except Exception:
        pass
    # Fallback: try sender name field (sometimes contains email)
    if not sender:
        try:
            name = msg.get_sender_name() or ""
            m = re.search(r"([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})", name)
            if m:
                sender = m.group(1).lower()
        except Exception:
            pass
    return sender or ""


def _matches_domain_filter(sender_email, filter_str):
    """
    Check if sender_email matches the domain filter string.
    filter_str is comma-separated, supporting:
      - domain names: "edf.com" matches *@edf.com and *@*.edf.com
      - full addresses: "billing@edf.com" matches exactly
      - wildcard domains: "*.edf.com" matches subdomains
    """
    if not sender_email or not filter_str:
        return False
    sender_email = sender_email.lower().strip()
    parts = [p.strip().lower() for p in filter_str.split(",") if p.strip()]
    for pattern in parts:
        if "@" in pattern:
            # Full email address match
            if sender_email == pattern:
                return True
        else:
            # Domain match — check exact domain or subdomain
            domain = pattern.lstrip("*").lstrip(".")
            sender_domain = sender_email.split("@")[-1] if "@" in sender_email else ""
            if sender_domain == domain or sender_domain.endswith("." + domain):
                return True
    return False


class EvidenceEngine:
    def __init__(self, config, update_ui_cb, progress_cb=None, cancel_event=None):
        self.config = config
        self.records = []
        self.filtered_records = []
        self.update_ui = update_ui_cb
        self.update_progress = progress_cb
        self.cancel_event = cancel_event or threading.Event()
        self.pdf_count = 0
        self.email_count = 0
        self.error_log = []
        self.seen_pdf_hashes = set()
        self.lock = threading.Lock()

    def is_cancelled(self):
        return self.cancel_event.is_set()

    def log_error(self, context, err):
        self.error_log.append(f"[{datetime.now().strftime('%H:%M:%S')}] {context} — {err}")

    def find_billing_period(self, text):
        m = PERIOD_RE.search(text)
        if m:
            return (
                parse_to_display_date(m.group(1).strip()),
                parse_to_display_date(m.group(2).strip()),
            )
        return "N/A", "N/A"

    def _add_record(self, rec):
        """Thread-safe record append after optional filter check."""
        amt = rec.get("Amount (£)", 0) or 0
        if self.config.get("filter_below", True) and amt < self.config["min_amount"]:
            with self.lock:
                self.filtered_records.append(
                    {
                        "Source": rec.get("Source", ""),
                        "Date": rec.get("Date", ""),
                        "Amount (£)": amt,
                        "Details": rec.get("Details", "")[:60],
                        "Logic Used": rec.get("Logic Used", ""),
                        "Reason": f"Below minimum threshold (£{self.config['min_amount']:,.2f})",
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

        # Account filter
        if self.config.get("use_acc_filter"):
            acc_raw = re.sub(r"[\s\-]", "", self.config.get("acc_num", ""))
            # New format account numbers are "A-31105244" — check both variants
            text_stripped = re.sub(r"[\s\-]", "", text)
            if acc_raw and acc_raw not in text_stripped:
                # Also try just the numeric part
                acc_numeric = re.sub(r"\D", "", acc_raw)
                if acc_numeric not in text_stripped:
                    return False

        r_type = "Unknown"
        for label, pat in READING_PATTERNS.items():
            if pat.search(text):
                r_type = label
                break

        # Classify entry type: New Bill if it has period charges, else Ongoing Balance
        entry_type = (
            "New Bill"
            if fields.get("period_charge") or fields.get("period_from")
            else "Ongoing Balance"
        )

        self._add_record(
            {
                "Source": source_label,
                "Sender": sender,
                "Date": fields.get("date", fallback_date),
                "Period From": fields.get("period_from", "N/A"),
                "Period To": fields.get("period_to", "N/A"),
                "Invoice #": fields.get("inv_num", "N/A"),
                "Amount (£)": fields["amount"],
                "Period Charge (£)": fields.get("period_charge", "N/A"),
                "Entry Type": entry_type,
                "Reading": r_type,
                "Units (kWh)": fields.get("units_used", "N/A"),
                "Standing Chg (p/day)": fields.get("standing_charge", "N/A"),
                "Attachment Name": attachment_name or "N/A",
                "Details": (detail_label or "New invoice")[:60],
                "Logic Used": "New Invoice Format",
            }
        )
        return True

    def _process_new_credit(
        self, text, source_label, detail_label, fallback_date, sender="", attachment_name=""
    ):
        fields = extract_new_credit_fields(text)
        if "amount" not in fields:
            return False

        if self.config.get("use_acc_filter"):
            acc_raw = re.sub(r"[\s\-]", "", self.config.get("acc_num", ""))
            text_stripped = re.sub(r"[\s\-]", "", text)
            if acc_raw and acc_raw not in text_stripped:
                acc_numeric = re.sub(r"\D", "", acc_raw)
                if acc_numeric not in text_stripped:
                    return False

        self._add_record(
            {
                "Source": source_label,
                "Sender": sender,
                "Date": fields.get("date", fallback_date),
                "Period From": "N/A",
                "Period To": "N/A",
                "Invoice #": fields.get("inv_num", "N/A"),
                "Amount (£)": fields["amount"],
                "Period Charge (£)": "N/A",
                "Entry Type": "Credit",
                "Reading": "Unknown",
                "Units (kWh)": "N/A",
                "Standing Chg (p/day)": "N/A",
                "Attachment Name": attachment_name or "N/A",
                "Details": (detail_label or "Credit note")[:60],
                "Logic Used": "New Credit Note Format",
            }
        )
        return True

    # ------------------------------------------------------------------
    # Generic text processing (old format + email bodies)
    # ------------------------------------------------------------------

    def process_text(self, text, source_type, detail, fallback_date, sender="", attachment_name=""):
        if not text:
            return

        clean_text = re.sub(r"\s+", " ", text)

        # Account filter
        if self.config.get("use_acc_filter"):
            acc = re.sub(r"[\s\-]", "", self.config.get("acc_num", ""))
            if acc and acc not in re.sub(r"[\s\-]", "", clean_text):
                return

        found_amt, strategy = None, ""
        matched_pattern_idx = -1

        if self.config.get("use_anchors", True):
            for idx, p in enumerate(AMOUNT_PATTERNS):
                m = re.search(p, clean_text, re.IGNORECASE)
                if m:
                    try:
                        found_amt = float(m.group(1).replace(",", ""))
                        strategy = "Smart Context"
                        matched_pattern_idx = idx
                        break
                    except Exception:
                        continue

        if not found_amt and self.config.get("use_large", True):
            matches = re.findall(r"£\s?(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)", clean_text)
            if matches:
                floats = [float(x.replace(",", "")) for x in matches]
                highs = [x for x in floats if x >= self.config["min_amount"]]
                if highs:
                    found_amt = max(highs)
                    strategy = "Large Amount Fallback"

        if not found_amt:
            return

        # Date extraction
        date_to_use = fallback_date
        if "PDF" in source_type or "old" in source_type.lower():
            date_m = re.search(
                r"(?:Bill date|Date issued):\s*[\",]*\s*(\d{1,2}\s+\w+\s+\d{4})",
                clean_text,
                re.IGNORECASE,
            )
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
            u_m = re.search(r"([\d,]+)\s*kWh", clean_text, re.IGNORECASE)
            sc_m = re.search(r"(\d+\.\d{2})p\s*per day", clean_text, re.IGNORECASE)
            in_m = re.search(r"Invoice number[\s:,\"\'\n]*([A-Z0-9\-]+)", clean_text, re.IGNORECASE)
            if u_m:
                units_used = u_m.group(1)
            if sc_m:
                standing_charge = sc_m.group(1)
            if in_m:
                inv_num = in_m.group(1)

        period_from, period_to = self.find_billing_period(clean_text)

        # Attempt to extract period charge separately from cumulative balance
        period_charge: str | float = "N/A"
        pc_m = re.search(
            r"total charges for this (?:period|bill|invoice)\s+£\s?([\d,]+(?:\.\d{2})?)",
            clean_text,
            re.IGNORECASE,
        )
        if pc_m:
            try:
                period_charge = float(pc_m.group(1).replace(",", ""))
            except (ValueError, AttributeError):
                pass

        # Classify Entry Type based on content
        entry_type = self._classify_entry_type(
            clean_text, matched_pattern_idx, period_from, period_to, strategy
        )

        self._add_record(
            {
                "Source": source_type,
                "Sender": sender,
                "Date": date_to_use,
                "Period From": period_from,
                "Period To": period_to,
                "Invoice #": inv_num,
                "Amount (£)": found_amt,
                "Period Charge (£)": period_charge,
                "Entry Type": entry_type,
                "Reading": r_type,
                "Units (kWh)": units_used,
                "Standing Chg (p/day)": standing_charge,
                "Attachment Name": attachment_name or "N/A",
                "Details": detail[:60],
                "Logic Used": strategy,
            }
        )

    def _classify_entry_type(self, text, pattern_idx, period_from, period_to, strategy):
        """Classify a record as New Bill, Ongoing Balance, or Other based on content."""
        text_lower = text.lower()

        # If it has billing period dates AND charges/invoice details → New Bill
        has_period = period_from != "N/A" and period_to != "N/A"
        has_bill_markers = bool(
            re.search(
                r"(?:bill date|date issued|invoice number|total charges|your charges)", text_lower
            )
        )

        if has_period and has_bill_markers:
            return "New Bill"

        # Patterns 0-2 match current balance/total charges → these are new bill amounts
        # Pattern 3 matches "your new account balance" → ongoing cumulative balance
        if pattern_idx >= 0:
            if pattern_idx <= 2:
                return "New Bill"
            if pattern_idx == 3:
                return "Ongoing Balance"

        # If matched via "balance" pattern or has "account balance" language → Ongoing Balance
        if re.search(r"(?:account balance|running balance|balance brought forward)", text_lower):
            return "Ongoing Balance"

        # If matched via total/amount to pay with period info → New Bill
        if has_period:
            return "New Bill"

        # Fallback strategy check
        if strategy == "Large Amount Fallback":
            return "Other"

        # Default: if it looks like a bill (has kWh, standing charge) → New Bill
        if re.search(r"(?:kwh|standing charge|tariff)", text_lower):
            return "New Bill"

        return "Ongoing Balance"

    # ------------------------------------------------------------------
    # PDF file processing — detects format automatically
    # ------------------------------------------------------------------

    def process_pdf_file(
        self, path, source_label, detail_label, fallback_date, sender="", attachment_name=""
    ):
        if self.is_cancelled():
            return
        try:
            import io

            with open(path, "rb") as fh:
                raw = fh.read()
            pdf_hash = hashlib.sha1(raw).hexdigest()
            with self.lock:
                if pdf_hash in self.seen_pdf_hashes:
                    return
                self.seen_pdf_hashes.add(pdf_hash)

            with pdfplumber.open(io.BytesIO(raw)) as pdf:
                pdf_text = " ".join([p.extract_text() or "" for p in pdf.pages])
            del raw

            # Use original filename as attachment_name if not already set
            if not attachment_name:
                attachment_name = detail_label or ""

            fmt = detect_pdf_format(pdf_text)

            if fmt == "new_invoice":
                self._process_new_invoice(
                    pdf_text,
                    source_label,
                    detail_label,
                    fallback_date,
                    sender=sender,
                    attachment_name=attachment_name,
                )
            elif fmt == "new_credit":
                self._process_new_credit(
                    pdf_text,
                    source_label,
                    detail_label,
                    fallback_date,
                    sender=sender,
                    attachment_name=attachment_name,
                )
            else:
                self.process_text(
                    pdf_text,
                    source_label,
                    detail_label,
                    fallback_date,
                    sender=sender,
                    attachment_name=attachment_name,
                )

        except Exception as e:
            self.log_error(f"PDF: {detail_label}", str(e))

    # ------------------------------------------------------------------
    # HTM account history
    # ------------------------------------------------------------------

    def process_htm_file(self, path):
        try:
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

    # ------------------------------------------------------------------
    # PST / OST crawl
    # ------------------------------------------------------------------

    def crawl_pst(self, folder):
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
                date_str = (
                    parse_to_display_date(d_time.strftime("%Y-%m-%d")) if d_time else "Unknown"
                )

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
                                    att_name = None
                                    # Try multiple pypff methods to get the real filename
                                    getters = [
                                        lambda a=att: a.name,
                                        lambda a=att: a.get_name(),
                                        lambda a=att: a.get_long_filename(),
                                        lambda a=att: a.get_short_filename(),
                                    ]
                                    for _getter in getters:
                                        try:
                                            val = _getter()
                                            if val:
                                                att_name = val
                                                break
                                        except (AttributeError, TypeError, Exception):
                                            continue
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
        if not path or not os.path.exists(path):
            return
        pdf_files = [f for f in os.listdir(path) if f.lower().endswith(".pdf")]
        total = len(pdf_files)

        def _process_one(i_file):
            idx, fname = i_file
            if self.is_cancelled():
                return
            file_path = os.path.join(path, fname)
            fallback_date = parse_to_display_date(
                datetime.fromtimestamp(os.path.getmtime(file_path)).strftime("%Y-%m-%d")
            )
            with self.lock:
                self.pdf_count += 1
            self.process_pdf_file(
                file_path, "Local PDF Folder", fname, fallback_date, attachment_name=fname
            )
            if self.update_progress:
                self.update_progress(idx, total, f"Scanning local PDFs: {idx}/{total}")

        for item in enumerate(pdf_files, start=1):
            _process_one(item)

        self.update_ui(f"PDF folder: {self.pdf_count} PDFs processed")


# ---------------------------------------------------------------------------
# Excel helpers
# ---------------------------------------------------------------------------

THIN = Side(style="thin", color="DDDDDD")
CELL_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _hcell(ws, row, col, value, bg="FE5716"):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
    c.fill = PatternFill("solid", start_color=bg)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = CELL_BORDER
    return c


def _money(ws, r, c, val, bold=False, fill_hex=None):
    cell = ws.cell(row=r, column=c, value=val)
    cell.font = Font(name="Calibri", size=10, bold=bold)
    cell.border = CELL_BORDER
    cell.number_format = "£#,##0.00"
    cell.alignment = Alignment(horizontal="right", vertical="center")
    if fill_hex:
        cell.fill = PatternFill("solid", start_color=fill_hex)
    return cell


def _text(ws, r, c, val, bold=False, fill_hex=None, wrap=False, align="left", color="000000"):
    cell = ws.cell(row=r, column=c, value=val)
    cell.font = Font(name="Calibri", size=10, bold=bold, color=color)
    cell.border = CELL_BORDER
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if fill_hex:
        cell.fill = PatternFill("solid", start_color=fill_hex)
    return cell


def _num(ws, r, c, val, fmt="#,##0", bold=False, fill_hex=None):
    cell = ws.cell(row=r, column=c, value=val)
    cell.font = Font(name="Calibri", size=10, bold=bold)
    cell.border = CELL_BORDER
    cell.number_format = fmt
    cell.alignment = Alignment(horizontal="right", vertical="center")
    if fill_hex:
        cell.fill = PatternFill("solid", start_color=fill_hex)
    return cell


def _section_hdr(ws, r, label, ncols=3, bg="10367A"):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=r, column=c, value=label if c == 1 else "")
        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", start_color=bg)
        cell.border = CELL_BORDER
        cell.alignment = Alignment(horizontal="left", vertical="center")


# ---------------------------------------------------------------------------
# Write evidence sheet
# ---------------------------------------------------------------------------


def write_evidence_sheet(ws, df, is_duplicate=False):
    # Columns: A=Source B=Sender C=Date D=PeriodFrom E=PeriodTo F=Invoice
    #          G=Amount H=PeriodCharge I=UnitRate J=%Change K=EntryType
    #          L=Reading M=Units N=StandingChg O=AttachmentName P=Details
    #          Q=LogicUsed R=AnomalyFlag
    COL_AMOUNT = 7  # G
    COL_PERIOD_CHG = 8  # H
    COL_UNIT_RATE = 9  # I
    COL_PCT_CHANGE = 10  # J
    COL_READING_IDX = 11  # 0-based index in row for Reading (col L, position 12, but 0-based=11)
    COL_ANOMALY = 18  # R

    headers = [
        "Source",
        "Sender",
        "Date",
        "Period From",
        "Period To",
        "Invoice #",
        "Amount (£)",
        "Period Charge (£)",
        "Unit Rate (p/kWh)",
        "% Change",
        "Entry Type",
        "Reading",
        "Units (kWh)",
        "Standing Chg (p/day)",
        "Attachment Name",
        "Details",
        "Logic Used",
        "Anomaly Flag",
    ]
    bg = "888888" if is_duplicate else "FE5716"
    for col, h in enumerate(headers, 1):
        _hcell(ws, 1, col, h, bg=bg)
    ws.row_dimensions[1].height = 28

    alt_fill = PatternFill("solid", start_color="FFF3EE")

    last_data_row = len(df) + 1
    for r_idx, row in enumerate(df.values, 2):
        row_fill = alt_fill if r_idx % 2 == 0 else PatternFill()

        for c_idx, val in enumerate(row, 1):
            if c_idx == COL_PCT_CHANGE and not is_duplicate:
                # % Change as live formula — Amount is col G
                c = ws.cell(
                    row=r_idx,
                    column=COL_PCT_CHANGE,
                    value=f'=IFERROR((G{r_idx}-G{r_idx - 1})/G{r_idx - 1},"")',
                )
                c.number_format = "0.0%"
                c.alignment = Alignment(horizontal="right", vertical="top")
                c.font = Font(name="Calibri", size=10)
                c.border = CELL_BORDER
                c.fill = row_fill
            else:
                # Convert date columns to real Excel date serials (C=3, D=4, E=5)
                excel_val = val
                if c_idx in (3, 4, 5):
                    dt = to_excel_date(val)
                    if dt is not None:
                        excel_val = dt
                c = ws.cell(row=r_idx, column=c_idx, value=excel_val)
                if c_idx == COL_AMOUNT and isinstance(val, (int, float)):
                    c.number_format = "£#,##0.00"
                if c_idx == COL_PERIOD_CHG and isinstance(val, (int, float)):
                    c.number_format = "£#,##0.00"
                if c_idx == COL_UNIT_RATE and isinstance(val, (int, float)):
                    c.number_format = "0.00"
                if c_idx in (3, 4, 5) and hasattr(excel_val, "year"):
                    c.number_format = "dd/mm/yyyy"
                c.font = Font(name="Calibri", size=10)
                c.fill = (
                    row_fill if not is_duplicate else PatternFill("solid", start_color=DUP_GREY)
                )
                c.border = CELL_BORDER
                c.alignment = Alignment(vertical="top")

            # Highlight estimated readings (Reading is col L = 0-based index 11)
            if (
                not is_duplicate
                and len(row) > COL_READING_IDX
                and row[COL_READING_IDX] == "Estimated"
            ):
                c.fill = PatternFill("solid", start_color=EST_YELLOW)

        # Anomaly flag col R (18) — Amount is col G
        if not is_duplicate and r_idx > 2:
            ca = ws.cell(
                row=r_idx,
                column=COL_ANOMALY,
                value=f'=IF(AND(G{r_idx - 1}>0,G{r_idx}>G{r_idx - 1}*2),"⚠ >100% INCREASE","")',
            )
            ca.font = Font(name="Calibri", size=10, bold=True)
            ca.border = CELL_BORDER
            ca.fill = row_fill

    # Conditional formatting: only colour anomaly column red when non-empty
    if not is_duplicate and last_data_row > 2:
        ws.conditional_formatting.add(
            f"R2:R{last_data_row}",
            FormulaRule(
                formula=['$R2<>""'],
                fill=PatternFill("solid", start_color=JUMP_RED),
                font=Font(name="Calibri", size=10, bold=True),
            ),
        )

    widths = {
        "A": 18,
        "B": 26,
        "C": 13,
        "D": 13,
        "E": 13,
        "F": 16,
        "G": 13,
        "H": 15,
        "I": 15,
        "J": 10,
        "K": 14,
        "L": 11,
        "M": 12,
        "N": 18,
        "O": 28,
        "P": 38,
        "Q": 18,
        "R": 20,
    }
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width
    ws.freeze_panes = "A2"


# ---------------------------------------------------------------------------
# Write summary sheet — uses _xlfn.MAXIFS/_xlfn.MINIFS so Excel evaluates
# on load without the dynamic-array compatibility dialog
# ---------------------------------------------------------------------------


def write_summary_sheet(ws, years, evidence_sheet_name, last_data_row=5000):
    ws.title = "Annual Summary"

    headers = [
        "Year",
        "Balance Range (£)",
        "Records",
        "Avg Balance (£)",
        "Peak Balance (£)",
        "Lowest Balance (£)",
    ]
    for col, h in enumerate(headers, 1):
        _hcell(ws, 1, col, h, bg="10367A")
    ws.row_dimensions[1].height = 28

    alt_fill = PatternFill("solid", start_color="EEF2FF")
    esn = evidence_sheet_name

    date_col = f"'{esn}'!$C$2:$C${last_data_row}"
    amt_col = f"'{esn}'!$G$2:$G${last_data_row}"

    for r_idx, year_val in enumerate(years, 2):
        row_fill = alt_fill if r_idx % 2 == 0 else PatternFill()
        yr_cell = f"A{r_idx}"

        # _xlfn. prefix tells Excel to evaluate MAXIFS/MINIFS on load without
        # the dynamic-array compatibility dialog.
        peak_f = f'=IFERROR(_xlfn.MAXIFS({amt_col},{date_col},">="&DATE({yr_cell},1,1),{date_col},"<"&DATE({yr_cell}+1,1,1)),"")'
        low_f = f'=IFERROR(_xlfn.MINIFS({amt_col},{date_col},">="&DATE({yr_cell},1,1),{date_col},"<"&DATE({yr_cell}+1,1,1)),"")'
        range_f = f'=IFERROR(_xlfn.MAXIFS({amt_col},{date_col},">="&DATE({yr_cell},1,1),{date_col},"<"&DATE({yr_cell}+1,1,1))-_xlfn.MINIFS({amt_col},{date_col},">="&DATE({yr_cell},1,1),{date_col},"<"&DATE({yr_cell}+1,1,1)),"")'

        row_values = [
            int(year_val),
            range_f,
            f'=COUNTIFS({date_col},">="&DATE({yr_cell},1,1),{date_col},"<"&DATE({yr_cell}+1,1,1))',
            f'=IFERROR(AVERAGEIFS({amt_col},{date_col},">="&DATE({yr_cell},1,1),{date_col},"<"&DATE({yr_cell}+1,1,1)),"")',
            peak_f,
            low_f,
        ]
        for c_idx, val in enumerate(row_values, 1):
            c = ws.cell(row=r_idx, column=c_idx, value=val)
            c.font = Font(name="Calibri", size=10)
            c.fill = row_fill
            c.border = CELL_BORDER
            c.alignment = Alignment(
                horizontal="center" if c_idx == 1 else "right",
                vertical="top",
            )
            if c_idx == 2:
                c.number_format = "£#,##0.00"
            elif c_idx == 3:
                c.number_format = "#,##0"
            elif c_idx > 3:
                c.number_format = "£#,##0.00"

    # Grand total row — SUM/MAX/MIN over the year rows only, no dynamic-array functions
    n = len(years) + 2
    first_r = 2
    last_r = n - 1
    tot_fill = PatternFill("solid", start_color="10367A")
    tot_specs = [
        ("OVERALL", None, "center"),
        (f'=IFERROR(MAX(E{first_r}:E{last_r})-MIN(F{first_r}:F{last_r}),"")', "£#,##0.00", "right"),
        (f"=SUM(C{first_r}:C{last_r})", "#,##0", "right"),
        (f'=IFERROR(AVERAGE(D{first_r}:D{last_r}),"")', "£#,##0.00", "right"),
        (f'=IFERROR(MAX(E{first_r}:E{last_r}),"")', "£#,##0.00", "right"),
        (f'=IFERROR(MIN(F{first_r}:F{last_r}),"")', "£#,##0.00", "right"),
    ]
    for c_idx, (val, num_fmt, align) in enumerate(tot_specs, 1):
        c = ws.cell(row=n, column=c_idx, value=val)
        c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        c.fill = tot_fill
        c.border = CELL_BORDER
        c.alignment = Alignment(horizontal=align)
        if num_fmt:
            c.number_format = num_fmt

    for col_letter in ["A", "B", "C", "D", "E", "F"]:
        ws.column_dimensions[col_letter].width = 22
    ws.freeze_panes = "A2"


# ---------------------------------------------------------------------------
# Main export function
# ---------------------------------------------------------------------------


def export_to_excel(data, output_path, error_log, config, filtered=None):
    import numpy as np

    NAVY = "10367A"
    ORANGE = "FE5716"
    RED = "FF6B6B"
    AMBER = "FFD166"
    GREEN = "06D6A0"
    LGREY = "F0F0F0"
    DGREY = "888888"

    df = pd.DataFrame(data)
    df["_sort"] = df["Date"].apply(parse_to_sort_date)
    df = df.sort_values(by=["_sort", "Invoice #"], ascending=[True, False]).reset_index(drop=True)
    df["% Change"] = None

    # Deduplication — multi-pass to match the same bill across sources
    # Pass 1: Period To + Amount  (catches HTM ↔ PST where billing period matches)
    # Pass 2: Amount within 60-day window for records with no period info (Local PDF)
    dup_df = pd.DataFrame()
    if config.get("use_dedup", True):
        # Sort by source priority so the richest-metadata record is kept first
        src_pri = {
            "HTM Account History": 0,
            "PST PDF Attachment": 1,
            "Email Body": 2,
            "Local PDF Folder": 3,
        }
        df["_src_pri"] = df["Source"].map(src_pri).fillna(9).astype(int)
        df = df.sort_values(["_sort", "_src_pri"]).reset_index(drop=True)

        # Dedup key: prefer Period To (consistent across sources for same bill),
        # fall back to Date for records without period info
        df["_dedup_date"] = df["_sort"].where(
            (df["Period To"] != "N/A") & df["Period To"].notna(), df["_sort"]
        )
        is_dup = df.duplicated(subset=["_dedup_date", "Amount (£)"], keep="first")

        # Pass 2: records with no period info (e.g. Local PDF) — match by Amount
        # within a 60-day window of any already-kept record
        no_period = (df["Period To"] == "N/A") | df["Period To"].isna()
        for idx in df[~is_dup & no_period].index:
            amt = df.loc[idx, "Amount (£)"]
            rec_date = df.loc[idx, "_sort"]
            if pd.isna(rec_date):
                continue
            kept = df[(~is_dup) & (df.index != idx)]
            matches = kept[kept["Amount (£)"] == amt]
            for m_idx in matches.index:
                m_date = df.loc[m_idx, "_sort"]
                if pd.notna(m_date) and abs((rec_date - m_date).days) <= 60:
                    is_dup.at[idx] = True
                    break

        if config.get("save_dups", True):
            dup_df = df[is_dup].copy()
        df = df[~is_dup].reset_index(drop=True)
        df = df.drop(columns=["_src_pri", "_dedup_date"], errors="ignore")

    df = df.drop(columns=["_sort"], errors="ignore")
    dup_df = (
        dup_df.drop(columns=["_sort", "_src_pri", "_dedup_date"], errors="ignore")
        if not dup_df.empty
        else dup_df
    )

    # Compute Unit Rate (p/kWh) where both Period Charge and Units are available
    def _compute_unit_rate(row):
        pc = row.get("Period Charge (£)")
        units = row.get("Units (kWh)")
        try:
            pc_f = float(pc)
            u_f = float(str(units).replace(",", ""))
            if u_f > 0 and pc_f > 0:
                return round((pc_f / u_f) * 100, 2)
        except (ValueError, TypeError):
            pass
        return "N/A"

    df["Unit Rate (p/kWh)"] = df.apply(_compute_unit_rate, axis=1)
    if not dup_df.empty:
        dup_df["Unit Rate (p/kWh)"] = dup_df.apply(_compute_unit_rate, axis=1)

    col_order = [
        "Source",
        "Sender",
        "Date",
        "Period From",
        "Period To",
        "Invoice #",
        "Amount (£)",
        "Period Charge (£)",
        "Unit Rate (p/kWh)",
        "% Change",
        "Entry Type",
        "Reading",
        "Units (kWh)",
        "Standing Chg (p/day)",
        "Attachment Name",
        "Details",
        "Logic Used",
    ]
    df = df.reindex(columns=col_order)
    dup_df = dup_df.reindex(columns=col_order) if not dup_df.empty else dup_df

    # Years for summary tab
    years = sorted(
        y for y in df["Date"].apply(parse_to_sort_date).dropna().dt.year.astype(int).unique()
    )

    wb = openpyxl.Workbook()
    wb.calculation.fullCalcOnLoad = True

    # Tab 1: Evidence (created first — summary formulas reference it by name)
    ws_main = wb.active
    ws_main.title = "EDF Evidence Report"
    write_evidence_sheet(ws_main, df, is_duplicate=False)

    # Tab 2: Annual Summary
    ws_summary = wb.create_sheet(title="Annual Summary", index=0)
    write_summary_sheet(ws_summary, years, ws_main.title, last_data_row=len(df) + 1)

    # Tab 3: Duplicates
    if not dup_df.empty:
        ws_dup = wb.create_sheet(title="Duplicate Entries")
        write_evidence_sheet(ws_dup, dup_df, is_duplicate=True)

    # Tab 4: Filtered
    if filtered and config.get("save_filtered", True):
        ws_filt = wb.create_sheet(title="Filtered (Below Min)")
        filt_headers = ["Source", "Date", "Amount (£)", "Details", "Logic Used", "Reason"]
        for ci, h in enumerate(filt_headers, 1):
            _hcell(ws_filt, 1, ci, h, bg="888888")
        filt_df = pd.DataFrame(filtered).sort_values("Amount (£)", ascending=False)
        for r_idx, frow in enumerate(filt_df.values, 2):
            bg_hex = "F5F5F5" if r_idx % 2 == 0 else None
            for c_idx, val in enumerate(frow, 1):
                c = ws_filt.cell(row=r_idx, column=c_idx, value=val)
                c.font = Font(name="Calibri", size=10)
                c.border = CELL_BORDER
                if bg_hex:
                    c.fill = PatternFill("solid", start_color=bg_hex)
                if c_idx == 3:
                    c.number_format = "£#,##0.00"
        for col, w in zip(["A", "B", "C", "D", "E", "F"], [18, 13, 14, 38, 18, 28], strict=False):
            ws_filt.column_dimensions[col].width = w
        ws_filt.freeze_panes = "A2"

    # Tab 5: Parse errors
    if error_log:
        ws_err = wb.create_sheet(title="Parse Errors")
        _hcell(ws_err, 1, 1, "Time", bg="888888")
        _hcell(ws_err, 1, 2, "Context", bg="888888")
        _hcell(ws_err, 1, 3, "Error", bg="888888")
        for r_idx, entry in enumerate(error_log, 2):
            ts_m = re.match(r"\[(.+?)\]\s*(.*?)\s*—\s*(.*)", entry)
            if ts_m:
                ts, ctx, err = ts_m.group(1), ts_m.group(2), ts_m.group(3)
            else:
                ts, ctx, err = "", entry, ""
            for c_idx, val in enumerate([ts, ctx, err], 1):
                c = ws_err.cell(row=r_idx, column=c_idx, value=val)
                c.font = Font(name="Calibri", size=10)
                c.border = CELL_BORDER
        ws_err.column_dimensions["A"].width = 10
        ws_err.column_dimensions["B"].width = 45
        ws_err.column_dimensions["C"].width = 60

    # =====================================================================
    # ANALYSIS SUITE
    # Uses bills above analysis_min threshold only.
    # =====================================================================

    df_an = df.copy()
    df_an["_dt"] = df_an["Date"].apply(parse_to_sort_date)
    df_an = df_an.sort_values("_dt").reset_index(drop=True)
    analysis_min = float(config.get("analysis_min", 500.0))
    balance_types = ("New Bill", "Ongoing Balance", "Payment")
    dfc = (
        df_an[(df_an["Amount (£)"] >= analysis_min) & (df_an["Entry Type"].isin(balance_types))]
        .copy()
        .reset_index(drop=True)
    )
    dfc["year"] = dfc["_dt"].dt.year
    dfc["month"] = dfc["_dt"].dt.month

    if len(dfc) < 2:
        wb.save(output_path)
        return

    amounts = dfc["Amount (£)"].values.astype(float)
    dates_lbl = dfc["Date"].tolist()
    n = len(amounts)

    raw_diffs = np.diff(amounts)
    pos_diffs = raw_diffs[raw_diffs > 0]

    yearly = (
        dfc.groupby("year")
        .agg(
            count=("Amount (£)", "count"),
            avg_bal=("Amount (£)", "mean"),
            peak=("Amount (£)", "max"),
            low=("Amount (£)", "min"),
        )
        .reset_index()
    )

    # ----- TAB A: KEY STATISTICS -----
    ws_ks = wb.create_sheet(title="Key Statistics")
    ws_ks.column_dimensions["A"].width = 44
    ws_ks.column_dimensions["B"].width = 22
    ws_ks.column_dimensions["C"].width = 44

    tc = ws_ks.cell(row=1, column=1, value="EDF ENERGY DISPUTE  —  KEY STATISTICS")
    tc.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    tc.fill = PatternFill("solid", start_color=ORANGE)
    tc.border = CELL_BORDER
    tc.alignment = Alignment(horizontal="left", vertical="center")
    for c in [2, 3]:
        x = ws_ks.cell(row=1, column=c)
        x.fill = PatternFill("solid", start_color=ORANGE)
        x.border = CELL_BORDER
    ws_ks.row_dimensions[1].height = 26

    def ks_row(r, label, value, note="", fmt=None, bold=False, alt=False):
        bg = LGREY if alt else None
        _text(ws_ks, r, 1, label, bold=bold, fill_hex=bg)
        if fmt == "£":
            _money(ws_ks, r, 2, value, bold=bold, fill_hex=bg)
        elif fmt == "%":
            _num(ws_ks, r, 2, value, fmt="0.0%", bold=bold, fill_hex=bg)
        elif fmt == "date":
            cell = ws_ks.cell(row=r, column=2, value=value)
            cell.number_format = "dd/mm/yyyy"
            cell.font = Font(name="Calibri", size=10, bold=bold)
            cell.border = CELL_BORDER
            cell.alignment = Alignment(horizontal="right", vertical="center")
            if bg:
                cell.fill = PatternFill("solid", start_color=bg)
        elif fmt:
            _num(ws_ks, r, 2, value, fmt=fmt, bold=bold, fill_hex=bg)
        else:
            _text(ws_ks, r, 2, value, bold=bold, fill_hex=bg, align="right")
        _text(ws_ks, r, 3, note, fill_hex=bg, color=DGREY)

    acc_ref = str(config.get("report_account_ref") or config.get("acc_num") or "N/A")

    r = 2
    _section_hdr(ws_ks, r, "ACCOUNT OVERVIEW")
    r = 3
    ks_row(r, "Account reference", acc_ref, alt=True)
    r = 4
    ks_row(
        r,
        "First bill on record",
        "='Balance Trend'!A2",
        fmt="date",
        note="From Balance Trend sheet",
    )
    r = 5
    ks_row(
        r,
        "Most recent bill",
        "=INDEX('Balance Trend'!A:A,MATCH(9.99E+307,'Balance Trend'!B:B)+1)",
        fmt="date",
        alt=True,
    )
    r = 6
    ks_row(
        r,
        "Period covered (days)",
        "=IFERROR(INT(INDEX('Balance Trend'!A:A,MATCH(9.99E+307,'Balance Trend'!B:B)+1)-'Balance Trend'!A2),\"\")",
        fmt="#,##0",
        note="Days between first and last bill",
    )
    r = 7
    ks_row(
        r,
        "Total bills on record",
        "=IFERROR(COUNT('Balance Trend'!B:B),\"\")",
        fmt="#,##0",
        alt=True,
    )

    r = 8
    _section_hdr(ws_ks, r, "BALANCE FIGURES")
    r = 9
    ks_row(
        r,
        "Opening balance (first bill)",
        "='Balance Trend'!B2",
        fmt="£",
        alt=True,
        note="First entry in Balance Trend",
    )
    r = 10
    ks_row(
        r,
        "Current balance (latest bill)",
        "=INDEX('Balance Trend'!B:B,MATCH(9.99E+307,'Balance Trend'!B:B))",
        fmt="£",
        bold=True,
        note="Last numeric entry in Balance Trend",
    )
    r = 11
    ks_row(
        r,
        "Total balance increase",
        '=IFERROR(B10-B9,"")',
        fmt="£",
        bold=True,
        alt=True,
        note="Latest minus earliest",
    )
    r = 12
    ks_row(r, "% increase over full period", '=IFERROR((B10-B9)/B9,"")', fmt="%", bold=True)
    r = 13
    ks_row(
        r,
        "Mean balance across all bills",
        "=IFERROR(AVERAGE('Balance Trend'!B:B),\"\")",
        fmt="£",
        alt=True,
    )
    r = 14
    ks_row(r, "Median balance", "=IFERROR(MEDIAN('Balance Trend'!B:B),\"\")", fmt="£")
    r = 15
    ks_row(r, "Peak balance recorded", "=IFERROR(MAX('Balance Trend'!B:B),\"\")", fmt="£", alt=True)
    r = 16
    ks_row(r, "Lowest balance recorded", "=IFERROR(MIN('Balance Trend'!B:B),\"\")", fmt="£")

    r = 17
    _section_hdr(ws_ks, r, "PERIODIC CHARGES")
    r = 18
    ks_row(
        r,
        "Note",
        "Bills are a running cumulative balance — periodic charge = closing minus opening balance",
        alt=True,
    )
    r = 19
    ks_row(
        r,
        "Mean charge per period (positive only)",
        '=IFERROR(AVERAGEIF(\'Period Charges\'!F:F,">0"),"")',
        fmt="£",
    )
    r = 20
    ks_row(
        r,
        "Largest single-period charge",
        "=IFERROR(MAX('Period Charges'!F:F),\"\")",
        fmt="£",
        bold=True,
        alt=True,
    )
    r = 21
    ks_row(
        r,
        "Smallest positive charge",
        "=IFERROR(_xlfn.MINIFS('Period Charges'!F:F,'Period Charges'!F:F,\">0\"),\"\")",
        fmt="£",
    )
    r = 22
    ks_row(
        r,
        "Periods where balance increased",
        '=IFERROR(COUNTIF(\'Period Charges\'!F:F,">0"),"")',
        fmt="#,##0",
        alt=True,
    )
    r = 23
    ks_row(
        r,
        "Periods where balance fell (payments/credits)",
        '=IFERROR(COUNTIF(\'Period Charges\'!F:F,"<0"),"")',
        fmt="#,##0",
    )
    r = 24
    ks_row(
        r,
        "Implied annual rate (avg last 6 charges ×12)",
        "=IFERROR(AVERAGE(OFFSET('Period Charges'!F1,MAX(1,COUNTIF('Period Charges'!F:F,\">0\")-5),0,6,1))*12,\"\")",
        fmt="£",
        bold=True,
        alt=True,
        note="Assumes ~monthly billing — may overstate if billing is quarterly",
    )

    r = 25
    _section_hdr(ws_ks, r, "READING & DATA QUALITY")
    r = 26
    ks_row(
        r,
        "Estimated readings",
        '=IFERROR(COUNTIF(\'EDF Evidence Report\'!L:L,"Estimated"),"")',
        fmt="#,##0",
        alt=True,
    )
    r = 27
    ks_row(
        r,
        "Actual / customer readings",
        '=IFERROR(COUNTIF(\'EDF Evidence Report\'!L:L,"Actual"),"")',
        fmt="#,##0",
    )
    r = 28
    ks_row(
        r,
        "Smart meter readings",
        '=IFERROR(COUNTIF(\'EDF Evidence Report\'!L:L,"Smart"),"")',
        fmt="#,##0",
        alt=True,
    )
    r = 29
    ks_row(
        r,
        "% of bills with estimated readings",
        "=IFERROR(B26/COUNT('EDF Evidence Report'!G:G),\"\")",
        fmt="%",
    )

    r = 30
    _section_hdr(ws_ks, r, "UNIT RATES")
    r = 31
    ks_row(
        r,
        "Average unit rate (p/kWh)",
        "=IFERROR(AVERAGE('EDF Evidence Report'!I:I),\"\")",
        fmt="0.00",
        alt=True,
        note="Across all bills with valid period charge and kWh",
    )
    r = 32
    ks_row(
        r,
        "Maximum unit rate (p/kWh)",
        "=IFERROR(MAX('EDF Evidence Report'!I:I),\"\")",
        fmt="0.00",
        note="Highest effective rate — potential overcharge",
    )
    r = 33
    ks_row(
        r,
        "Minimum unit rate (p/kWh)",
        "=IFERROR(MIN('EDF Evidence Report'!I:I),\"\")",
        fmt="0.00",
        alt=True,
    )

    ws_ks.freeze_panes = "A2"

    # ----- TAB B: BALANCE TREND -----
    ws_bt = wb.create_sheet(title="Balance Trend")
    for ci, h in enumerate(
        ["Date", "Balance (£)", "6-Bill Rolling Avg (£)", "Linear Trend (£)", "Period Charge (£)"],
        1,
    ):
        _hcell(ws_bt, 1, ci, h, bg=NAVY)
    ws_bt.row_dimensions[1].height = 22

    last_data_row = n + 1
    for i in range(n):
        r = i + 2
        bg = LGREY if i % 2 == 0 else None

        # Write date as a true Excel date serial
        excel_dt = to_excel_date(dates_lbl[i])
        c1 = ws_bt.cell(row=r, column=1, value=excel_dt)
        c1.number_format = "dd/mm/yyyy"
        c1.font = Font(name="Calibri", size=10)
        c1.border = CELL_BORDER
        c1.alignment = Alignment(horizontal="left")
        if bg:
            c1.fill = PatternFill("solid", start_color=bg)

        _money(ws_bt, r, 2, float(amounts[i]), fill_hex=bg)

        start_r = max(2, r - 5)
        for col_i, formula in [
            (3, f'=IFERROR(AVERAGE(B{start_r}:B{r}),"")'),
            (
                4,
                f'=IFERROR(FORECAST.LINEAR(ROW(),B$2:B${last_data_row},ROW(B$2:B${last_data_row})),"")',
            ),
        ]:
            cx = ws_bt.cell(row=r, column=col_i, value=formula)
            cx.number_format = "£#,##0.00"
            cx.font = Font(name="Calibri", size=10)
            cx.border = CELL_BORDER
            cx.alignment = Alignment(horizontal="right")
            if bg:
                cx.fill = PatternFill("solid", start_color=bg)

        if i > 0:
            c5 = ws_bt.cell(row=r, column=5, value=f"=B{r}-B{r - 1}")
            c5.number_format = "£#,##0.00"
            c5.font = Font(name="Calibri", size=10)
            c5.border = CELL_BORDER
            c5.alignment = Alignment(horizontal="right")
            if bg:
                c5.fill = PatternFill("solid", start_color=bg)

    # Line chart
    lc = LineChart()
    lc.title = "Account Balance Over Time"
    lc.style = 10
    lc.y_axis.title = "Balance (£)"
    lc.x_axis.title = "Bill Date"
    lc.width, lc.height = 30, 18
    data_ref = Reference(ws_bt, min_col=2, max_col=4, min_row=1, max_row=n + 1)
    dates_ref = Reference(ws_bt, min_col=1, min_row=2, max_row=n + 1)
    lc.add_data(data_ref, titles_from_data=True)
    lc.set_categories(dates_ref)
    lc.series[0].graphicalProperties.line.solidFill = ORANGE
    lc.series[0].graphicalProperties.line.width = 22000
    if len(lc.series) > 1:
        lc.series[1].graphicalProperties.line.solidFill = NAVY
        lc.series[1].graphicalProperties.line.width = 15000
        lc.series[1].graphicalProperties.line.dashDot = "dash"
    if len(lc.series) > 2:
        lc.series[2].graphicalProperties.line.solidFill = DGREY
        lc.series[2].graphicalProperties.line.width = 10000
        lc.series[2].graphicalProperties.line.dashDot = "sysDash"
    ws_bt.add_chart(lc, "G2")
    for col, w in zip(["A", "B", "C", "D", "E"], [14, 16, 20, 16, 16], strict=False):
        ws_bt.column_dimensions[col].width = w
    ws_bt.freeze_panes = "A2"

    # ----- TAB C: YEAR-ON-YEAR -----
    ws_yoy = wb.create_sheet(title="Year-on-Year")
    for ci, h in enumerate(
        [
            "Year",
            "Bills",
            "Peak Balance (£)",
            "Avg Balance (£)",
            "Lowest Balance (£)",
            "YoY Avg Δ (£)",
            "YoY Avg Δ (%)",
            "Est. Readings",
            "Biggest Jump (£)",
        ],
        1,
    ):
        _hcell(ws_yoy, 1, ci, h, bg=ORANGE)
    ws_yoy.row_dimensions[1].height = 22

    prev_avg = None
    yoy_data = []
    for r_off, row_y in enumerate(yearly.itertuples(), 2):
        yr = row_y.year
        cnt = row_y.count
        pk = row_y.peak
        av = row_y.avg_bal
        lo = row_y.low
        yoy_chg_pct = ((av - prev_avg) / prev_avg) if prev_avg else None

        yr_rows = dfc[dfc["year"] == yr]
        yr_idx = yr_rows.index.tolist()
        max_jump = None
        for ii in yr_idx:
            if ii > 0 and ii in dfc.index and ii - 1 in dfc.index:
                jmp = dfc.at[ii, "Amount (£)"] - dfc.at[ii - 1, "Amount (£)"]
                if max_jump is None or jmp > max_jump:
                    max_jump = jmp

        alt = r_off % 2 == 0
        bg = LGREY if alt else None

        _num(ws_yoy, r_off, 1, yr, fmt="#,##0", fill_hex=bg, bold=True)
        _num(ws_yoy, r_off, 2, cnt, fmt="#,##0", fill_hex=bg)
        _money(ws_yoy, r_off, 3, pk, fill_hex=bg, bold=True)
        _money(ws_yoy, r_off, 4, av, fill_hex=bg)
        _money(ws_yoy, r_off, 5, lo, fill_hex=bg)

        if r_off > 2:
            c6 = ws_yoy.cell(row=r_off, column=6, value=f"=D{r_off}-D{r_off - 1}")
            c6.number_format = "£#,##0.00"
            c6.font = Font(name="Calibri", size=10, bold=True)
            c6.border = CELL_BORDER
            c6.alignment = Alignment(horizontal="right")
            if bg:
                c6.fill = PatternFill("solid", start_color=bg)

            c7 = ws_yoy.cell(row=r_off, column=7, value=f'=IFERROR(F{r_off}/D{r_off - 1},"")')
            c7.number_format = "+0.0%;-0.0%;—"
            c7.font = Font(name="Calibri", size=10, bold=True)
            c7.border = CELL_BORDER
            c7.alignment = Alignment(horizontal="right")
            yoy_fill = (
                RED
                if yoy_chg_pct is not None and yoy_chg_pct > 0.5
                else (
                    AMBER
                    if yoy_chg_pct is not None and yoy_chg_pct > 0.2
                    else (GREEN if yoy_chg_pct is not None and yoy_chg_pct < -0.1 else bg)
                )
            )
            if yoy_fill:
                c7.fill = PatternFill("solid", start_color=yoy_fill)
        else:
            ws_yoy.cell(row=r_off, column=6, value="—").border = CELL_BORDER
            ws_yoy.cell(row=r_off, column=7, value="—").border = CELL_BORDER

        yr_est = (
            int((dfc[dfc["year"] == yr]["Reading"] == "Estimated").sum())
            if "Reading" in dfc.columns
            else 0
        )
        _num(ws_yoy, r_off, 8, yr_est, fmt="#,##0", fill_hex=bg)
        if max_jump is not None:
            _money(ws_yoy, r_off, 9, max_jump, fill_hex=(RED if max_jump > 5000 else bg))

        yoy_data.append((yr, av))
        prev_avg = av

    bc = BarChart()
    bc.type = "col"
    bc.title = "Average Balance by Year"
    bc.y_axis.title = "Average Balance (£)"
    bc.style = 10
    bc.width, bc.height = 22, 14
    n_yrs = len(yoy_data)
    avg_ref = Reference(ws_yoy, min_col=4, min_row=1, max_row=n_yrs + 1)
    yr_ref = Reference(ws_yoy, min_col=1, min_row=2, max_row=n_yrs + 1)
    bc.add_data(avg_ref, titles_from_data=True)
    bc.set_categories(yr_ref)
    bc.series[0].graphicalProperties.solidFill = ORANGE
    ws_yoy.add_chart(bc, "K2")
    for col, w in zip(
        ["A", "B", "C", "D", "E", "F", "G", "H", "I"],
        [8, 8, 18, 18, 18, 16, 14, 14, 18],
        strict=False,
    ):
        ws_yoy.column_dimensions[col].width = w
    ws_yoy.freeze_panes = "A2"

    # ----- TAB D: PERIOD CHARGES -----
    ws_pc = wb.create_sheet(title="Period Charges")
    for ci, h in enumerate(
        [
            "From Date",
            "To Date",
            "Days",
            "Opening Balance (£)",
            "Closing Balance (£)",
            "Charge (£)",
            "Daily Rate (£/day)",
            "Flag",
        ],
        1,
    ):
        _hcell(ws_pc, 1, ci, h, bg=NAVY)
    ws_pc.row_dimensions[1].height = 22

    mean_daily = float(np.mean(pos_diffs)) / 30.0 if len(pos_diffs) else 0
    pc_rows_data = []

    pc_r = 2
    for i in range(1, n):
        p = dfc.iloc[i - 1]
        c_ = dfc.iloc[i]
        days = (c_["_dt"] - p["_dt"]).days
        charge = float(c_["Amount (£)"]) - float(p["Amount (£)"])
        daily = charge / days if days > 0 else None

        flag = ""
        if days > 90:
            flag = f"⚠ {days}-day gap — possible missed bill(s)"
        elif charge < 0:
            flag = f"↓ Balance reduced by £{abs(charge):,.2f} (payment or credit)"
        elif daily and mean_daily > 0 and daily > mean_daily * 2.5:
            flag = f"⚠ Daily rate {daily / mean_daily:.1f}× average"

        bg = LGREY if pc_r % 2 == 0 else None
        if flag.startswith("⚠"):
            bg = AMBER
        elif charge < 0:
            bg = GREEN

        _text(ws_pc, pc_r, 1, p["Date"], fill_hex=bg)
        _text(ws_pc, pc_r, 2, c_["Date"], fill_hex=bg)
        _num(ws_pc, pc_r, 3, days, fmt="#,##0", fill_hex=bg)
        _money(ws_pc, pc_r, 4, float(p["Amount (£)"]), fill_hex=bg)
        _money(ws_pc, pc_r, 5, float(c_["Amount (£)"]), fill_hex=bg)

        c6 = ws_pc.cell(row=pc_r, column=6, value=f"=E{pc_r}-D{pc_r}")
        c6.number_format = "£#,##0.00"
        c6.font = Font(name="Calibri", size=10)
        c6.border = CELL_BORDER
        c6.alignment = Alignment(horizontal="right")
        if bg:
            c6.fill = PatternFill("solid", start_color=bg)

        c7 = ws_pc.cell(row=pc_r, column=7, value=f'=IFERROR(F{pc_r}/C{pc_r},"")')
        c7.number_format = "£#,##0.00"
        c7.font = Font(name="Calibri", size=10)
        c7.border = CELL_BORDER
        c7.alignment = Alignment(horizontal="right")
        if bg:
            c7.fill = PatternFill("solid", start_color=bg)

        _text(ws_pc, pc_r, 8, flag, fill_hex=bg, wrap=True)

        if charge > 0:
            pc_rows_data.append((c_["Date"], charge))
        pc_r += 1

    if pc_r > 2:
        sr = pc_r + 2
        _section_hdr(ws_pc, sr, "SUMMARY STATISTICS", ncols=8, bg=ORANGE)
        sr += 1
        dr = f"F2:F{pc_r - 1}"
        cr = f"C2:C{pc_r - 1}"

        def pc_stat(r, lbl, formula, fmt="£"):
            _text(ws_pc, r, 1, lbl, bold=True, fill_hex=LGREY)
            c = ws_pc.cell(row=r, column=2, value=formula)
            c.font = Font(name="Calibri", size=10, bold=True)
            c.fill = PatternFill("solid", start_color=LGREY)
            c.border = CELL_BORDER
            c.alignment = Alignment(horizontal="right")
            c.number_format = "£#,##0.00" if fmt == "£" else fmt
            for cc in range(3, 9):
                ws_pc.cell(row=r, column=cc).fill = PatternFill("solid", start_color=LGREY)
                ws_pc.cell(row=r, column=cc).border = CELL_BORDER

        pc_stat(sr, "Mean charge per period (positive only)", f'=IFERROR(AVERAGEIF({dr},">0"),"")')
        pc_stat(sr + 1, "Largest single charge", f'=IFERROR(MAX({dr}),"")')
        pc_stat(sr + 2, "Largest credit / reduction", f'=IFERROR(MIN({dr}),"")')
        pc_stat(sr + 3, "Charge periods", f'=IFERROR(COUNTIF({dr},">0"),"")', fmt="#,##0")
        pc_stat(sr + 4, "Credit periods", f'=IFERROR(COUNTIF({dr},"<0"),"")', fmt="#,##0")
        pc_stat(sr + 5, "Average days between bills", f'=IFERROR(AVERAGE({cr}),"")', fmt="#,##0.0")

    if len(pc_rows_data) > 1:
        bc2 = BarChart()
        bc2.type = "col"
        bc2.title = "Charge Added Each Period"
        bc2.y_axis.title = "Charge (£)"
        bc2.style = 10
        bc2.width, bc2.height = 28, 14
        chg_ref2 = Reference(ws_pc, min_col=6, min_row=1, max_row=pc_r - 1)
        date_ref2 = Reference(ws_pc, min_col=2, min_row=2, max_row=pc_r - 1)
        bc2.add_data(chg_ref2, titles_from_data=True)
        bc2.set_categories(date_ref2)
        bc2.series[0].graphicalProperties.solidFill = NAVY
        ws_pc.add_chart(bc2, "J2")

    for col, w in zip(
        ["A", "B", "C", "D", "E", "F", "G", "H"], [13, 13, 7, 18, 18, 16, 14, 42], strict=False
    ):
        ws_pc.column_dimensions[col].width = w
    ws_pc.freeze_panes = "A2"

    # ----- TAB E: DISPUTE FLAGS -----
    ws_df = wb.create_sheet(title="Dispute Flags")

    def _banner(ws, r, text, bg):
        c = ws.cell(row=r, column=1, value=text)
        c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color=bg)
        c.border = CELL_BORDER
        c.alignment = Alignment(horizontal="left", vertical="center")
        for col in range(2, 7):
            x = ws.cell(row=r, column=col)
            x.fill = PatternFill("solid", start_color=bg)
            x.border = CELL_BORDER
        ws.row_dimensions[r].height = 20

    _banner(ws_df, 1, "EDF ENERGY DISPUTE  —  AUTOMATED ANALYSIS FLAGS", ORANGE)
    ws_df.cell(
        row=2,
        column=1,
        value=f"Generated {datetime.now().strftime('%d/%m/%Y %H:%M')}  |  Period: {dates_lbl[0]} to {dates_lbl[-1]}",
    )
    ws_df.cell(row=2, column=1).font = Font(name="Calibri", size=9, italic=True, color=DGREY)

    for ci, (txt, col_hex) in enumerate(
        [
            ("■ RED = HIGH severity", RED),
            ("■ AMBER = MEDIUM", AMBER),
            ("■ GREEN = Payment/credit", GREEN),
        ],
        1,
    ):
        lc2 = ws_df.cell(row=3, column=ci * 2 - 1, value=txt)
        lc2.font = Font(name="Calibri", size=9, bold=True)
        lc2.fill = PatternFill("solid", start_color=col_hex)
        lc2.border = CELL_BORDER

    hdr_row = 5
    for ci, h in enumerate(["#", "Date", "Balance (£)", "Flag Type", "Detail", "Severity"], 1):
        _hcell(ws_df, hdr_row, ci, h, bg=NAVY)

    flags = []

    for i in range(1, n):
        p = dfc.iloc[i - 1]
        c_ = dfc.iloc[i]
        chg = float(c_["Amount (£)"]) - float(p["Amount (£)"])
        pct = chg / float(p["Amount (£)"]) if float(p["Amount (£)"]) > 0 else 0
        days = (c_["_dt"] - p["_dt"]).days
        if pct > 0.25 and 0 < days <= 90:
            flags.append(
                (
                    "LARGE JUMP",
                    c_["Date"],
                    c_["Amount (£)"],
                    f"+£{chg:,.2f} (+{pct * 100:.1f}%) in {days} days (from {p['Date']}: £{p['Amount (£)']:,.2f})",
                    "HIGH" if pct > 0.5 else "MEDIUM",
                )
            )

    for i in range(1, n):
        p = dfc.iloc[i - 1]
        c_ = dfc.iloc[i]
        days = (c_["_dt"] - p["_dt"]).days
        if days > 60:
            flags.append(
                (
                    "BILLING GAP",
                    c_["Date"],
                    c_["Amount (£)"],
                    f"{days} days without a bill (previous: {p['Date']}). Balance accumulated unchecked.",
                    "HIGH" if days > 120 else "MEDIUM",
                )
            )

    if "Reading" in dfc.columns:
        run = 0
        run_start = None
        for i, rv in enumerate(dfc["Reading"].tolist()):
            if str(rv).lower() in ("estimated", "est."):
                run += 1
                if run == 1:
                    run_start = dfc.iloc[i]["Date"]
            else:
                if run >= 3:
                    flags.append(
                        (
                            "ESTIMATED RUN",
                            run_start,
                            None,
                            f"{run} consecutive estimated readings from {run_start}.",
                            "HIGH",
                        )
                    )
                run = 0
                run_start = None
        if run >= 3:
            flags.append(
                (
                    "ESTIMATED RUN",
                    run_start,
                    None,
                    f"{run} consecutive estimated readings from {run_start} (ongoing).",
                    "HIGH",
                )
            )

    if mean_daily > 0:
        for i in range(1, n):
            p = dfc.iloc[i - 1]
            c_ = dfc.iloc[i]
            days = (c_["_dt"] - p["_dt"]).days
            charge = float(c_["Amount (£)"]) - float(p["Amount (£)"])
            if days > 0 and charge > 0:
                daily = charge / days
                ratio = daily / mean_daily
                if ratio > 2.5:
                    flags.append(
                        (
                            "HIGH DAILY RATE",
                            c_["Date"],
                            c_["Amount (£)"],
                            f"£{daily:,.2f}/day ({ratio:.1f}× avg £{mean_daily:,.2f}/day) over {days} days",
                            "HIGH" if ratio > 4 else "MEDIUM",
                        )
                    )

    for i in range(1, n):
        p = dfc.iloc[i - 1]
        c_ = dfc.iloc[i]
        chg = float(c_["Amount (£)"]) - float(p["Amount (£)"])
        if chg < -500:
            flags.append(
                (
                    "BALANCE REDUCTION",
                    c_["Date"],
                    c_["Amount (£)"],
                    f"Balance fell £{abs(chg):,.2f} (from £{p['Amount (£)']:,.2f} to £{c_['Amount (£)']:,.2f}).",
                    "INFO",
                )
            )

    # Reconciliation mismatch: where consecutive New Bill records have period_charge,
    # compare balance delta vs stated period charge
    if "Period Charge (£)" in dfc.columns:
        for i in range(1, n):
            p = dfc.iloc[i - 1]
            c_ = dfc.iloc[i]
            if str(c_.get("Entry Type", "")) == "New Bill" and str(p.get("Entry Type", "")) in (
                "New Bill",
                "Ongoing Balance",
            ):
                pc = c_.get("Period Charge (£)")
                try:
                    pc_val = float(pc)
                except (ValueError, TypeError):
                    continue
                balance_delta = float(c_["Amount (£)"]) - float(p["Amount (£)"])
                diff = abs(balance_delta - pc_val)
                threshold = max(pc_val * 0.10, 50.0) if pc_val > 0 else 50.0
                if diff > threshold:
                    flags.append(
                        (
                            "RECONCILIATION MISMATCH",
                            c_["Date"],
                            c_["Amount (£)"],
                            f"Balance delta £{balance_delta:,.2f} vs period charge £{pc_val:,.2f} "
                            f"(difference: £{diff:,.2f}). Possible payment, credit, or billing error "
                            f"between {p['Date']} and {c_['Date']}.",
                            "HIGH" if diff > pc_val * 0.5 else "MEDIUM",
                        )
                    )

    sev_fill = {"HIGH": RED, "MEDIUM": AMBER, "INFO": GREEN}
    for fi, (ftype, date, amt, detail, sev) in enumerate(flags, hdr_row + 1):
        bg = sev_fill.get(sev, LGREY)
        _num(ws_df, fi, 1, fi - hdr_row, fmt="#,##0", fill_hex=bg)
        _text(ws_df, fi, 2, date or "—", fill_hex=bg)
        if amt:
            _money(ws_df, fi, 3, float(amt), fill_hex=bg)
        else:
            ws_df.cell(row=fi, column=3).fill = PatternFill("solid", start_color=bg)
            ws_df.cell(row=fi, column=3).border = CELL_BORDER
        _text(ws_df, fi, 4, ftype, bold=True, fill_hex=bg)
        _text(ws_df, fi, 5, detail, fill_hex=bg, wrap=True)
        _text(ws_df, fi, 6, sev, bold=True, fill_hex=bg, align="center")
        ws_df.row_dimensions[fi].height = 30

    if flags:
        fr = len(flags) + hdr_row + 2
        counts = {s: sum(1 for f in flags if f[4] == s) for s in ("HIGH", "MEDIUM", "INFO")}
        _banner(
            ws_df,
            fr,
            f"TOTAL FLAGS: {len(flags)}   |   HIGH: {counts['HIGH']}   |   MEDIUM: {counts['MEDIUM']}   |   INFO: {counts['INFO']}",
            NAVY,
        )

    for col, w in zip(["A", "B", "C", "D", "E", "F"], [5, 13, 16, 20, 60, 10], strict=False):
        ws_df.column_dimensions[col].width = w
    ws_df.freeze_panes = f"A{hdr_row + 1}"

    # ----- TAB F: DISPUTE TIMELINE -----
    ws_tl = wb.create_sheet(title="Dispute Timeline")
    _banner(ws_tl, 1, "EDF ENERGY DISPUTE  —  CHRONOLOGICAL TIMELINE", ORANGE)
    ws_tl.cell(
        row=2, column=1, value=f"Account: {acc_ref}  |  Period: {dates_lbl[0]} to {dates_lbl[-1]}"
    )
    ws_tl.cell(row=2, column=1).font = Font(name="Calibri", size=9, italic=True, color=DGREY)

    for ci, h in enumerate(["Date", "Event Type", "Description"], 1):
        _hcell(ws_tl, 4, ci, h, bg=NAVY)

    timeline_events = []

    # Bookend: first record
    timeline_events.append(
        (dates_lbl[0], "ACCOUNT START", f"First bill on record. Balance: £{amounts[0]:,.2f}.")
    )

    # Top 5 largest balance jumps
    jumps = []
    for i in range(1, n):
        delta = float(amounts[i]) - float(amounts[i - 1])
        days = (dfc.iloc[i]["_dt"] - dfc.iloc[i - 1]["_dt"]).days
        if delta > 0:
            jumps.append((delta, i, days))
    jumps.sort(key=lambda x: x[0], reverse=True)
    for delta, idx, days in jumps[:5]:
        timeline_events.append(
            (
                dfc.iloc[idx]["Date"],
                "LARGE INCREASE",
                f"Balance rose £{delta:,.2f} in {days} days "
                f"(from £{amounts[idx - 1]:,.2f} to £{amounts[idx]:,.2f}).",
            )
        )

    # Billing gaps > 60 days
    for i in range(1, n):
        days = (dfc.iloc[i]["_dt"] - dfc.iloc[i - 1]["_dt"]).days
        if days > 60:
            timeline_events.append(
                (
                    dfc.iloc[i]["Date"],
                    "BILLING GAP",
                    f"{days} days without a bill (previous: {dfc.iloc[i - 1]['Date']}). "
                    f"Balance accumulated unchecked.",
                )
            )

    # Estimated reading runs (reuse existing detection)
    if "Reading" in dfc.columns:
        run = 0
        run_start = None
        run_start_date = None
        for i, rv in enumerate(dfc["Reading"].tolist()):
            if str(rv).lower() in ("estimated", "est."):
                run += 1
                if run == 1:
                    run_start_date = dfc.iloc[i]["Date"]
            else:
                if run >= 3:
                    timeline_events.append(
                        (
                            run_start_date,
                            "ESTIMATED READINGS",
                            f"{run} consecutive bills used estimated meter readings.",
                        )
                    )
                run = 0
                run_start_date = None
        if run >= 3:
            timeline_events.append(
                (
                    run_start_date,
                    "ESTIMATED READINGS",
                    f"{run} consecutive estimated readings (ongoing).",
                )
            )

    # Payment events (balance reductions)
    for i in range(1, n):
        delta = float(amounts[i]) - float(amounts[i - 1])
        if delta < -200:
            timeline_events.append(
                (
                    dfc.iloc[i]["Date"],
                    "PAYMENT/CREDIT",
                    f"Balance reduced by £{abs(delta):,.2f} "
                    f"(from £{amounts[i - 1]:,.2f} to £{amounts[i]:,.2f}).",
                )
            )

    # Reconciliation mismatches (from flags)
    for ftype, fdate, _famt, fdetail, _fsev in flags:
        if ftype == "RECONCILIATION MISMATCH":
            timeline_events.append((fdate, "RECONCILIATION", fdetail))

    # Bookend: latest record
    timeline_events.append(
        (
            dates_lbl[-1],
            "CURRENT STATE",
            f"Latest bill on record. Balance: £{amounts[-1]:,.2f}. "
            f"Total increase from first record: £{amounts[-1] - amounts[0]:,.2f}.",
        )
    )

    # Sort by date and write
    timeline_events.sort(key=lambda e: parse_to_sort_date(e[0]) or pd.Timestamp.min)
    tl_r = 5
    for date, etype, desc in timeline_events:
        bg_hex = LGREY if tl_r % 2 == 0 else None
        _text(ws_tl, tl_r, 1, date, fill_hex=bg_hex)
        _text(ws_tl, tl_r, 2, etype, bold=True, fill_hex=bg_hex)
        _text(ws_tl, tl_r, 3, desc, fill_hex=bg_hex, wrap=True)
        ws_tl.row_dimensions[tl_r].height = 40
        tl_r += 1

    for col, w in zip(["A", "B", "C"], [14, 22, 90], strict=False):
        ws_tl.column_dimensions[col].width = w
    ws_tl.freeze_panes = "A5"

    # =====================================================================
    # NEW ANALYSIS TABS (added after Dispute Timeline)
    # =====================================================================

    # Statistical Analysis
    write_statistical_analysis_sheet(wb.create_sheet(title="Statistical Analysis"), dfc, config)

    # Payment Analysis
    write_payment_analysis_sheet(wb.create_sheet(title="Payment Analysis"), dfc)

    # Forecast & Projection
    write_forecast_sheet(wb.create_sheet(title="Forecast & Projection"), dfc)

    # Data Quality Report
    write_data_quality_sheet(wb.create_sheet(title="Data Quality Report"), df)

    # Tariff Analysis (if data available)
    write_tariff_analysis_sheet(wb.create_sheet(title="Tariff Analysis"), dfc)

    wb.save(output_path)


# =====================================================================
# NEW ANALYSIS FUNCTIONS (pandas-powered enhancements)
# =====================================================================


def _compute_rolling_stats(series, window=6):
    """Compute rolling statistics for a time series."""
    return {
        "mean": series.rolling(window=window, min_periods=1).mean(),
        "std": series.rolling(window=window, min_periods=1).std(),
        "min": series.rolling(window=window, min_periods=1).min(),
        "max": series.rolling(window=window, min_periods=1).max(),
        "median": series.rolling(window=window, min_periods=1).median(),
    }


def _compute_ema(series, span=6):
    """Compute Exponential Moving Average."""
    return series.ewm(span=span, adjust=False).mean()


def _compute_momentum(series, period=3):
    """Compute momentum (rate of change)."""
    return series.diff(period)


def _compute_volatility(series, window=6):
    """Compute rolling volatility (std of returns)."""
    returns = series.pct_change()
    return returns.rolling(window=window, min_periods=1).std()


def _zscore_anomalies(series, threshold=2.5):
    """Detect anomalies using z-score method."""
    if len(series) < 3:
        return pd.Series(False, index=series.index)
    mean = series.mean()
    std = series.std()
    if std == 0:
        return pd.Series(False, index=series.index)
    z_scores = np.abs((series - mean) / std)
    return z_scores > threshold


def _iqr_anomalies(series, multiplier=1.5):
    """Detect anomalies using IQR method."""
    if len(series) < 4:
        return pd.Series(False, index=series.index)
    q1 = series.quantile(0.25)
    q3 = series.quantile(0.75)
    iqr = q3 - q1
    if iqr == 0:
        return pd.Series(False, index=series.index)
    lower = q1 - multiplier * iqr
    upper = q3 + multiplier * iqr
    return (series < lower) | (series > upper)


def _linear_forecast(series, steps=6):
    """Simple linear regression forecast."""
    if len(series) < 3:
        return None
    x = np.arange(len(series))
    y = series.values
    # Handle NaN values
    mask = ~np.isnan(y)
    if mask.sum() < 3:
        return None
    x_clean = x[mask]
    y_clean = y[mask]
    try:
        coeffs = np.polyfit(x_clean, y_clean, 1)
        future_x = np.arange(len(series), len(series) + steps)
        forecast = np.polyval(coeffs, future_x)
        return forecast
    except Exception:
        return None


def _holt_winters_forecast(series, steps=6, seasonal_periods=None):
    """Holt-Winters exponential smoothing forecast (if statsmodels available)."""
    if not HAS_STATSMODELS or len(series) < 4:
        return None
    try:
        # Clean series
        clean_series = series.dropna()
        if len(clean_series) < 4:
            return None

        # Determine seasonal periods from data frequency if not provided
        if seasonal_periods is None:
            # Guess from data length
            seasonal_periods = min(12, len(clean_series) // 2) if len(clean_series) >= 8 else None

        model = ExponentialSmoothing(
            clean_series,
            trend="add",
            seasonal="add" if seasonal_periods else None,
            seasonal_periods=seasonal_periods,
            initialization_method="estimated",
        )
        fitted = model.fit(optimized=True)
        forecast = fitted.forecast(steps)
        return forecast.values
    except Exception:
        return None


def _detect_payment_patterns(df):
    """Analyze payment/credit patterns in the data."""
    payments = df[df["Entry Type"].isin(["Payment", "Credit"])].copy()
    if payments.empty:
        return {}

    payments["_dt"] = payments["Date"].apply(parse_to_sort_date)
    payments = payments.sort_values("_dt")

    # Calculate days between payments
    pay_dates = payments["_dt"].dropna()
    intervals = pay_dates.diff().dt.days.dropna()

    # Payment amounts (negative values for credits/payments)
    pay_amounts = payments["Amount (£)"].astype(float)

    return {
        "count": len(payments),
        "total_paid": abs(pay_amounts.sum()),
        "avg_payment": abs(pay_amounts.mean()),
        "median_payment": abs(pay_amounts.median()),
        "max_payment": abs(pay_amounts.max()),
        "min_payment": abs(pay_amounts.min()),
        "avg_interval_days": float(intervals.mean()) if len(intervals) > 0 else None,
        "median_interval_days": float(intervals.median()) if len(intervals) > 0 else None,
        "last_payment_date": payments.iloc[-1]["Date"] if len(payments) > 0 else None,
        "last_payment_amount": abs(pay_amounts.iloc[-1]) if len(pay_amounts) > 0 else None,
    }


def _analyze_tariff_impact(df):
    """Analyze the impact of tariff changes on unit rates and charges."""
    if "Tariff" not in df.columns or "Unit Rate (p/kWh)" not in df.columns:
        return {}

    tariff_data = df[df["Tariff"].notna() & (df["Tariff"] != "N/A")].copy()
    if tariff_data.empty:
        return {}

    # Convert unit rate to numeric
    tariff_data["unit_rate_num"] = pd.to_numeric(tariff_data["Unit Rate (p/kWh)"], errors="coerce")
    tariff_data = tariff_data.dropna(subset=["unit_rate_num"])

    if tariff_data.empty:
        return {}

    # Group by tariff
    tariff_stats = (
        tariff_data.groupby("Tariff")
        .agg(
            count=("unit_rate_num", "count"),
            avg_unit_rate=("unit_rate_num", "mean"),
            median_unit_rate=("unit_rate_num", "median"),
            min_unit_rate=("unit_rate_num", "min"),
            max_unit_rate=("unit_rate_num", "max"),
            avg_charge=("Period Charge (£)", lambda x: pd.to_numeric(x, errors="coerce").mean()),
        )
        .reset_index()
    )

    # Find tariff changes
    tariff_data = tariff_data.sort_values("_dt" if "_dt" in tariff_data.columns else "Date")
    tariff_changes = tariff_data["Tariff"].ne(tariff_data["Tariff"].shift()).cumsum()

    return {
        "tariff_stats": tariff_stats,
        "num_tariffs": tariff_data["Tariff"].nunique(),
        "tariff_changes": int(tariff_changes.max()) if not tariff_changes.empty else 0,
    }


def _data_quality_report(df):
    """Generate a comprehensive data quality report."""
    total_records = len(df)
    if total_records == 0:
        return {}

    # Date parsing success
    df["_dt_parsed"] = df["Date"].apply(parse_to_sort_date)
    date_parsed = df["_dt_parsed"].notna().sum()
    date_failed = total_records - date_parsed

    # Amount completeness
    amt_complete = df["Amount (£)"].notna().sum()
    amt_missing = total_records - amt_complete

    # Period info completeness
    period_from_complete = (df["Period From"] != "N/A").sum()
    _ = (df["Period To"] != "N/A").sum()  # Not used, but computed for completeness
    period_complete = period_from_complete  # At least from date

    # Reading classification
    reading_classified = (df["Reading"] != "Unknown").sum() if "Reading" in df.columns else 0

    # Unit rate computable
    ur_computable = (
        df["Unit Rate (p/kWh)"].apply(lambda x: isinstance(x, (int, float)) and x != "N/A")
    ).sum()

    # Duplicates (same date + amount)
    dup_count = df.duplicated(subset=["Date", "Amount (£)"]).sum()

    # Source distribution
    source_dist = df["Source"].value_counts().to_dict()

    # Entry type distribution
    entry_dist = df["Entry Type"].value_counts().to_dict() if "Entry Type" in df.columns else {}

    return {
        "total_records": total_records,
        "date_parsed": date_parsed,
        "date_failed": date_failed,
        "date_parse_rate": date_parsed / total_records if total_records > 0 else 0,
        "amt_complete": amt_complete,
        "amt_missing": amt_missing,
        "period_complete": period_complete,
        "period_completeness_rate": period_complete / total_records if total_records > 0 else 0,
        "reading_classified": reading_classified,
        "reading_classify_rate": reading_classified / total_records if total_records > 0 else 0,
        "ur_computable": ur_computable,
        "ur_computable_rate": ur_computable / total_records if total_records > 0 else 0,
        "duplicate_count": int(dup_count),
        "duplicate_rate": dup_count / total_records if total_records > 0 else 0,
        "source_distribution": source_dist,
        "entry_type_distribution": entry_dist,
    }


# ---------------------------------------------------------------------------
# NEW ANALYSIS TAB WRITERS
# ---------------------------------------------------------------------------


def write_statistical_analysis_sheet(ws, dfc, config):
    """Write Statistical Analysis tab with advanced pandas analytics."""
    ws.title = "Statistical Analysis"

    NAVY = "10367A"
    ORANGE = "FE5716"
    AMBER = "FFD166"
    LGREY = "F0F0F0"
    DGREY = "888888"

    # Prepare data
    dfc = dfc.copy()
    dfc["_dt"] = dfc["Date"].apply(parse_to_sort_date)
    dfc = dfc.sort_values("_dt").reset_index(drop=True)
    amounts = dfc["Amount (£)"].astype(float).values
    dates = dfc["Date"].tolist()
    n = len(amounts)

    if n < 3:
        _hcell(ws, 1, 1, "Insufficient data for statistical analysis", bg=NAVY)
        ws.column_dimensions["A"].width = 50
        return

    # Headers
    headers = [
        "Metric",
        "Value",
        "Notes",
    ]
    for col, h in enumerate(headers, 1):
        _hcell(ws, 1, col, h, bg=NAVY)
    ws.row_dimensions[1].height = 28

    # Title
    tc = ws.cell(row=1, column=1, value="EDF ENERGY DISPUTE  —  STATISTICAL ANALYSIS")
    tc.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    tc.fill = PatternFill("solid", start_color=ORANGE)
    tc.border = CELL_BORDER
    tc.alignment = Alignment(horizontal="left", vertical="center")
    for c in [2, 3]:
        x = ws.cell(row=1, column=c)
        x.fill = PatternFill("solid", start_color=ORANGE)
        x.border = CELL_BORDER

    # Summary stats
    r = 2
    _section_hdr(ws, r, "DESCRIPTIVE STATISTICS")

    amounts_series = pd.Series(amounts)
    stats_data = [
        ("Count", len(amounts), "#,##0", "Number of billing records"),
        ("Mean (£)", float(amounts_series.mean()), "£#,##0.00", "Average balance"),
        ("Median (£)", float(amounts_series.median()), "£#,##0.00", "Median balance"),
        ("Std Dev (£)", float(amounts_series.std()), "£#,##0.00", "Standard deviation"),
        ("Min (£)", float(amounts_series.min()), "£#,##0.00", "Minimum balance"),
        ("Max (£)", float(amounts_series.max()), "£#,##0.00", "Maximum balance"),
        ("Range (£)", float(amounts_series.max() - amounts_series.min()), "£#,##0.00", "Max - Min"),
        (
            "Skewness",
            float(amounts_series.skew()) if hasattr(amounts_series, "skew") else 0,
            "0.00",
            "Asymmetry of distribution",
        ),
        (
            "Kurtosis",
            float(amounts_series.kurtosis()) if hasattr(amounts_series, "kurtosis") else 0,
            "0.00",
            "Tailedness of distribution",
        ),
        (
            "CV (%)",
            float(amounts_series.std() / amounts_series.mean() * 100)
            if amounts_series.mean() > 0
            else 0,
            "0.00",
            "Coefficient of variation",
        ),
    ]

    for label, value, fmt, note in stats_data:
        r += 1
        bg = LGREY if r % 2 == 0 else None
        _text(ws, r, 1, label, fill_hex=bg)
        if fmt == "£":
            _money(ws, r, 2, value, fill_hex=bg)
        elif fmt == "%":
            _num(ws, r, 2, value, fmt="0.0%", fill_hex=bg)
        else:
            _num(ws, r, 2, value, fmt=fmt, fill_hex=bg)
        _text(ws, r, 3, note, fill_hex=bg, color=DGREY)

    # Rolling statistics
    r += 1
    _section_hdr(ws, r, "ROLLING STATISTICS (6-period window)")
    r += 1
    _text(ws, r, 1, "Current 6-Period Rolling Mean (£)", bold=True)
    rolling_mean = float(pd.Series(amounts).rolling(6, min_periods=1).mean().iloc[-1])
    _money(ws, r, 2, rolling_mean)

    r += 1
    _text(ws, r, 1, "Current 6-Period Rolling Std (£)", bold=True)
    rolling_std = float(pd.Series(amounts).rolling(6, min_periods=1).std().iloc[-1])
    _money(ws, r, 2, rolling_std)

    r += 1
    _text(ws, r, 1, "Current 6-Period Rolling Min (£)", bold=True)
    rolling_min = float(pd.Series(amounts).rolling(6, min_periods=1).min().iloc[-1])
    _money(ws, r, 2, rolling_min)

    r += 1
    _text(ws, r, 1, "Current 6-Period Rolling Max (£)", bold=True)
    rolling_max = float(pd.Series(amounts).rolling(6, min_periods=1).max().iloc[-1])
    _money(ws, r, 2, rolling_max)

    r += 1
    _text(ws, r, 1, "Current 6-Period Rolling Median (£)", bold=True)
    rolling_median = float(pd.Series(amounts).rolling(6, min_periods=1).median().iloc[-1])
    _money(ws, r, 2, rolling_median)

    # Exponential Moving Average
    r += 1
    _section_hdr(ws, r, "EXPONENTIAL MOVING AVERAGE")
    r += 1
    _text(ws, r, 1, "Current EMA (span=6) (£)", bold=True)
    ema = float(pd.Series(amounts).ewm(span=6, adjust=False).mean().iloc[-1])
    _money(ws, r, 2, ema)

    r += 1
    _text(ws, r, 1, "EMA vs Simple SMA Difference (£)", bold=True)
    sma = float(pd.Series(amounts).rolling(6, min_periods=1).mean().iloc[-1])
    _money(ws, r, 2, ema - sma)

    # Momentum & Volatility
    r += 1
    _section_hdr(ws, r, "MOMENTUM & VOLATILITY")
    r += 1
    mom = float(pd.Series(amounts).diff(3).iloc[-1]) if n >= 4 else 0
    _text(ws, r, 1, "3-Period Momentum (£)", bold=True)
    _money(ws, r, 2, mom)

    r += 1
    vol = (
        float(pd.Series(amounts).pct_change().rolling(6, min_periods=1).std().iloc[-1])
        if n >= 3
        else 0
    )
    _text(ws, r, 1, "6-Period Volatility (σ of returns)", bold=True)
    _num(ws, r, 2, vol, fmt="0.00%")

    # Anomaly Detection
    r += 1
    _section_hdr(ws, r, "ANOMALY DETECTION")
    series = pd.Series(amounts, index=pd.to_datetime(dates, dayfirst=True, errors="coerce"))

    z_anoms = _zscore_anomalies(series, threshold=2.5)
    iqr_anoms = _iqr_anomalies(series, multiplier=1.5)

    z_count = int(z_anoms.sum())
    iqr_count = int(iqr_anoms.sum())

    r += 1
    _text(ws, r, 1, "Z-Score Anomalies (threshold=2.5σ)", bold=True)
    _num(ws, r, 2, z_count, fmt="#,##0")

    r += 1
    _text(ws, r, 1, "IQR Anomalies (multiplier=1.5)", bold=True)
    _num(ws, r, 2, iqr_count, fmt="#,##0")

    # List detected anomalies
    if z_count > 0:
        r += 1
        _text(ws, r, 1, "Z-Score Anomaly Dates:", bold=True)
        anom_dates = series[z_anoms].index
        for dt in anom_dates:
            r += 1
            _text(
                ws,
                r,
                1,
                f"  • {dt.strftime('%d/%m/%Y') if hasattr(dt, 'strftime') else dt} ({series[dt]:,.2f})",
            )

    if iqr_count > 0:
        r += 1
        _text(ws, r, 1, "IQR Anomaly Dates:", bold=True)
        anom_dates = series[iqr_anoms].index
        for dt in anom_dates:
            r += 1
            _text(
                ws,
                r,
                1,
                f"  • {dt.strftime('%d/%m/%Y') if hasattr(dt, 'strftime') else dt} ({series[dt]:,.2f})",
            )

    # Normality test (if scipy available)
    r += 1
    _section_hdr(ws, r, "DISTRIBUTION TESTS")
    if HAS_SCIPY:
        try:
            from scipy import stats as sp_stats

            shapiro_stat, shapiro_p = sp_stats.shapiro(amounts_series.dropna())
            r += 1
            _text(ws, r, 1, "Shapiro-Wilk Test (Normality)", bold=True)
            _num(ws, r, 2, shapiro_stat, fmt="0.0000")
            _text(
                ws,
                r,
                3,
                f"p-value: {shapiro_p:.4f} — {'Normal' if shapiro_p > 0.05 else 'Non-normal'}",
            )

            # Jarque-Bera
            jb_stat, jb_p = sp_stats.jarque_bera(amounts_series.dropna())
            r += 1
            _text(ws, r, 1, "Jarque-Bera Test (Normality)", bold=True)
            _num(ws, r, 2, jb_stat, fmt="0.00")
            _text(ws, r, 3, f"p-value: {jb_p:.4f} — {'Normal' if jb_p > 0.05 else 'Non-normal'}")
        except Exception:
            r += 1
            _text(ws, r, 1, "Scipy tests failed", fill_hex=AMBER)
    else:
        r += 1
        _text(ws, r, 1, "Scipy not available — install for normality tests", fill_hex=AMBER)

    # Column widths
    for col_letter, width in zip(["A", "B", "C"], [45, 22, 80], strict=False):
        ws.column_dimensions[col_letter].width = width
    ws.freeze_panes = "A2"


def write_payment_analysis_sheet(ws, dfc):
    """Write Payment/Credit Analysis tab."""
    ws.title = "Payment Analysis"

    NAVY = "10367A"
    ORANGE = "FE5716"
    GREEN = "06D6A0"
    LGREY = "F0F0F0"
    DGREY = "888888"

    payments = dfc[dfc["Entry Type"].isin(["Payment", "Credit"])].copy()
    if payments.empty:
        _hcell(ws, 1, 1, "No payment/credit records found", bg=NAVY)
        ws.column_dimensions["A"].width = 50
        return

    payments["_dt"] = payments["Date"].apply(parse_to_sort_date)
    payments = payments.sort_values("_dt").reset_index(drop=True)

    headers = ["Metric", "Value", "Notes"]
    for col, h in enumerate(headers, 1):
        _hcell(ws, 1, col, h, bg=NAVY)
    ws.row_dimensions[1].height = 28

    tc = ws.cell(row=1, column=1, value="EDF ENERGY DISPUTE  —  PAYMENT & CREDIT ANALYSIS")
    tc.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    tc.fill = PatternFill("solid", start_color=ORANGE)
    tc.border = CELL_BORDER
    tc.alignment = Alignment(horizontal="left", vertical="center")
    for c in [2, 3]:
        x = ws.cell(row=1, column=c)
        x.fill = PatternFill("solid", start_color=ORANGE)
        x.border = CELL_BORDER

    pat = _detect_payment_patterns(dfc)

    r = 2
    _section_hdr(ws, r, "PAYMENT SUMMARY")

    payment_items = [
        ("Total Payments/Credits", pat["count"], "#,##0", "Number of payment events"),
        ("Total Amount Paid (£)", pat["total_paid"], "£#,##0.00", "Sum of all payments/credits"),
        ("Average Payment (£)", pat["avg_payment"], "£#,##0.00", "Mean payment amount"),
        ("Median Payment (£)", pat["median_payment"], "£#,##0.00", "Median payment amount"),
        ("Largest Payment (£)", pat["max_payment"], "£#,##0.00", "Maximum single payment"),
        ("Smallest Payment (£)", pat["min_payment"], "£#,##0.00", "Minimum single payment"),
    ]

    for label, value, fmt, note in payment_items:
        r += 1
        bg = LGREY if r % 2 == 0 else None
        _text(ws, r, 1, label, fill_hex=bg)
        if fmt == "£":
            _money(ws, r, 2, value, fill_hex=bg)
        else:
            _num(ws, r, 2, value, fmt=fmt, fill_hex=bg)
        _text(ws, r, 3, note, fill_hex=bg, color=DGREY)

    # Payment intervals
    r += 1
    _section_hdr(ws, r, "PAYMENT TIMING")
    interval_items = [
        ("Avg Interval (days)", pat["avg_interval_days"], "#,##0.0", "Mean days between payments"),
        (
            "Median Interval (days)",
            pat["median_interval_days"],
            "#,##0.0",
            "Median days between payments",
        ),
    ]
    for label, value, fmt, note in interval_items:
        r += 1
        bg = LGREY if r % 2 == 0 else None
        _text(ws, r, 1, label, fill_hex=bg)
        if value is not None:
            _num(ws, r, 2, value, fmt=fmt, fill_hex=bg)
        else:
            _text(ws, r, 2, "N/A", fill_hex=bg)
        _text(ws, r, 3, note, fill_hex=bg, color=DGREY)

    # Last payment
    r += 1
    _section_hdr(ws, r, "LAST PAYMENT")
    r += 1
    _text(ws, r, 1, "Last Payment Date", bold=True)
    _text(ws, r, 2, pat["last_payment_date"] or "N/A")

    r += 1
    _text(ws, r, 1, "Last Payment Amount (£)", bold=True)
    _money(ws, r, 2, pat["last_payment_amount"] or 0)

    # Payment detail table
    r += 2
    _section_hdr(ws, r, "ALL PAYMENTS & CREDITS (Chronological)")
    r += 1
    pay_headers = ["Date", "Entry Type", "Amount (£)", "Balance After (£)", "Details"]
    for ci, h in enumerate(pay_headers, 1):
        _hcell(ws, r, ci, h, bg=NAVY)

    for i, (_, row) in enumerate(payments.iterrows()):
        r += 1
        bg = LGREY if i % 2 == 0 else None
        _text(ws, r, 1, row["Date"], fill_hex=bg)
        _text(ws, r, 2, row["Entry Type"], fill_hex=bg, bold=True)
        _money(ws, r, 3, float(row["Amount (£)"]), fill_hex=bg)
        _money(ws, r, 4, float(row["Amount (£)"]), fill_hex=bg)  # Balance after = current amount
        _text(ws, r, 5, str(row.get("Details", ""))[:60], fill_hex=bg, wrap=True)

    # Chart - Payment amounts over time
    if len(payments) > 1:
        bc = BarChart()
        bc.type = "col"
        bc.title = "Payment/Credit Amounts Over Time"
        bc.y_axis.title = "Amount (£)"
        bc.x_axis.title = "Payment Date"
        bc.style = 10
        bc.width, bc.height = 28, 14
        # Add data
        pay_amounts = payments["Amount (£)"].astype(float).tolist()
        for i, amt in enumerate(pay_amounts):
            ws.cell(row=r - len(payments) + i + 1, column=6, value=amt)
        chg_ref = Reference(ws, min_col=6, min_row=r - len(payments) + 1, max_row=r)
        date_ref = Reference(ws, min_col=1, min_row=r - len(payments) + 1, max_row=r)
        bc.add_data(chg_ref, titles_from_data=False)
        bc.set_categories(date_ref)
        bc.series[0].graphicalProperties.solidFill = GREEN
        ws.add_chart(bc, f"H{r + 2}")

    for col_letter, width in zip(["A", "B", "C", "D", "E"], [14, 16, 16, 16, 60], strict=False):
        ws.column_dimensions[col_letter].width = width
    ws.freeze_panes = f"A{r - len(payments)}"


def write_forecast_sheet(ws, dfc):
    """Write Forecast/Projection tab with multiple forecasting methods."""
    ws.title = "Forecast & Projection"

    NAVY = "10367A"
    ORANGE = "FE5716"
    AMBER = "FFD166"
    LGREY = "F0F0F0"
    DGREY = "888888"

    dfc = dfc.copy()
    dfc["_dt"] = dfc["Date"].apply(parse_to_sort_date)
    dfc = dfc.sort_values("_dt").reset_index(drop=True)
    amounts = dfc["Amount (£)"].astype(float).values
    dates = dfc["Date"].tolist()
    n = len(amounts)

    if n < 3:
        _hcell(ws, 1, 1, "Insufficient data for forecasting (need 3+ records)", bg=NAVY)
        ws.column_dimensions["A"].width = 60
        return

    headers = [
        "Date",
        "Actual (£)",
        "Linear Forecast (£)",
        "Holt-Winters (£)",
        "EMA Projection (£)",
        "Confidence (±£)",
    ]
    for col, h in enumerate(headers, 1):
        _hcell(ws, 1, col, h, bg=NAVY)
    ws.row_dimensions[1].height = 28

    tc = ws.cell(row=1, column=1, value="EDF ENERGY DISPUTE  —  BALANCE FORECAST")
    tc.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    tc.fill = PatternFill("solid", start_color=ORANGE)
    tc.border = CELL_BORDER
    tc.alignment = Alignment(horizontal="left", vertical="center")
    for c in range(2, 7):
        x = ws.cell(row=1, column=c)
        x.fill = PatternFill("solid", start_color=ORANGE)
        x.border = CELL_BORDER

    # Generate forecasts (6 steps ahead)
    forecast_steps = 6
    series = pd.Series(amounts, index=pd.to_datetime(dates, dayfirst=True, errors="coerce"))

    linear_fc = _linear_forecast(series, forecast_steps)
    hw_fc = _holt_winters_forecast(series, forecast_steps)
    ema_fc = _compute_ema(series, span=6).iloc[-1] if n >= 2 else amounts[-1]
    # Simple EMA projection (extend last EMA)
    ema_vals = [ema_fc] * forecast_steps

    # Historical volatility for confidence intervals
    returns = pd.Series(amounts).pct_change().dropna()
    hist_vol = returns.std() if len(returns) > 1 else 0.05

    # Write historical data + forecasts
    r = 2
    # Historical
    for i in range(n):
        bg = LGREY if i % 2 == 0 else None
        _text(ws, r, 1, dates[i], fill_hex=bg)
        _money(ws, r, 2, float(amounts[i]), fill_hex=bg)
        _text(ws, r, 3, "—", fill_hex=bg)
        _text(ws, r, 4, "—", fill_hex=bg)
        _text(ws, r, 5, "—", fill_hex=bg)
        _text(ws, r, 6, "—", fill_hex=bg)
        r += 1

    # Separator
    ws.cell(row=r, column=1, value="—" * 20).font = Font(bold=True, color=DGREY)
    r += 1

    # Forecasts
    forecast_dates = []
    last_date = parse_to_sort_date(dates[-1])
    from datetime import timedelta

    if not pd.isna(last_date):
        for i in range(1, forecast_steps + 1):
            next_date = last_date + timedelta(days=30 * i)  # Approximate monthly
            forecast_dates.append(next_date.strftime("%d/%m/%Y"))
    else:
        forecast_dates = [f"Forecast +{i + 1}" for i in range(forecast_steps)]

    for i in range(forecast_steps):
        bg = AMBER
        _text(ws, r, 1, forecast_dates[i], fill_hex=bg, bold=True)
        _text(ws, r, 2, "—", fill_hex=bg)  # No actual
        lin_val = linear_fc[i] if linear_fc is not None else "N/A"
        hw_val = hw_fc[i] if hw_fc is not None else "N/A"
        _money(ws, r, 3, lin_val, fill_hex=bg)
        _money(ws, r, 4, hw_val, fill_hex=bg)
        _money(ws, r, 5, ema_vals[i], fill_hex=bg)
        # Confidence interval ±2σ
        if linear_fc is not None:
            conf = linear_fc[i] * hist_vol * 2
            _money(ws, r, 6, conf, fill_hex=bg)
        else:
            _text(ws, r, 6, "N/A", fill_hex=bg)
        r += 1

    # Model comparison
    r += 1
    _section_hdr(ws, r, "MODEL COMPARISON")
    r += 1
    _text(ws, r, 1, "Linear Trend", bold=True)
    _text(ws, r, 2, "Simple linear regression on time index")
    r += 1
    _text(ws, r, 1, "Holt-Winters", bold=True)
    _text(
        ws, r, 2, "Exponential smoothing with trend" + (" + seasonality" if HAS_STATSMODELS else "")
    )
    r += 1
    _text(ws, r, 1, "EMA Projection", bold=True)
    _text(ws, r, 2, "Extends last Exponential Moving Average (span=6)")
    r += 1
    _text(ws, r, 1, "Historical Volatility", bold=True)
    _num(ws, r, 2, hist_vol, fmt="0.00%")
    _text(ws, r, 3, "Monthly return std used for confidence bands")

    # Accuracy metrics (in-sample)
    r += 1
    _section_hdr(ws, r, "IN-SAMPLE ACCURACY (Last 6 periods)")
    if n >= 7:
        test_series = pd.Series(amounts[:-6])
        true_vals = amounts[-6:]
        lin_hist = _linear_forecast(test_series, 6)
        if lin_hist is not None:
            mae = np.mean(np.abs(lin_hist - true_vals))
            rmse = np.sqrt(np.mean((lin_hist - true_vals) ** 2))
            mape = np.mean(np.abs((lin_hist - true_vals) / true_vals)) * 100

            r += 1
            _text(ws, r, 1, "Linear Forecast MAE (£)", bold=True)
            _money(ws, r, 2, mae)
            r += 1
            _text(ws, r, 1, "Linear Forecast RMSE (£)", bold=True)
            _money(ws, r, 2, rmse)
            r += 1
            _text(ws, r, 1, "Linear Forecast MAPE (%)", bold=True)
            _num(ws, r, 2, mape, fmt="0.00%")

    for col_letter, width in zip(
        ["A", "B", "C", "D", "E", "F"], [14, 16, 18, 18, 18, 16], strict=False
    ):
        ws.column_dimensions[col_letter].width = width
    ws.freeze_panes = "A2"


def write_data_quality_sheet(ws, df):
    """Write Data Quality Report tab."""
    ws.title = "Data Quality Report"

    NAVY = "10367A"
    ORANGE = "FE5716"
    LGREY = "F0F0F0"
    DGREY = "888888"

    def _banner(ws, r, text, bg):
        c = ws.cell(row=r, column=1, value=text)
        c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color=bg)
        c.border = CELL_BORDER
        c.alignment = Alignment(horizontal="left", vertical="center")
        for col in range(2, 6):
            x = ws.cell(row=r, column=col)
            x.fill = PatternFill("solid", start_color=bg)
            x.border = CELL_BORDER
        ws.row_dimensions[r].height = 20

    dq = _data_quality_report(df)

    if not dq:
        _hcell(ws, 1, 1, "No data to analyze", bg=NAVY)
        ws.column_dimensions["A"].width = 40
        return

    headers = ["Check", "Result", "Rate/Count", "Status"]
    for col, h in enumerate(headers, 1):
        _hcell(ws, 1, col, h, bg=NAVY)
    ws.row_dimensions[1].height = 28

    tc = ws.cell(row=1, column=1, value="EDF ENERGY DISPUTE  —  DATA QUALITY REPORT")
    tc.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    tc.fill = PatternFill("solid", start_color=ORANGE)
    tc.border = CELL_BORDER
    tc.alignment = Alignment(horizontal="left", vertical="center")
    for c in range(2, 5):
        x = ws.cell(row=1, column=c)
        x.fill = PatternFill("solid", start_color=ORANGE)
        x.border = CELL_BORDER

    def _check_row(r, check, result, rate, status, note=""):
        bg = LGREY if r % 2 == 0 else None
        _text(ws, r, 1, check, fill_hex=bg)
        _text(ws, r, 2, str(result), fill_hex=bg)
        _text(ws, r, 3, str(rate), fill_hex=bg)
        _text(ws, r, 4, status, bold=True, fill_hex=bg)
        if note:
            ws.cell(row=r, column=5, value=note).font = Font(name="Calibri", size=9, color=DGREY)

    r = 2
    _section_hdr(ws, r, "COMPLETENESS CHECKS")

    checks = [
        ("Total Records", dq["total_records"], "—", "PASS" if dq["total_records"] > 0 else "FAIL"),
        (
            "Date Parsing Success",
            dq["date_parsed"],
            f"{dq['date_parse_rate']:.1%}",
            "PASS"
            if dq["date_parse_rate"] > 0.8
            else "WARN"
            if dq["date_parse_rate"] > 0.5
            else "FAIL",
        ),
        (
            "Amount Complete",
            dq["amt_complete"],
            f"{(dq['amt_complete'] / dq['total_records']):.1%}",
            "PASS" if dq["amt_complete"] == dq["total_records"] else "WARN",
        ),
        (
            "Period Info Complete",
            dq["period_complete"],
            f"{dq['period_completeness_rate']:.1%}",
            "PASS" if dq["period_completeness_rate"] > 0.7 else "WARN",
        ),
        (
            "Reading Classified",
            dq["reading_classified"],
            f"{dq['reading_classify_rate']:.1%}",
            "PASS" if dq["reading_classify_rate"] > 0.5 else "WARN",
        ),
        (
            "Unit Rate Computable",
            dq["ur_computable"],
            f"{dq['ur_computable_rate']:.1%}",
            "PASS" if dq["ur_computable_rate"] > 0.3 else "INFO",
        ),
    ]
    for check, result, rate, status in checks:
        _check_row(r, check, result, rate, status)
        r += 1

    r += 1
    _section_hdr(ws, r, "DUPLICATION CHECKS")
    r += 1
    _check_row(
        r,
        "Duplicate Records (Date+Amount)",
        dq["duplicate_count"],
        f"{dq['duplicate_rate']:.2%}",
        "PASS"
        if dq["duplicate_rate"] < 0.05
        else "WARN"
        if dq["duplicate_rate"] < 0.15
        else "FAIL",
    )
    r += 1

    r += 1
    _section_hdr(ws, r, "SOURCE DISTRIBUTION")
    for src, cnt in dq.get("source_distribution", {}).items():
        r += 1
        _check_row(r, f"Source: {src}", cnt, f"{cnt / dq['total_records']:.1%}", "INFO")

    r += 1
    _section_hdr(ws, r, "ENTRY TYPE DISTRIBUTION")
    for etype, cnt in dq.get("entry_type_distribution", {}).items():
        r += 1
        _check_row(r, f"Type: {etype}", cnt, f"{cnt / dq['total_records']:.1%}", "INFO")

    # Summary banner
    r += 2
    total_checks = (
        len(checks)
        + 1
        + len(dq.get("source_distribution", {}))
        + len(dq.get("entry_type_distribution", {}))
    )
    pass_count = sum(1 for c in checks if c[3] == "PASS") + (
        1 if dq["duplicate_rate"] < 0.05 else 0
    )
    warn_count = sum(1 for c in checks if c[3] == "WARN") + (
        1 if 0.05 <= dq["duplicate_rate"] < 0.15 else 0
    )
    fail_count = sum(1 for c in checks if c[3] == "FAIL") + (
        1 if dq["duplicate_rate"] >= 0.15 else 0
    )

    _banner(
        ws,
        r,
        f"QUALITY SUMMARY: {total_checks} checks  |  PASS: {pass_count}  |  WARN: {warn_count}  |  FAIL: {fail_count}",
        NAVY,
    )

    for col_letter, width in zip(["A", "B", "C", "D", "E"], [40, 20, 18, 12, 60], strict=False):
        ws.column_dimensions[col_letter].width = width
    ws.freeze_panes = "A2"


def write_tariff_analysis_sheet(ws, dfc):
    """Write Tariff Impact Analysis tab."""
    ws.title = "Tariff Analysis"

    NAVY = "10367A"
    ORANGE = "FE5716"
    LGREY = "F0F0F0"

    tariff_info = _analyze_tariff_impact(dfc)

    if not tariff_info:
        _hcell(ws, 1, 1, "No tariff data available in records", bg=NAVY)
        ws.column_dimensions["A"].width = 50
        return

    headers = [
        "Tariff",
        "Records",
        "Avg Unit Rate (p/kWh)",
        "Median Unit Rate",
        "Min Rate",
        "Max Rate",
        "Avg Period Charge (£)",
    ]
    for col, h in enumerate(headers, 1):
        _hcell(ws, 1, col, h, bg=NAVY)
    ws.row_dimensions[1].height = 28

    tc = ws.cell(row=1, column=1, value="EDF ENERGY DISPUTE  —  TARIFF IMPACT ANALYSIS")
    tc.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    tc.fill = PatternFill("solid", start_color=ORANGE)
    tc.border = CELL_BORDER
    tc.alignment = Alignment(horizontal="left", vertical="center")
    for c in range(2, 8):
        x = ws.cell(row=1, column=c)
        x.fill = PatternFill("solid", start_color=ORANGE)
        x.border = CELL_BORDER

    tariff_stats = tariff_info.get("tariff_stats")
    if tariff_stats is not None and not tariff_stats.empty:
        r = 2
        for _, row in tariff_stats.iterrows():
            bg = LGREY if r % 2 == 0 else None
            _text(ws, r, 1, str(row["Tariff"]), fill_hex=bg)
            _num(ws, r, 2, int(row["count"]), fmt="#,##0", fill_hex=bg)
            _num(ws, r, 3, float(row["avg_unit_rate"]), fmt="0.00", fill_hex=bg)
            _num(ws, r, 4, float(row["median_unit_rate"]), fmt="0.00", fill_hex=bg)
            _num(ws, r, 5, float(row["min_unit_rate"]), fmt="0.00", fill_hex=bg)
            _num(ws, r, 6, float(row["max_unit_rate"]), fmt="0.00", fill_hex=bg)
            avg_chg = row["avg_charge"]
            _money(ws, r, 7, float(avg_chg) if pd.notna(avg_chg) else 0, fill_hex=bg)
            r += 1

    r += 1
    _section_hdr(ws, r, "SUMMARY")
    r += 1
    _text(ws, r, 1, "Unique Tariffs Identified")
    _num(ws, r, 2, tariff_info.get("num_tariffs", 0), fmt="#,##0")
    r += 1
    _text(ws, r, 1, "Tariff Changes Detected")
    _num(ws, r, 2, tariff_info.get("tariff_changes", 0), fmt="#,##0")

    for col_letter, width in zip(
        ["A", "B", "C", "D", "E", "F", "G"], [28, 10, 22, 18, 16, 16, 20], strict=False
    ):
        ws.column_dimensions[col_letter].width = width
    ws.freeze_panes = "A2"


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("EDF Master Evidence Collector")
        self.root.geometry("780x860")
        self.root.configure(bg=EDF_OFFWHITE)

        self.pst_path = tk.StringVar()
        self.pdf_dir = tk.StringVar()
        self.htm_path = tk.StringVar()
        self.acc_num = tk.StringVar(value="")
        self.status = tk.StringVar(value="Ready.")
        self.progress_v = tk.DoubleVar(value=0)

        self.use_anchors = tk.BooleanVar(value=True)
        self.use_large = tk.BooleanVar(value=True)
        self.use_reading_class = tk.BooleanVar(value=True)
        self.use_pdf_fields = tk.BooleanVar(value=True)
        self.use_acc_filt = tk.BooleanVar(value=False)
        self.filter_below = tk.BooleanVar(value=True)
        self.save_filtered = tk.BooleanVar(value=True)
        self.use_dedup = tk.BooleanVar(value=True)
        self.save_dups = tk.BooleanVar(value=True)
        self.use_domain_filter = tk.BooleanVar(value=True)
        self.domain_filter = tk.StringVar(value="edfenergy.com")
        self.min_amount = tk.DoubleVar(value=500.0)
        self.analysis_min = tk.DoubleVar(value=500.0)
        self.output_name = tk.StringVar(value="EDF_Dispute_Evidence.xlsx")
        self.report_account_ref = tk.StringVar(value="")

        self.cancel_event = threading.Event()
        self.build_ui()

    def build_ui(self):
        hdr = tk.Frame(self.root, bg=EDF_ORANGE, height=60)
        hdr.pack(fill=tk.X)
        tk.Label(
            hdr,
            text="EDF BILLING EVIDENCE COLLECTOR",
            bg=EDF_ORANGE,
            fg="white",
            font=("Calibri", 14, "bold"),
        ).pack(pady=15)

        container = ttk.Frame(self.root)
        container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(container, bg=EDF_OFFWHITE, highlightthickness=0)
        yscroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=yscroll.set)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        main = ttk.Frame(canvas, padding=16)
        cw = canvas.create_window((0, 0), window=main, anchor="nw")

        def _reconfig(_e=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(cw, width=canvas.winfo_width())

        main.bind("<Configure>", _reconfig)
        canvas.bind("<Configure>", _reconfig)

        # --- Section 1: Source Data ---
        s1 = ttk.LabelFrame(main, text=" 1. Source Data ", padding=10)
        s1.pack(fill=tk.X, pady=5)

        def browse_row(parent, label, var, cmd):
            r = ttk.Frame(parent)
            r.pack(fill=tk.X, pady=2)
            ttk.Label(r, text=label, width=14).pack(side=tk.LEFT)
            ttk.Entry(r, textvariable=var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            ttk.Button(r, text="Browse", command=cmd).pack(side=tk.LEFT)

        browse_row(s1, "PST/OST File:", self.pst_path, self._pick_pst)
        browse_row(s1, "PDF Folder:", self.pdf_dir, self._pick_pdf_dir)
        browse_row(
            s1,
            "HTM Export:",
            self.htm_path,
            lambda: self.htm_path.set(
                filedialog.askopenfilename(filetypes=[("HTM/HTML", "*.htm *.html")])
            ),
        )

        # --- Section 2: Extraction options ---
        s2 = ttk.LabelFrame(main, text=" 2. Search & Filter Options ", padding=10)
        s2.pack(fill=tk.X, pady=5)
        for text, var in [
            ("Smart Context Search", self.use_anchors),
            ("Large Number Fallback", self.use_large),
            ("Classify Reading Type", self.use_reading_class),
            ("Deep PDF Mine (kWh, standing charge, invoice #)", self.use_pdf_fields),
        ]:
            tk.Checkbutton(s2, text=text, variable=var, bg=EDF_OFFWHITE).pack(anchor=tk.W)

        r3 = ttk.Frame(s2)
        r3.pack(fill=tk.X, pady=4)
        tk.Checkbutton(
            r3, text="Filter by Account #:", variable=self.use_acc_filt, bg=EDF_OFFWHITE
        ).pack(side=tk.LEFT)
        ttk.Entry(r3, textvariable=self.acc_num, width=16).pack(side=tk.LEFT, padx=5)

        r3d = ttk.Frame(s2)
        r3d.pack(fill=tk.X, pady=4)
        tk.Checkbutton(
            r3d,
            text="Filter PST emails by sender domain:",
            variable=self.use_domain_filter,
            bg=EDF_OFFWHITE,
        ).pack(side=tk.LEFT)
        ttk.Entry(r3d, textvariable=self.domain_filter, width=40).pack(side=tk.LEFT, padx=5)
        ttk.Label(r3d, text="(comma-separated domains/addresses)", font=("Calibri", 8)).pack(
            side=tk.LEFT
        )

        r4 = ttk.Frame(s2)
        r4.pack(fill=tk.X, pady=2)
        chk_filt = tk.Checkbutton(
            r4, text="Filter results below minimum £:", variable=self.filter_below, bg=EDF_OFFWHITE
        )
        chk_filt.pack(side=tk.LEFT)
        ttk.Entry(r4, textvariable=self.min_amount, width=8).pack(side=tk.LEFT, padx=5)

        r4c = ttk.Frame(s2)
        r4c.pack(fill=tk.X, pady=2)
        ttk.Label(r4c, text="Analysis threshold (£):", width=24).pack(side=tk.LEFT)
        ttk.Entry(r4c, textvariable=self.analysis_min, width=8).pack(side=tk.LEFT, padx=5)

        r4d = ttk.Frame(s2)
        r4d.pack(fill=tk.X, pady=2)
        ttk.Label(r4d, text="Report account reference:", width=24).pack(side=tk.LEFT)
        ttk.Entry(r4d, textvariable=self.report_account_ref, width=20).pack(side=tk.LEFT, padx=5)

        r4e = ttk.Frame(s2)
        r4e.pack(fill=tk.X, pady=2)
        ttk.Label(r4e, text="Output filename:", width=24).pack(side=tk.LEFT)
        ttk.Entry(r4e, textvariable=self.output_name, width=30).pack(side=tk.LEFT, padx=5)

        chk_sf = tk.Checkbutton(
            s2,
            text="Save filtered-out records to worksheet",
            variable=self.save_filtered,
            bg=EDF_OFFWHITE,
        )
        chk_sf.pack(anchor=tk.W, padx=20)
        chk_filt.config(
            command=lambda: chk_sf.config(state="normal" if self.filter_below.get() else "disabled")
        )

        # --- Section 3: Deduplication ---
        s3 = ttk.LabelFrame(main, text=" 3. Deduplication ", padding=10)
        s3.pack(fill=tk.X, pady=5)
        chk_dup = tk.Checkbutton(
            s3,
            text="Filter duplicate records (same date & amount)",
            variable=self.use_dedup,
            bg=EDF_OFFWHITE,
        )
        chk_dup.pack(anchor=tk.W)
        chk_sd = tk.Checkbutton(
            s3,
            text="Save duplicates to separate worksheet",
            variable=self.save_dups,
            bg=EDF_OFFWHITE,
        )
        chk_sd.pack(anchor=tk.W, padx=20)
        chk_dup.config(
            command=lambda: chk_sd.config(state="normal" if self.use_dedup.get() else "disabled")
        )

        # --- Progress ---
        self.pb = ttk.Progressbar(main, mode="determinate", maximum=100, variable=self.progress_v)
        self.pb.pack(fill=tk.X, pady=10)
        ttk.Label(
            main, textvariable=self.status, foreground=EDF_NAVY, font=("Calibri", 11, "bold")
        ).pack()

        btns = ttk.Frame(main)
        btns.pack(fill=tk.X, pady=8)
        self.run_btn = tk.Button(
            btns,
            text="EXTRACT TO EXCEL",
            bg=EDF_ORANGE,
            fg="white",
            font=("Calibri", 12, "bold"),
            command=self.start_thread,
            relief="flat",
        )
        self.run_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8)
        self.cancel_btn = ttk.Button(btns, text="Cancel", command=self._cancel, state="disabled")
        self.cancel_btn.pack(side=tk.LEFT, padx=8)
        self.pdf_report_btn = tk.Button(
            btns,
            text="EXPORT PDF REPORT",
            bg=EDF_NAVY,
            fg="white",
            font=("Calibri", 12, "bold"),
            command=self.export_pdf_report,
            relief="flat",
            state="disabled" if not HAS_PDF_REPORT else "normal",
        )
        self.pdf_report_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8, padx=(8, 0))

    # -- Helpers --

    def _pick_pst(self):
        p = filedialog.askopenfilename(filetypes=[("Mail Stores", "*.pst *.ost")])
        if p:
            self.pst_path.set(p)

    def _pick_pdf_dir(self):
        p = filedialog.askdirectory()
        if p:
            self.pdf_dir.set(p)

    def set_status(self, text):
        def _apply():
            self.status.set(text)
            self.root.update_idletasks()

        if threading.current_thread() is threading.main_thread():
            _apply()
        else:
            self.root.after(0, _apply)

    def set_progress(self, current, total, text=None):
        pct = max(0, min(100, (current / total) * 100)) if total else 0

        def _apply():
            self.progress_v.set(pct)
            if text:
                self.status.set(text)

        if threading.current_thread() is threading.main_thread():
            _apply()
        else:
            self.root.after(0, _apply)

    def _show(self, level, title, text):
        def _s():
            if level == "info":
                messagebox.showinfo(title, text)
            elif level == "warning":
                messagebox.showwarning(title, text)
            else:
                messagebox.showerror(title, text)

        if threading.current_thread() is threading.main_thread():
            _s()
        else:
            self.root.after(0, _s)

    def _finish(self):
        self.run_btn.config(state="normal")
        self.cancel_btn.config(state="disabled")
        if hasattr(self, "pdf_report_btn"):
            self.pdf_report_btn.config(state="normal" if HAS_PDF_REPORT else "disabled")
        self.progress_v.set(0)
        self.set_status("Cancelled." if self.cancel_event.is_set() else "Ready.")
        gc.collect()

    def export_pdf_report(self):
        """Generate professional PDF report for Ombudsman submission."""
        if not HAS_PDF_REPORT:
            self._show(
                "error",
                "PDF Report Unavailable",
                "PDF report generation requires the 'reportlab' package.\n"
                "Install with: pip install reportlab",
            )
            return

        if not hasattr(self, "engine") or not self.engine or not self.engine.records:
            self._show("warning", "No Data", "No records available. Run extraction first.")
            return

        # Ask for output path
        base_dir = (
            os.path.dirname(self.pst_path.get().strip())
            if self.pst_path.get().strip()
            else self.pdf_dir.get().strip()
            if self.pdf_dir.get().strip()
            else os.path.dirname(self.htm_path.get().strip())
            if self.htm_path.get().strip()
            else os.getcwd()
        )
        default_name = "EDF_Ombudsman_Report.pdf"
        out_path = filedialog.asksaveasfilename(
            initialdir=base_dir,
            initialfile=default_name,
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
            title="Save Ombudsman PDF Report As",
        )
        if not out_path:
            return  # User cancelled

        self.set_status("Generating professional PDF report…")
        self.pdf_report_btn.config(state="disabled")
        self.run_btn.config(state="disabled")
        self.cancel_btn.config(state="disabled")

        # Run in background thread
        def _generate():
            try:
                config = {
                    "min_amount": self.min_amount.get(),
                    "analysis_min": self.analysis_min.get(),
                    "acc_num": self.acc_num.get(),
                    "report_account_ref": self.report_account_ref.get().strip(),
                }
                success, msg = generate_pdf_from_gui(
                    records=self.engine.records,
                    output_path=out_path,
                    config=config,
                    engine=self.engine,
                    filtered=self.engine.filtered_records,
                )
                if success:
                    self.root.after(0, lambda: self._show("info", "Success", msg))
                else:
                    self.root.after(0, lambda: self._show("error", "PDF Generation Failed", msg))
            except Exception as e:
                self.root.after(
                    0, lambda err=e: self._show("error", "Error", f"An error occurred:\\n\\n{err}")
                )
            finally:
                self.root.after(
                    0,
                    lambda: (
                        self.pdf_report_btn.config(
                            state="normal" if HAS_PDF_REPORT else "disabled"
                        ),
                        self.run_btn.config(state="normal"),
                        self.cancel_btn.config(state="disabled"),
                        self.set_status("Ready."),
                    ),
                )

        threading.Thread(target=_generate, daemon=True).start()

    def _cancel(self):
        self.cancel_event.set()
        self.set_status("Cancelling…")

    def start_thread(self):
        try:
            self.min_amount.get()
            self.analysis_min.get()
        except Exception:
            messagebox.showerror(
                "Error", "Minimum amount and analysis threshold must be valid numbers."
            )
            return

        has_sources = any(
            [
                self.pst_path.get().strip(),
                self.pdf_dir.get().strip(),
                self.htm_path.get().strip(),
            ]
        )
        if not has_sources:
            messagebox.showerror(
                "Error",
                "Please select at least one source:\nPST/OST file, PDF folder, or HTM export.",
            )
            return
        self.cancel_event.clear()
        self.run_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")
        self.progress_v.set(0)
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        config = {
            "use_anchors": self.use_anchors.get(),
            "use_large": self.use_large.get(),
            "use_reading_classification": self.use_reading_class.get(),
            "use_pdf_fields": self.use_pdf_fields.get(),
            "use_acc_filter": self.use_acc_filt.get(),
            "acc_num": self.acc_num.get(),
            "min_amount": self.min_amount.get(),
            "analysis_min": self.analysis_min.get(),
            "report_account_ref": self.report_account_ref.get().strip(),
            "filter_below": self.filter_below.get(),
            "save_filtered": self.save_filtered.get(),
            "use_dedup": self.use_dedup.get(),
            "save_dups": self.save_dups.get(),
            "use_domain_filter": self.use_domain_filter.get(),
            "domain_filter": self.domain_filter.get().strip(),
        }

        engine = EvidenceEngine(config, self.set_status, self.set_progress, self.cancel_event)
        self.engine = engine

        try:
            pst_path = self.pst_path.get().strip()
            if pst_path and os.path.exists(pst_path) and not self.cancel_event.is_set():
                if not HAS_PYPFF:
                    self._show("warning", "PST", "pypff not installed — PST/OST scanning skipped.")
                else:
                    self.set_status("Scanning PST/OST…")
                    # Handle pypff API differences: some versions use .file(), others .File()
                    if hasattr(pypff, "file"):
                        pff = pypff.file()
                    elif hasattr(pypff, "File"):
                        pff = pypff.File()
                    else:
                        raise AttributeError("pypff module has no 'file' or 'File' attribute")
                    pff.open(os.path.abspath(pst_path))
                    try:
                        engine.crawl_pst(pff.get_root_folder())
                    finally:
                        pff.close()

            htm_path = self.htm_path.get().strip()
            if htm_path and os.path.exists(htm_path) and not self.cancel_event.is_set():
                self.set_status("Parsing HTM account history…")
                engine.process_htm_file(htm_path)

            pdf_path = self.pdf_dir.get().strip()
            if pdf_path and os.path.exists(pdf_path) and not self.cancel_event.is_set():
                engine.crawl_local_pdfs(pdf_path)

            if self.cancel_event.is_set():
                self._show("warning", "Cancelled", "Extraction cancelled.")
                return

            if engine.records:
                self.set_status("Writing Excel report…")
                base_dir = (
                    os.path.dirname(pst_path)
                    if pst_path
                    else pdf_path
                    if pdf_path
                    else os.path.dirname(htm_path)
                    if htm_path
                    else os.getcwd()
                )
                out_name = self.output_name.get().strip() or "EDF_Dispute_Evidence.xlsx"
                if not out_name.lower().endswith(".xlsx"):
                    out_name += ".xlsx"
                out_path = os.path.join(base_dir, out_name)
                export_to_excel(
                    engine.records,
                    out_path,
                    engine.error_log,
                    config,
                    filtered=engine.filtered_records,
                )
                summary = (
                    f"Extraction complete.\n\n"
                    f"  Emails matched: {engine.email_count}\n"
                    f"  PDFs processed: {engine.pdf_count}\n"
                    f"  Records found:  {len(engine.records)}\n"
                )
                if engine.error_log:
                    summary += f"\n  Parse errors: {len(engine.error_log)} (see Parse Errors tab)"
                summary += f"\n\nSaved to:\n{out_path}"
                self._show("info", "Success", summary)
            else:
                self._show(
                    "warning",
                    "No Data",
                    "No billing amounts found.\n\nTips:\n"
                    "• Uncheck the Account Filter\n"
                    "• Lower the minimum threshold\n"
                    "• Check your source files contain EDF billing data",
                )

        except Exception:
            self._show("error", "Error", f"An error occurred:\n\n{traceback.format_exc()}")
        finally:
            self.root.after(0, self._finish)


def run_cli_extract(args: list[str]) -> None:
    """Run extraction from command line (headless mode)."""
    import argparse
    import json
    import os
    import sys

    parser = argparse.ArgumentParser(
        description="Extract EDF billing data from PST/OST, PDF folder, or HTM export",
        prog="edf-collector --extract",
    )
    parser.add_argument("--pst", help="Path to PST/OST file")
    parser.add_argument("--pdf-dir", help="Path to directory containing PDF bills")
    parser.add_argument("--htm", help="Path to HTM account history export")
    parser.add_argument("--output", "-o", required=True, help="Output Excel file path")
    parser.add_argument("--records-json", help="Also save extracted records as JSON")
    parser.add_argument("--config", "-c", help="Path to config JSON file (optional)")
    parser.add_argument("--acc-filter", help="Filter by account number (e.g., A-12345678)")
    parser.add_argument(
        "--domain-filter",
        default="edfenergy.com",
        help="Comma-separated sender domains for PST filtering",
    )
    parser.add_argument("--min-amount", type=float, default=500.0, help="Minimum amount threshold")
    parser.add_argument("--no-dedup", action="store_true", help="Disable deduplication")
    parser.add_argument("--no-anchors", action="store_true", help="Disable smart context search")
    parser.add_argument("--no-large", action="store_true", help="Disable large amount fallback")
    parser.add_argument(
        "--no-reading-class", action="store_true", help="Disable reading classification"
    )
    parser.add_argument(
        "--no-pdf-fields", action="store_true", help="Disable deep PDF field extraction"
    )
    parser.add_argument(
        "--no-filter-below", action="store_true", help="Don't filter records below minimum amount"
    )
    parsed = parser.parse_args(args)

    # Check at least one source
    if not any([parsed.pst, parsed.pdf_dir, parsed.htm]):
        sys.stderr.write("ERROR: At least one source required (--pst, --pdf-dir, or --htm)\n")
        sys.exit(1)

    # Load config from file if provided
    config = {}
    if parsed.config:
        try:
            with open(parsed.config, encoding="utf-8") as f:
                config = json.load(f)
        except Exception as e:
            sys.stderr.write(f"ERROR: Failed to load config: {e}\n")
            sys.exit(1)

    # Override with CLI args
    config.update(
        {
            "use_acc_filter": bool(parsed.acc_filter),
            "acc_num": parsed.acc_filter or "",
            "use_domain_filter": True,
            "domain_filter": parsed.domain_filter,
            "min_amount": parsed.min_amount,
            "filter_below": not parsed.no_filter_below,
            "use_dedup": not parsed.no_dedup,
            "use_anchors": not parsed.no_anchors,
            "use_large": not parsed.no_large,
            "use_reading_classification": not parsed.no_reading_class,
            "use_pdf_fields": not parsed.no_pdf_fields,
            "save_filtered": True,
            "save_dups": True,
        }
    )

    # Check PST dependency
    if parsed.pst and not HAS_PYPFF:
        sys.stderr.write(
            "ERROR: PST/OST support requires 'libpff-python'. Install with: pip install libpff-python\n"
        )
        sys.exit(1)

    engine = EvidenceEngine(config, print, None, None)

    try:
        if parsed.pst and os.path.exists(parsed.pst):
            print(f"Scanning PST/OST: {parsed.pst}")
            # Handle pypff API differences: some versions use .file(), others .File()
            if hasattr(pypff, "file"):
                pff = pypff.file()
            elif hasattr(pypff, "File"):
                pff = pypff.File()
            else:
                raise AttributeError("pypff module has no 'file' or 'File' attribute")
            pff.open(os.path.abspath(parsed.pst))
            try:
                engine.crawl_pst(pff.get_root_folder())
            finally:
                pff.close()

        if parsed.htm and os.path.exists(parsed.htm):
            print(f"Parsing HTM: {parsed.htm}")
            engine.process_htm_file(parsed.htm)

        if parsed.pdf_dir and os.path.exists(parsed.pdf_dir):
            print(f"Scanning PDF folder: {parsed.pdf_dir}")
            engine.crawl_local_pdfs(parsed.pdf_dir)

        if not engine.records:
            sys.stderr.write("WARNING: No billing records found\n")
            sys.exit(1)

        # Export to Excel
        print(f"Writing Excel report: {parsed.output}")
        export_to_excel(
            engine.records,
            parsed.output,
            engine.error_log,
            config,
            filtered=engine.filtered_records,
        )

        # Optionally save records as JSON
        if parsed.records_json:
            import datetime

            output_data = {
                "extracted_at": datetime.datetime.now().isoformat(),
                "config": config,
                "records": engine.records,
                "filtered_records": engine.filtered_records,
                "error_log": engine.error_log,
            }
            with open(parsed.records_json, "w", encoding="utf-8") as f:
                json.dump(output_data, f, indent=2, default=str)
            print(f"Records saved as JSON: {parsed.records_json}")

        print("Extraction complete!")
        print(f"  PDFs processed: {engine.pdf_count}")
        print(f"  Emails matched: {engine.email_count}")
        print(f"  Records found:  {len(engine.records)}")
        if engine.error_log:
            print(f"  Parse errors:   {len(engine.error_log)}")

    except Exception as e:
        sys.stderr.write(f"ERROR: {e}\n")
        import traceback

        traceback.print_exc()
        sys.exit(1)


def run_cli_pdf_report(args: list[str]) -> None:
    """Run PDF report generation from command line."""
    import argparse
    import json
    import sys

    parser = argparse.ArgumentParser(
        description="Generate professional PDF report for Energy Ombudsman submission",
        prog="edf-collector --pdf-report",
    )
    parser.add_argument(
        "--records",
        "-i",
        required=True,
        help="Path to extracted records JSON file (exported from GUI or script)",
    )
    parser.add_argument("--output", "-o", required=True, help="Output PDF file path")
    parser.add_argument("--config", "-c", help="Path to config JSON file (optional)")
    parser.add_argument(
        "--engine-data",
        "-e",
        help="Path to engine data pickle file (optional, for filtered records)",
    )
    parsed = parser.parse_args(args)

    try:
        with open(parsed.records, encoding="utf-8") as f:
            records = json.load(f)

        config = {}
        if parsed.config:
            with open(parsed.config, encoding="utf-8") as f:
                config = json.load(f)

        engine = None
        filtered = None
        if parsed.engine_data:
            import pickle

            with open(parsed.engine_data, "rb") as f:
                engine = pickle.load(f)
                filtered = getattr(engine, "filtered_records", None)

        success, msg = generate_pdf_from_gui(
            records=records,
            output_path=parsed.output,
            config=config,
            engine=engine,
            filtered=filtered,
        )
        if success:
            sys.stdout.write(msg + "\n")
            sys.exit(0)
        else:
            sys.stderr.write(f"ERROR: {msg}\n")
            sys.exit(1)

    except Exception as e:
        sys.stderr.write(f"ERROR: {e}\n")
        sys.exit(1)


def main() -> None:
    """Entry point for the EDF Evidence Collector GUI."""
    import sys

    if len(sys.argv) > 1:
        if sys.argv[1] in ("--pdf-report", "--report", "-r"):
            run_cli_pdf_report(sys.argv[2:])
            return
        elif sys.argv[1] in ("--extract", "-e"):
            run_cli_extract(sys.argv[2:])
            return

    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
