"""SAP-CSV-in-PDF dump parsers and reconciliation statement extractors.

Extracted from ``edf_collector.py`` as part of the modularization refactor
(Task 4 — Phase 3).  This module is the single source of truth for:

- ``detect_sap_dump`` — header-row detection of which SAP dump type
  (``contract`` / ``meter_read`` / ``financial``) a PDF contains.
- ``parse_sap_contract_history`` — converts SAP Contract-History dump rows
  into evidence-engine records.
- ``parse_sap_meter_read_history`` — converts SAP Meter-Readings dump rows.
- ``parse_sap_financial_transactions`` — converts SAP Financial-Transactions
  dump rows; the widened parser emits 26 columns per row.
- ``extract_new_invoice_fields`` — extracts KI invoice (new-style) fields.
- ``extract_new_credit_fields`` — extracts KCR credit-note (new-style) fields.
- ``detect_reconciliation_statement`` — boolean detector for EDF
  consolidation reconciliation statement PDFs.
- ``extract_reconciliation_statement_rows`` — emits one row per charge,
  reversal, late payment, and payment found in the statement.

Internal helpers (regexes + small functions) live alongside the function
that uses them so the module is self-contained (no cross-import back into
``edf_collector``).

Compat re-exports live in ``edf_collector.py`` so callers using
``from edf_collector import parse_sap_contract_history`` continue to
work; stripped by Task 7.
"""

from __future__ import annotations

import csv as _stdcsv
import io as _io
import re

from edf_bill_fetcher.helpers.date_utils import parse_to_display_date
from edf_bill_fetcher.processors.patterns import (
    _BILLING_PERIOD_RE,
    _CREDIT_NUMBER_RE,
    _CREDIT_TOTAL_RE,
    _INV_NUMBER_RE,
    _PERIOD_CHARGE_RE,
)

# ---------------------------------------------------------------------------
# SAP dump header detection regexes — used by ``detect_sap_dump``.
# Compile once at module load; the dump-type decision is a cheap presence
# test against the first 600–1500 chars of the PDF text.
# ---------------------------------------------------------------------------
_SAP_HEADER_RE = re.compile(r'"Kraken ID"\s*,\s*"SAP Account [Nn]umber"', re.IGNORECASE)
_SAP_CONTRACT_COLS = re.compile(
    r'"Contract [Tt]ariff [Cc]ode"|"Contract [Ss]tatus"|"Start [Dd]ate"|"End [Dd]ate"'
)
_SAP_METER_COLS = re.compile(
    r'"Meter [Ss]erial [Nn]umber"|"Register [Nn]umber"|"Meter [Rr]ead [Tt]ype"'
)
_SAP_FINANCIAL_COLS = re.compile(
    r'"Posting [Dd]ate"|"Document [Nn]umber"|"Clearing [Dd]ocument"|"Amount"'
)


# ---------------------------------------------------------------------------
# Invoice / credit-note field regexes — used by ``extract_new_invoice_fields``
# and ``extract_new_credit_fields``.  These stayed in ``edf_collector`` for
# years; the moved extractors need them too, so we copy the canonical
# definitions here.  Backward compat (``from edf_collector import _ACC_NUM_RE``
# etc.) is preserved by keeping the originals in ``edf_collector`` unchanged.
# ---------------------------------------------------------------------------
_ACC_NUM_RE = re.compile(r"Account number:\s*(A-\d+|\d[\d ]*\d)", re.IGNORECASE)
_DATE_ISSUED_RE = re.compile(r"Date issued:\s*(\d{1,2}\s+\w+\s+\d{4})", re.IGNORECASE)
_CURRENT_BAL_RE = re.compile(
    r"Current balance\s+£([\d,]+\.\d{2})(?:\s+(debit|credit))?",
    re.IGNORECASE,
)
_UNITS_USED_RE = re.compile(r"Electricity used\s+([\d,]+\.?\d*)\s+kWh", re.IGNORECASE)
_STANDING_CHARGE_RE = re.compile(r"Standing charge\s+\d+\s+days\s+@\s+([\d.]+)p/day", re.IGNORECASE)
_TARIFF_NAME_RE = re.compile(r"Tariff name\s+(\w[\w\s]+?)(?:Payment type|$)", re.IGNORECASE)


# ---------------------------------------------------------------------------
# Reconciliation month-name map — used by ``_recon_to_iso`` to convert
# "DD Mon YYYY" strings to ISO date.
# ---------------------------------------------------------------------------
_RECON_MONTH_MAP = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}


def detect_sap_dump(text: str) -> str | None:
    """Return ``'contract'`` / ``'meter_read'`` / ``'financial'`` / ``None``.

    Detection is header-row based: the dump's first non-empty CSV row is
    ``"Kraken ID","SAP Account Number", ...``. Robust to any filename.
    """
    if not _SAP_HEADER_RE.search(text[:600]):
        return None
    if _SAP_CONTRACT_COLS.search(text[:1500]):
        return "contract"
    if _SAP_METER_COLS.search(text[:1500]):
        return "meter_read"
    if _SAP_FINANCIAL_COLS.search(text[:1500]):
        return "financial"
    return None


def _sap_to_iso_date(s: str) -> str:
    """Convert SAP date string ("DD.MM.YYYY", "DD-MM-YYYY", or "YYYY-MM-DD") to ISO date."""
    s = s.strip()
    # SAP format: "31.12.2023"
    if "." in s and len(s) == 10:
        d, m, y = s.split(".")
        return f"{y}-{m}-{d}"
    # DD-MM-YYYY format: "26-03-2020" — dashes in non-ISO order.
    # The first 2 chars are digits forming a valid day (01-31).
    if "-" in s and len(s) == 10:
        try:
            day = int(s[0:2])
            month = int(s[3:5])
            year = int(s[6:10])
            if 1 <= day <= 31 and 1 <= month <= 12 and year > 1900:
                return f"{year:04d}-{month:02d}-{day:02d}"
        except ValueError:
            pass
        # ISO format: "2023-12-31"
        return s
    # Fallback: try parsing
    return s


def _parse_sap_csv(text: str) -> list[dict]:
    """Parse a SAP-data-dump CSV-in-PDF body into a list of dict rows keyed

    by the CSV header column names. Empty input -> empty list. Page-break
    artifacts in pdfplumber output are handled transparently by the
    standard ``csv`` quoting rule that says a quoted field may contain
    newlines.
    """
    if not text:
        return []
    reader = _stdcsv.reader(_io.StringIO(text), skipinitialspace=True)
    rows: list[list[str]] = []
    for raw_row in reader:
        rows.append(raw_row)
    while rows and not rows[0]:
        rows.pop(0)
    if not rows:
        return []
    header = rows[0]
    out: list[dict] = []
    for raw_row in rows[1:]:
        if not raw_row or all(not (c or "").strip() for c in raw_row):
            continue
        row = {header[i]: (raw_row[i] if i < len(raw_row) else "") for i in range(len(header))}
        out.append(row)
    return out


def parse_sap_contract_history(text: str, source_file: str = "") -> list[dict]:
    """Parse the Contract-and-Product-Change-History CSV-in-PDF.

    One dict row per SAP contract record. Output columns:
      Contract From, Contract To, Product Code, Product Description,
      Contract Reason, Set Up By, Notes, Cancelled Flag, Source File.
    """
    rows = _parse_sap_csv(text)
    out: list[dict] = []
    for r in rows:
        out.append(
            {
                "Contract From": _sap_to_iso_date(r.get("Start Date", "")),
                "Contract To": _sap_to_iso_date(r.get("End Date", "")),
                "Product Code": r.get("Product", ""),
                "Product Description": r.get("Product Description", ""),
                "Contract Reason": r.get("Contract Reason", ""),
                "Set Up By": r.get("Created by", ""),
                "Notes": "",
                "Cancelled Flag": r.get("Cancelled Flag", ""),
                "Source File": source_file,
            }
        )
    return out


def parse_sap_meter_read_history(text: str, source_file: str = "") -> list[dict]:
    """Parse the Meter-Read-History CSV-in-PDF.

    One dict row per read event. Read Type is derived from the
    ``Meter Read Status`` / ``Meter Read Category`` columns:
      * ``"Released by Agent"`` -> ``A`` (Actual)
      * Anything with "estimate" in the category or status -> ``E``
      * Otherwise blank (unknown)
    """
    rows = _parse_sap_csv(text)
    out: list[dict] = []
    for r in rows:
        status = (r.get("Meter Read Status") or "").strip()
        category = (r.get("Meter Read Category") or "").strip()
        if "Released" in status:
            rtype = "A"
            rsrc = (r.get("Meter Read Type") or "").strip()
        elif "estim" in (category + " " + status).lower():
            rtype = "E"
            rsrc = (r.get("Meter Read Type") or "Automatic estimation").strip()
        else:
            rtype = ""
            rsrc = (r.get("Meter Read Type") or "").strip()
        out.append(
            {
                "Scheduled Read Date": _sap_to_iso_date(r.get("Scheduled Meter Read Date", "")),
                "Meter Read Date": _sap_to_iso_date(r.get("Meter Read Date", "")),
                "Reading (kWh)": r.get("Meter Read", "N/A"),
                "Read Type": rtype,
                "Read Source": rsrc,
                "Read Status": status,
                "Meter Read Reason": (r.get("Meter Read Reason") or "").strip(),
                "Register": r.get("Register", ""),
                "Source File": source_file,
            }
        )
    return out


def parse_sap_financial_transactions(text: str, source_file: str = "") -> list[dict]:
    """Parse the Financial-Transactions CSV-in-PDF.

    One dict row per financial transaction. The real SAP header has a
    trailing-space variant ``"Posting Date "``; both that and the
    no-trailing-space form are accepted.

    The source PDF emits 32 CSV columns per row. This parser retains
    the 16 historically-surfaced columns plus 8 columns used by the
    SAP Back-billing analyser (Contract, Sub Item, Clearing Posting
    Date, Clearing Amount, Statistical Key Flag, Tax Code,
    Tax Code Description, G/L Account, G/L Description, Deferral
    Date). The remaining 8 source columns (Kraken ID, SAP Account
    Number, Business Partner, Account Determination ID, Fuel Type,
    Payment Method, Down Payment Flag, Restriction) carry values
    that are either constant per account or not consumed by any
    downstream analyser, and are intentionally dropped.
    """
    rows = _parse_sap_csv(text)
    out: list[dict] = []
    for r in rows:
        posting_raw = ""
        for k in ("Posting Date ", "Posting Date"):
            if k in r:
                posting_raw = r.get(k, "")
                break
        clearing_posting_raw = ""
        for k in ("Clearing Posting Date ", "Clearing Posting Date"):
            if k in r:
                clearing_posting_raw = r.get(k, "")
                break
        out.append(
            {
                "Document No.": r.get("Document No.", ""),
                "Item": r.get("Item", ""),
                "Document Date": _sap_to_iso_date(r.get("Document Date", "")),
                "Posting Date": _sap_to_iso_date(posting_raw.strip()),
                "Net Due Date": _sap_to_iso_date(r.get("Net Due Date", "")),
                "Main Transaction": r.get("Main Transactions", ""),
                "Sub Transaction": r.get("Sub Transactions", ""),
                "Transaction Text": r.get("Transaction Text", ""),
                "Amount": r.get("Amount", ""),
                "Clearing Status": r.get("Clearing Status", ""),
                "Clearing Document": r.get("Clearing Document", ""),
                "Clearing Date": _sap_to_iso_date(r.get("Clearing Date", "")),
                "Clearing Reason": r.get("Clearing Reason", ""),
                "Document Type": r.get("Document Type", ""),
                "Document Type Description": r.get("Document Type Description", ""),
                "Source File": source_file,
                # Analyser-relevant extensions (spec §3.1):
                "Contract": r.get("Contract", ""),
                "Sub Item": r.get("Sub Item", ""),
                "Clearing Posting Date": _sap_to_iso_date(clearing_posting_raw.strip()),
                "Clearing Amount": r.get("Clearing Amount", ""),
                "Statistical Key Flag": r.get("Statistical Key Flag", ""),
                "Tax Code": r.get("Tax Code", ""),
                "Tax Code Description": r.get("Tax Code Description", ""),
                "G/L Account": r.get("G/L Account", ""),
                "G/L Description": r.get("G/L Description", ""),
                "Deferral Date": _sap_to_iso_date(r.get("Deferral Date", "")),
            }
        )
    return out


# ---------------------------------------------------------------------------
# SAP Back-billing analysis (spec: 2026-07-21-sap-back-billing-analysis-design.md)
# ---------------------------------------------------------------------------
#
# EDF's SAP financial ledger is the *behind-the-scenes* truth of every
# billing event; the EDF-branded invoice PDFs only show what EDF chose to
# send the customer.  Two SAP-native signals reveal back-billing activity
# without any cross-system join:
#
# 1. ``Cr- Credit for Consum Billing`` rows — a credit posted against a
#    previously-raised consumption billing (18 instances in this account).
# 2. ``Clearing Document`` clusters — a SAP bookkeeping event that
#    simultaneously clears multiple prior document numbers; the canonical
#    back-billing signature is a cluster whose rows net to exactly £0.00
#    (one or more ``Dr- Consum Billing Receivable`` postings paired with
#    matching ``Cr- Credit for Consum Billing`` reversals on the same day).
#
# Joining SAP rows to EDF invoices by ``Invoice #`` is impossible (the two
# numbering schemes have zero string overlap — verified empirically), but
# a *cluster-level* join by ``Clearing Date`` ↔ EDF ``Period To`` finds
# real back-billing events.  See the spec for details.

_SAP_DEBT_MGMT_FLAG_VALUE = "Installment Plan Item"
_SAP_MIN_CLUSTER_SIZE = 4
_SAP_MATCH_DAY_BANDS = ((0, 50), (3, 25), (14, 5))
_SAP_MATCH_AMOUNT_BANDS = ((0.05, 40), (0.25, 20), (0.50, 5))
_SAP_CONFIDENCE_BANDS = (("High", 75), ("Medium", 40), ("Low", 10))


def _parse_amount_for_event(v: object) -> float:
    """Parse the SAP ``Amount`` field (often a string with commas)."""
    if v is None:
        return 0.0
    try:
        s = str(v).strip().lstrip("£").replace(",", "")
        if not s:
            return 0.0
        return float(s)
    except ValueError:
        return 0.0


def extract_new_invoice_fields(text):
    """Extract key fields from new-style KI-XXXXXXXX invoices."""
    fields = {}

    # Invoice number
    m = _INV_NUMBER_RE.search(text)
    if m:
        fields["inv_num"] = m.group(1).strip()

    # Account number. EDF has shipped multiple renderings on real
    # invoices — both
    #
    #     "Account number: A-31105244"
    #
    # and
    #
    #     "Account number: 671 078 701 920"
    #
    # show up in real exports. The compact form requires the A- prefix;
    # the spaced form omits the prefix and groups digits three-by-three.
    #
    # Pre-fix only the compact form was matched, so any invoice using
    # spaced digits silently dropped ``acc_num`` AND failed the
    # downstream --acc-filter (the user could not filter to their
    # own bill). The regex below matches both renderings and emits a
    # single normalised account number (either ``A-NNNNNNNN`` or
    # ``601 234 567 890`` depending on what was on the page). Callers
    # that need to compare against a filter value are responsible for
    # stripping spaces and the ``A-`` prefix themselves; see the
    # existing helper at the engine filter check.
    m = _ACC_NUM_RE.search(text)
    if m:
        fields["acc_num"] = m.group(1).strip()

    # Date issued
    m = _DATE_ISSUED_RE.search(text)
    if m:
        fields["date"] = parse_to_display_date(m.group(1).strip())

    # Billing period from "Your charges: DD Mon YYYY - DD Mon YYYY"
    m = _BILLING_PERIOD_RE.search(text)
    if m:
        fields["period_from"] = parse_to_display_date(m.group(1).strip())
        fields["period_to"] = parse_to_display_date(m.group(2).strip())

    # Current balance (the running account total — used as primary Amount).
    #
    # Pre-fix this hard-required "debit" — same gap as the #15 HTM parser:
    # if the KI invoice reports ``Current balance GBPX in credit`` (rare but
    # legal: e.g. over-payment or opening credit balance), this matcher
    # would drop the Amount cell. Accept either currency-side label.
    m = _CURRENT_BAL_RE.search(text)
    if m:
        fields["amount"] = float(m.group(1).replace(",", ""))
        fields["amount_side"] = (m.group(2) or "").lower()

    # Period charge (total for this invoice).
    #
    # Pre-fix, this regex hard-required "debit" after the amount — it
    # silently dropped period charges where the line was reported in
    # credit. The "debit"/"credit" labelling is a side-property of
    # the statement, not the period charge itself; accept either so a
    # credit-flagged period still populates the Period Charge column.
    # Captures: group(1) = amount, group(2) = debit|credit (may be empty).
    m = _PERIOD_CHARGE_RE.search(text)
    if m:
        fields["period_charge"] = float(m.group(1).replace(",", ""))

    # kWh used
    m = _UNITS_USED_RE.search(text)
    if m:
        fields["units_used"] = m.group(1)

    # Standing charge
    m = _STANDING_CHARGE_RE.search(text)
    if m:
        fields["standing_charge"] = m.group(1)

    # Tariff name
    m = _TARIFF_NAME_RE.search(text)
    if m:
        fields["tariff"] = m.group(1).strip()

    return fields


def extract_new_credit_fields(text):
    """Extract key fields from new-style KCR-XXXXXXXX credit notes."""
    fields = {}

    m = _CREDIT_NUMBER_RE.search(text)
    if m:
        fields["inv_num"] = m.group(1).strip()

    # Account number — accept both EDF renderings: compact
    # "A-NNNNNNNN" and spaced-digits "601 234 567 890". See the same
    # note in extract_new_invoice_fields above for context.
    m = _ACC_NUM_RE.search(text)
    if m:
        fields["acc_num"] = m.group(1).strip()

    m = _DATE_ISSUED_RE.search(text)
    if m:
        fields["date"] = parse_to_display_date(m.group(1).strip())

    # Total credits for this bill
    m = _CREDIT_TOTAL_RE.search(text)
    if m:
        fields["amount"] = float(m.group(1).replace(",", ""))

    return fields


# ---------------------------------------------------------------------------
# SAP CSV-in-PDF data dump parsers
# ---------------------------------------------------------------------------
#
# EDF exports three types of structured data dump from its SAP / Kraken
# back-end as quoted CSV records inside a PDF (one record per line, with
# the first row a quoted CSV header). These are the canonical source
# for contract history, meter-read history, and the financial ledger.
# Filename is not used as a routing signal — detection is header-row
# based so the same parser works on any filename the user supplies.


_SAP_HEADER_RE = re.compile(r'"Kraken ID"\s*,\s*"SAP Account [Nn]umber"', re.IGNORECASE)
_SAP_CONTRACT_COLS = re.compile(
    r'"Product"\s*,\s*"Product[\s\n]*Description"\s*,\s*"Contract[\s\n]*Reason"', re.I
)
_SAP_METER_COLS = re.compile(
    r'"Meter Read Reason"\s*,\s*"Scheduled[\s\n]*Meter[\s\n]*Read[\s\n]*Date"', re.I
)
_SAP_FINANCIAL_COLS = re.compile(
    r'"Main[\s\n]*Transactions"\s*,\s*"Sub[\s\n]*Transactions"\s*,\s*"Transaction[\s\n]*Text"', re.I
)
_SAP_DDMMYYYY_RE = re.compile(r"\b(\d{2})-(\d{2})-(\d{4})\b")


# ---------------------------------------------------------------------------
# Multi-regex fallback chain (Stream P3 / Task 5)
# ---------------------------------------------------------------------------
# Each fallback chain scans the input text in a fixed precedence order and
# returns ``(value, regex_name)`` so the Source Excerpt column can show the
# technical trace ("inv_num via _COVER_BLOCK_INV_RE; period via ..."). This
# reduces the N/A count on the analyser tabs (Back-billing, Rebilling,
# Meter Readings, Contract History) since many invoice PDFs sidestep the
# canonical "Invoice number: KI-<n>" / "Your charges: <from> - <to>" markers
# but still surface the data under alternative phrasings on the cover sheet.


# ---------------------------------------------------------------------------
# Reconciliation statement regexes — used by ``detect_reconciliation_statement``
# and ``extract_reconciliation_statement_rows``.
# ---------------------------------------------------------------------------
_RECON_STATEMENT_RE = re.compile(
    r"Bill\s+reference:\s*(\d+)\s*\(([^)]+)\)\s*\n?\s*"
    r"Account\s+number:\s*A-\d+",
    re.IGNORECASE,
)

# Electricity <from> - <to> £<amt>
# Months appear abbreviated or full (e.g. "Sept." vs "September"); the day
# suffix is sometimes wrapped in a thousands-separator comma for amounts.
_RECON_CHARGE_RE = re.compile(
    r"Electricity\s+"
    r"(\d{1,2}\s+[A-Za-z\.]{3,9}\.?\s+\d{4})\s*-\s*"
    r"(\d{1,2}\s+[A-Za-z\.]{3,9}\.?\s+\d{4})\s+"
    r"£([\d,]+\.\d{2})",
    re.IGNORECASE,
)

# Reversed electricity charge <date> £<amt>
# The next non-empty line typically carries the reversed period in parens,
# e.g. "(14 May 2024 - 30 Sept. 2024)" -- captured by a follow-up search.
_RECON_REVERSAL_RE = re.compile(
    r"Reversed\s+electricity\s+charge\s+"
    r"(\d{1,2}\s+[A-Za-z\.]{3,9}\.?\s+\d{4})\s+"
    r"£([\d,]+\.\d{2})",
    re.IGNORECASE,
)

# Parenthetical period sometimes follows a reversal row.
_RECON_REVERSAL_PERIOD_RE = re.compile(
    r"\(\s*(\d{1,2}\s+[A-Za-z\.]{3,9}\.?\s+\d{4})\s*-\s*"
    r"(\d{1,2}\s+[A-Za-z\.]{3,9}\.?\s+\d{4})\s*\)",
    re.IGNORECASE,
)

_RECON_LATE_PAYMENT_RE = re.compile(
    r"Late\s+Payment\s+Charge\s+£([\d,]+\.\d{2})",
    re.IGNORECASE,
)

# Payment rows: a date followed by £<amount>. Only matches inside the
# Payments section -- see ``extract_reconciliation_statement_rows`` which
# scopes the search to the post-"Payments" text block.
_RECON_PAYMENT_RE = re.compile(
    r"\b(\d{1,2}\s+[A-Za-z\.]{3,9}\.?\s+\d{4})\s+£([\d,]+\.\d{2})",
    re.IGNORECASE,
)

_RECON_BALANCE_LAST_RE = re.compile(
    r"Balance\s+on\s+your\s+last\s+bill\s+£([\d,]+\.\d{2})\s*(debit|credit)?",
    re.IGNORECASE,
)

_RECON_NEW_BALANCE_RE = re.compile(
    r"Your\s+new\s+balance\s+£([\d,]+\.\d{2})\s*(debit|credit)?",
    re.IGNORECASE,
)


def _recon_to_iso(s: str) -> str:
    """Convert a single date string like ``14 May 2024`` to ``14/05/2024``."""
    s = s.strip().rstrip(".")
    m = re.match(r"(\d{1,2})\s+([A-Za-z\.]+)\s+(\d{4})", s)
    if not m:
        return "N/A"
    day = int(m.group(1))
    month_str = m.group(2).rstrip(".").lower()
    year = int(m.group(3))
    month = _RECON_MONTH_MAP.get(month_str)
    if month is None:
        return "N/A"
    return f"{day:02d}/{month:02d}/{year:04d}"


def _recon_money(s: str) -> float:
    return float(s.replace(",", ""))


def detect_reconciliation_statement(text: str) -> bool:
    return bool(_RECON_STATEMENT_RE.search(text[:2000]))


def extract_reconciliation_statement_rows(text: str, attachment_name: str) -> list[dict]:
    """Extract every charge, reversal, late-payment, payment + one meta row

    from a consolidation reconciliation statement PDF.
    """
    rows: list[dict] = []
    src = "Statement Reconciliation"

    def _excerpt_around(m: re.Match, window: int = 400) -> str:
        """Return up to ``window`` chars around the regex match."""
        start = max(0, m.start(0) - 20)
        end = min(len(text), m.end(0) + window)
        return text[start:end]

    bill_ref = ""
    bill_date_display = "N/A"
    bill_ref_match = _RECON_STATEMENT_RE.search(text)
    if bill_ref_match:
        bill_ref = bill_ref_match.group(1)
        bill_date_display = _recon_to_iso(bill_ref_match.group(2))

    bal_last: object = "N/A"
    bal_last_match = _RECON_BALANCE_LAST_RE.search(text)
    if bal_last_match:
        bal_last = _recon_money(bal_last_match.group(1))

    new_bal: object = "N/A"
    new_bal_match = _RECON_NEW_BALANCE_RE.search(text)
    if new_bal_match:
        new_bal = _recon_money(new_bal_match.group(1))

    # Charge rows
    for m in _RECON_CHARGE_RE.finditer(text):
        rows.append(
            {
                "Source": src,
                "Sender": "",
                "Date": bill_date_display,
                "Period From": _recon_to_iso(m.group(1)),
                "Period To": _recon_to_iso(m.group(2)),
                "Invoice #": bill_ref or "N/A",
                "Amount (£)": _recon_money(m.group(3)),
                "Period Charge (£)": _recon_money(m.group(3)),
                "Entry Type": "Charge",
                "Reading": "N/A",
                "Units (kWh)": "N/A",
                "Standing Chg (p/day)": "N/A",
                "Tariff": "N/A",
                "Attachment Name": attachment_name,
                "Details": "Electricity charge (reconciliation statement)",
                "Logic Used": "Reconciliation Statement Charge",
                "Balance Last Bill (£)": bal_last,
                "Source PDF Text": _excerpt_around(m),
                "_regex_trace": "recon _RECON_CHARGE_RE",
            }
        )

    # Reversed-electricity-charge rows
    for m in _RECON_REVERSAL_RE.finditer(text):
        date_iso = _recon_to_iso(m.group(1))
        amount = _recon_money(m.group(2))
        # Look for a parenthetical period on the next non-empty line.
        details = "Reversed electricity charge"
        tail = text[m.end() : m.end() + 400]
        period_match = _RECON_REVERSAL_PERIOD_RE.search(tail)
        if period_match:
            details = (
                f"Reversed electricity charge ({period_match.group(1)} - {period_match.group(2)})"
            )
        rows.append(
            {
                "Source": src,
                "Sender": "",
                "Date": date_iso,
                "Period From": "N/A",
                "Period To": "N/A",
                "Invoice #": bill_ref or "N/A",
                "Amount (£)": -abs(amount),
                "Period Charge (£)": -abs(amount),
                "Entry Type": "Credit",
                "Reading": "N/A",
                "Units (kWh)": "N/A",
                "Standing Chg (p/day)": "N/A",
                "Tariff": "N/A",
                "Attachment Name": attachment_name,
                "Details": details,
                "Logic Used": "Reconciliation Statement Reversal",
                "Balance Last Bill (£)": bal_last,
                "Source PDF Text": _excerpt_around(m),
                "_regex_trace": "recon _RECON_REVERSAL_RE",
            }
        )

    # Late payment rows
    for m in _RECON_LATE_PAYMENT_RE.finditer(text):
        amount = _recon_money(m.group(1))
        rows.append(
            {
                "Source": src,
                "Sender": "",
                "Date": bill_date_display,
                "Period From": "N/A",
                "Period To": "N/A",
                "Invoice #": bill_ref or "N/A",
                "Amount (£)": amount,
                "Period Charge (£)": amount,
                "Entry Type": "Late Payment",
                "Reading": "N/A",
                "Units (kWh)": "N/A",
                "Standing Chg (p/day)": "N/A",
                "Tariff": "N/A",
                "Attachment Name": attachment_name,
                "Details": "Late Payment Charge (reconciliation statement)",
                "Logic Used": "Reconciliation Statement Late Payment",
                "Balance Last Bill (£)": bal_last,
                "Source PDF Text": _excerpt_around(m),
                "_regex_trace": "recon _RECON_LATE_PAYMENT_RE",
            }
        )

    # Payment rows -- scoped to the section starting "Payments" through either
    # "Your new balance" or end-of-text. EDF lists payments with a date column
    # then a £ column.
    payments_block = ""
    pay_section_match = re.search(r"Payments\s*\n", text, re.IGNORECASE)
    if pay_section_match:
        block_start = pay_section_match.end()
        # End payment block at "Your new balance" or end-of-text.
        end_match = re.search(r"Your\s+new\s+balance", text[block_start:], re.IGNORECASE)
        block_end = block_start + end_match.start() if end_match else len(text)
        payments_block = text[block_start:block_end]

    if payments_block:
        for m in _RECON_PAYMENT_RE.finditer(payments_block):
            rows.append(
                {
                    "Source": src,
                    "Sender": "",
                    "Date": _recon_to_iso(m.group(1)),
                    "Period From": "N/A",
                    "Period To": "N/A",
                    "Invoice #": bill_ref or "N/A",
                    "Amount (£)": _recon_money(m.group(2)),
                    "Period Charge (£)": "N/A",
                    "Entry Type": "Payment",
                    "Reading": "N/A",
                    "Units (kWh)": "N/A",
                    "Standing Chg (p/day)": "N/A",
                    "Tariff": "N/A",
                    "Attachment Name": attachment_name,
                    "Details": "Payment received (reconciliation statement)",
                    "Logic Used": "Reconciliation Statement Payment",
                    "Balance Last Bill (£)": bal_last,
                    "Source PDF Text": _excerpt_around(m),
                    "_regex_trace": "recon _RECON_PAYMENT_RE",
                }
            )

    # Always emit one meta row carrying the statement-level context.
    rows.append(
        {
            "Source": src,
            "Sender": "",
            "Date": bill_date_display,
            "Period From": "N/A",
            "Period To": "N/A",
            "Invoice #": bill_ref or "N/A",
            "Amount (£)": new_bal,
            "Period Charge (£)": "N/A",
            "Entry Type": "Statement Reconciliation",
            "Reading": "N/A",
            "Units (kWh)": "N/A",
            "Standing Chg (p/day)": "N/A",
            "Tariff": "N/A",
            "Attachment Name": attachment_name,
            "Balance Last Bill (£)": bal_last,
            "Details": f"Statement reconciliation: bill ref {bill_ref}",
            "Logic Used": "Reconciliation Statement Meta",
            # The meta row carries the statement-level context
            # (bill ref + balances); there is no single regex match
            # to excerpt. Provide the first 600 chars of the statement
            # so a reviewer sees the statement header context.
            "Source PDF Text": text[:600],
            "_regex_trace": "recon meta",
        }
    )
    return rows


# ---------------------------------------------------------------------------
# PST attachment filename recovery (Stream P6 / Task 9)
# ---------------------------------------------------------------------------
# The legacy code tried ``att.name``, ``att.get_name()``,
# ``att.get_long_filename()`` and ``att.get_short_filename()`` -- none of
# these exist on ``pypff.attachment``, so the lookup always raised
# ``AttributeError`` and every PST PDF row was emitted as ``Attachment_N.pdf``.
#
# Verified against ``scratch/input/edf.pst``: a real attachment exposes
# ``record_sets[i].get_entry(j)`` carrying:
#   * ``.entry_type``  -- the MAPI tag (e.g. ``0x3707`` for ``PR_ATTACH_LONG_FILENAME``).
#   * ``.value_type``  -- the MAPI data type (e.g. ``0x001F`` for ``PT_UNICODE``).
#   * ``.get_data_as_string()`` -- returns an already-decoded Python ``str``
#     (no manual UTF-16LE decode needed).
#
# MAPI tag constants from [MS-OXPROPS]:
_PST_PR_ATTACH_LONG_FILENAME = 0x3707
_PST_PR_ATTACH_FILENAME = 0x3704


__all__ = [
    "detect_sap_dump",
    "parse_sap_contract_history",
    "parse_sap_financial_transactions",
    "parse_sap_meter_read_history",
    "extract_new_credit_fields",
    "extract_new_invoice_fields",
    "extract_reconciliation_statement_rows",
    "detect_reconciliation_statement",
]
