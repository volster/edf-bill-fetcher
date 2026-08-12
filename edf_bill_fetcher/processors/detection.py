"""Pure-pandas detectors for back-billing, rebilling, meter rollovers, PDF format, and reconciliation statements.

Extracted from ``edf_collector.py`` as part of the modularization refactor
(Task 5 - Phase 4).  Each detector returns a derived DataFrame; no LLM,
no external service.

Compat re-exports live in ``edf_collector.py`` so callers using
``from edf_collector import detect_back_billing`` continue to work;
stripped by Task 7.
"""

from __future__ import annotations

import re
from collections import defaultdict

import pandas as pd

from edf_bill_fetcher.helpers.date_utils import _safe_to_datetime, parse_to_display_date
from edf_bill_fetcher.io.adapters.pdf import legal_context as legal_context_fn  # noqa: E402,F401
from edf_bill_fetcher.processors.sap_parsers import (
    detect_reconciliation_statement,
    extract_reconciliation_statement_rows,
)
from edf_bill_fetcher.writers._helpers import _disclosed_label  # noqa: F401

# KI / KCR invoice-format presence regexes (still defined in edf_collector.py
# for backward compat).  Re-declare locally so detection.py is self-contained.
_KI_PRESENCE_RE = re.compile(r"invoice number:\s*KI-", re.IGNORECASE)
_KCR_PRESENCE_RE = re.compile(r"credit note number:\s*KCR-", re.IGNORECASE)

# Reconciliation statement regexes (kept local; private to this module).
_RECON_STATEMENT_RE = re.compile(
    r"Statement\s+reference:?\s*([A-Z0-9-]+).*?Statement\s+date:?\s*(\d{1,2}\s+\w+\s+\d{4})",
    re.IGNORECASE | re.DOTALL,
)
_RECON_BALANCE_LAST_RE = re.compile(
    r"Balance\s+(?:brought\s+forward|last\s+bill)\s*£([\d,]+\.\d{2})",
    re.IGNORECASE,
)
_RECON_NEW_BALANCE_RE = re.compile(
    r"Your\s+new\s+balance\s*£([\d,]+\.\d{2})",
    re.IGNORECASE,
)
_RECON_CHARGE_RE = re.compile(
    r"Electricity\s+charge[^\n]*?(\d{1,2}\s+\w+\s+\d{4})\s*[-–]\s*(\d{1,2}\s+\w+\s+\d{4})[^\n]*?£([\d,]+\.\d{2})",
    re.IGNORECASE,
)
_RECON_REVERSAL_RE = re.compile(
    r"Reversed\s+electricity\s+charge[^\n]*?(\d{1,2}\s+\w+\s+\d{4})[^\n]*?£([\d,]+\.\d{2})",
    re.IGNORECASE,
)
_RECON_REVERSAL_PERIOD_RE = re.compile(
    r"for\s+the\s+period\s+(\d{1,2}\s+\w+\s+\d{4})\s*[-–]\s*(\d{1,2}\s+\w+\s+\d{4})",
    re.IGNORECASE,
)
_RECON_LATE_PAYMENT_RE = re.compile(
    r"Late\s+payment[^\n]*?£([\d,]+\.\d{2})",
    re.IGNORECASE,
)
_RECON_PAYMENT_RE = re.compile(
    r"(\d{1,2}\s+\w+\s+\d{4})[^\n]*?£([\d,]+\.\d{2})",
    re.IGNORECASE,
)


# Reconciliation statement helpers (kept local; private to this module).
def _recon_to_iso(s: str) -> str:
    """Convert "DD Mon YYYY" to ISO "YYYY-MM-DD"; return "N/A" on failure."""
    try:
        return parse_to_display_date(s.strip())  # type: ignore[no-any-return]
    except Exception:
        return "N/A"


def _recon_money(s: str) -> float:
    """Parse a "1,234.56" string to float; return 0.0 on failure."""
    try:
        return float(s.replace(",", "").replace("£", "").strip())
    except (ValueError, AttributeError):
        return 0.0


# Helper used by detect_back_billing.  Stays local to keep module self-contained.
def _assess_reason(
    invoice: str,
    bill_date: pd.Timestamp,
    excess: int,
    admitted: bool,
    period_from: pd.Timestamp,
    period_to: pd.Timestamp,
) -> str:
    """Return a short, deterministic narrative for the Reason Assessment column of the Back-billing sheet.

    Template-driven (no LLM).  The narrative is keyed to the legally
    correct back-billing rule (SLC 7A / Electricity Act 1989 s.84B):
    a bill is back-billing when it charges for consumption supplied
    more than 12 months before the bill Date.  ``excess`` is the count
    of consumption days in the period that fall more than 365 days
    before ``bill_date``.
    """
    pf = period_from.strftime("%d %b %Y")
    pt = period_to.strftime("%d %b %Y")
    bd = bill_date.strftime("%d %b %Y")
    if admitted:
        head = (
            f"Invoice {invoice} billed on {bd} for consumption from {pf} to {pt}; "
            f"{excess} days of consumption were supplied more than 12 months before the bill, "
            "exceeding the SLC 7A back-billing limit. "
            "EDF's cover page admits a cancellation/reversal, which is "
            "direct evidence the bill is a back-billing remedy."
        )
    else:
        head = (
            f"Invoice {invoice} billed on {bd} for consumption from {pf} to {pt}; "
            f"{excess} days of consumption were supplied more than 12 months before the bill, "
            "exceeding the SLC 7A back-billing limit. No "
            "admit-phrase was found on the cover page."
        )
    return head


# Legal-context block for back-billing tab. Delegates to the canonical
# helper in ``io.adapters.pdf`` so any wording updates (and the test
# ``tests/test_legal_context.py`` that pins the first line) reach
# the back-billing sheet without further wiring.
legal_context = legal_context_fn


def detect_pdf_format(text):
    """Return 'new_invoice', 'new_credit', or 'old' based on document markers."""
    if _KI_PRESENCE_RE.search(text):
        return "new_invoice"
    if _KCR_PRESENCE_RE.search(text):
        return "new_credit"
    return "old"


# extract_new_invoice_fields moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); see re-export block at top of file.

# extract_new_credit_fields moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); see re-export block at top of file.

# _SAP_HEADER_RE helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.

# _SAP_CONTRACT_COLS helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.

# _SAP_METER_COLS helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.

# _SAP_FINANCIAL_COLS helper moved to ``edf_bill_fetcher.processors.sap_parsers`` during the modularization refactor (Task 4 — Phase 3); private helper.

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


def _pull_period_charge(r: pd.Series) -> tuple[float, str]:
    """Pull ``Period Charge (£)`` from the source row; fall back to ``Amount (£)``.

    Returns ``(charge, value_source)`` where ``value_source`` is
    ``"Period Charge"`` when the Period Charge column was used, or
    ``"Amount (fallback)"`` when Period Charge was absent, N/A, or
    unparseable and the Amount column was used instead.
    """
    pc_raw = r.get("Period Charge (£)")
    if pc_raw is not None:
        try:
            return float(pc_raw), "Period Charge"
        except (TypeError, ValueError):
            pass
    amt_raw = r.get("Amount (£)", 0)
    try:
        return float(amt_raw), "Amount (fallback)"
    except (TypeError, ValueError):
        return 0.0, "Amount (fallback)"


def detect_back_billing(df: pd.DataFrame) -> pd.DataFrame:
    """Return invoices that are back-billing under SLC 7A / Electricity Act 1989 s.84B.

    A bill is back-billing when it charges for consumption supplied
    more than 12 months before the bill Date.  The eligibility gate is
    ``Date - Period From > 365 days`` — i.e. the bill charges for
    consumption supplied more than 12 months before the bill Date.
    If even the earliest consumption (Period From) is within 365 days
    of the bill Date, none of the period's consumption is unlawful
    and the invoice is NOT back-billing.

    ``Excess Days = max(0, (Date - 365 days - Period From).days)`` —
    the count of consumption days in the period that fall more than
    365 days before the bill Date.

    The detector also pulls ``Period Charge (£)`` from the source
    record; if that column is absent, N/A, or unparseable, it falls
    back to ``Amount (£)`` and records the provenance in the
    ``Value Source`` column.

    The function tolerates a missing ``Cancel/Rebill Admitted``
    column (treated as ``False``).

    Output columns:
        Invoice #, Bill Date, Period From, Period To, Days Billed,
        Period Charge (£), Value Source, 12-Month Limit (days),
        Excess Days, Unlawful Charge (£), Cancel/Rebill Admitted,
        Reason Assessment.

    ``Unlawful Charge (£)`` is the prorated share of the Period Charge
    attributable to the Excess Days — i.e.
    ``round(charge * (min(excess, days) / days), 2)`` where ``days`` is
    the full Days Billed span. The ratio is capped at 1.0: when the bill
    date falls more than 365 days after Period To (excess > days) the
    whole period is unlawful, so the unlawful charge never exceeds the
    Period Charge. A reviewer seeing the full Period Charge might
    otherwise mistake the entire amount as at issue, when only the
    Excess Days portion is unlawful.

    Rows with unparseable ``Period From``/``Period To`` are skipped
    silently. Output is sorted by ``Bill Date`` and re-indexed.

    Architectural note (SAP cross-feeding):
    This detector takes only the inferred-evidence dataframe. SAP
    data-dump rows (Contract-and-Product-Change-History,
    Meter-Read-History, Financial-Transactions) are surfaced in
    their own tabs (SAP Contract History / SAP Meter Readings /
    SAP Financial Transactions) plus the cross-source
    Reconciliation tab; they are NOT joined back into
    ``detect_back_billing`` because:

      * SAP financial transactions carry a Document No. (e.g.
        ``531000424090``) not an Invoice #, and their Transaction
        Text is the generic ledger description
        (``Dr- Consum Billing Receivable`` etc.) -- they cannot
        be unambiguously matched to an inferred invoice.
      * SAP records have no ``Period From`` / ``Period To`` span
        (only Posting Date / Document Date) so they cannot
        independently drive a back-billing judgement.
      * The Reconciliation sheet is the proper place to surface
        agreements and disagreements between the inferred and
        SAP samples; naively joining SAP amounts into the
        backbilling tab would mislead the reviewer.
    If a future resource joins the two sources by a higher-fidelity
    key (e.g. PDF receipt number + SAP Document No. mapping
    table), wire the intersection through ``run_analysers`` here.
    """
    columns = [
        "Invoice #",
        "Bill Date",
        "Period From",
        "Period To",
        "Days Billed",
        "Period Charge (£)",
        "Value Source",
        "12-Month Limit (days)",
        "Excess Days",
        "Unlawful Charge (£)",
        "Cancel/Rebill Admitted",
        "Reason Assessment",
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=columns)
    has_admit = "Cancel/Rebill Admitted" in df.columns
    rows = []
    for _, r in df.iterrows():
        pf = _safe_to_datetime(r.get("Period From"))
        pt = _safe_to_datetime(r.get("Period To"))
        if pd.isna(pf) or pd.isna(pt):
            continue
        bill_date_dt = _safe_to_datetime(r.get("Date"))
        if pd.isna(bill_date_dt):
            continue
        # Legal gate: bill Date must be more than 365 days after Period From.
        # Any consumption day supplied more than 12 months before the bill
        # Date is unlawful (SLC 21BA per-unit test). If even the EARLIEST
        # consumption (Period From) is within 365 days, the whole period is
        # lawful and we skip.
        gap_from = int((bill_date_dt - pf).days)
        if gap_from <= 365:
            continue
        days = int((pt - pf).days)
        # Skip inverted (Period From > Period To) and zero-day periods: a
        # negative or zero day span carries no consumption and would otherwise
        # surface with nonsensical Days Billed / prorated charge values.
        if days <= 0:
            continue
        # Excess Days: consumption days supplied more than 365 days before bill Date.
        excess = max(0, int((bill_date_dt - pd.Timedelta(days=365) - pf).days))
        # Period Charge (£) with Amount (£) fallback.
        charge, value_source = _pull_period_charge(r)
        admitted = bool(r.get("Cancel/Rebill Admitted")) if has_admit else False
        bill_date_raw = r.get("Date", "")
        unlawful_charge = round(charge * (min(excess, days) / days), 2) if days > 0 else 0.0
        rows.append(
            {
                "Invoice #": r.get("Invoice #", ""),
                "Bill Date": bill_date_raw,
                "_bill_date_sort": bill_date_dt,
                "Period From": pf,
                "Period To": pt,
                "Days Billed": days,
                "Period Charge (£)": charge,
                "Value Source": value_source,
                "12-Month Limit (days)": 365,
                "Excess Days": excess,
                "Unlawful Charge (£)": unlawful_charge,
                "Cancel/Rebill Admitted": admitted,
                "Reason Assessment": _assess_reason(
                    r.get("Invoice #", ""),
                    bill_date_dt,
                    excess,
                    admitted,
                    pf,
                    pt,
                ),
            }
        )
    out = pd.DataFrame(rows)
    if out.empty:
        return pd.DataFrame(columns=columns)
    sort_key = out["_bill_date_sort"]
    out = out.drop(columns=["_bill_date_sort"])
    # Reorder rows by the sort key (parsed Bill Date, ascending).
    out = out.loc[sort_key.sort_values().index].reset_index(drop=True)
    return out[columns]


def _reversal_match(
    evidence_df: pd.DataFrame | None,
    killed_inv: str,
    killed_amount: float | None,
    killed_pf: pd.Timestamp,
    killed_pt: pd.Timestamp,
) -> bool:
    """Return whether a reversal-credit row in *evidence_df* matches the killed invoice well enough to count as rebilling evidence.

    Spec ref: 2026-07-16 §11. A reversal credit accepts the killed
    invoice when its amount is within ±£0.50 AND either its period
    overlaps the killed period by ≥ 30 days OR its period is
    unparseable (so we accept on amount alone, Entry Type == Credit).
    """
    if evidence_df is None or evidence_df.empty:
        return False
    if "Entry Type" not in evidence_df.columns:
        return False
    try:
        amount = abs(float(killed_amount or 0.0))
    except (TypeError, ValueError):
        return False
    matching = evidence_df[evidence_df["Entry Type"].isin(["Credit", "Payment"])]
    for _, row in matching.iterrows():
        try:
            row_amt = abs(float(row.get("Amount (£)", 0) or 0))
        except (TypeError, ValueError):
            continue
        if abs(row_amt - amount) > 0.50:
            continue
        rpf = _safe_to_datetime(row.get("Period From"))
        rpt = _safe_to_datetime(row.get("Period To"))
        if pd.isna(rpf) or pd.isna(rpt):
            return True
        overlap = (min(killed_pt, rpt) - max(killed_pf, rpf)).days
        if overlap >= 30:
            return True
    return False


def detect_rebilling(
    df: pd.DataFrame,
    *,
    evidence_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """Return cancel-and-repost pairs identified by the rebilling heuristic (spec §11, tightened gate).

    For each ordered pair ``(Killer, Killed)`` where ``Killer.Date``
    is strictly later than ``Killed.Date``, emit a row IFF ALL hold:

    1. ``Killer.Period From ≤ Killed.Period From AND Killer.Period To ≥
       Killed.Period To`` -- the killer's billing window fully contains
       the killed's billing window.
    2. ANY of these signals also fires:
       - ``Killer.Days Billed ≥ 365`` (wholesale cancel-and-repost of a
         long period),
       - the killer invoice has ``Cancel/Rebill Admitted = True``
         (an admission phrase like ``corrected`` / ``amended`` was
         detected on the source PDF), OR
       - a reversal credit row in ``evidence_df`` matches the killed
         invoice's amount within ±£0.50 and period overlap ≥ 30 days
         (or its period is unparseable, in which case amount alone
         suffices).

    Output columns:
        Killer Invoice, Killed Invoice, Killer Date, Killed Date,
        Period Overlap (days), Jump-back (days), Trigger Reason,
        Cancel/Rebill Admitted (Killer).

    ``Cancel/Rebill Admitted (Killer)`` is the admit-phrase flag
    lifted from the killer invoice.

    ``evidence_df`` is optional -- when omitted, the reversal-credit
    check is skipped and only the long-period / admit-phrase signals
    fire. ``run_analysers`` passes the evidence DataFrame so the
    reversal signal participates in normal pipeline use.
    """
    columns = [
        "Killer Invoice",
        "Killed Invoice",
        "Killer Date",
        "Killed Date",
        "Period Overlap (days)",
        "Jump-back (days)",
        "Trigger Reason",
        "Cancel/Rebill Admitted (Killer)",
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=columns)
    has_admit = "Cancel/Rebill Admitted" in df.columns
    rows = []
    parsed = []
    for _, r in df.iterrows():
        pf = _safe_to_datetime(r.get("Period From"))
        pt = _safe_to_datetime(r.get("Period To"))
        bd = _safe_to_datetime(r.get("Date"))
        if pd.isna(pf) or pd.isna(pt) or pd.isna(bd):
            continue
        try:
            amount = float(r.get("Amount (£)", 0) or 0)
        except (TypeError, ValueError):
            amount = None
        admitted = bool(r.get("Cancel/Rebill Admitted")) if has_admit else False
        parsed.append(
            {
                "Invoice #": r.get("Invoice #", ""),
                "Date_raw": r.get("Date", ""),
                "Date": bd,
                "Period From": pf,
                "Period To": pt,
                "Days Billed": int((pt - pf).days),
                "Amount": amount,
                "Admitted": admitted,
            }
        )
    if len(parsed) < 2:
        return pd.DataFrame(columns=columns)
    parsed.sort(key=lambda x: x["Date"])
    for i, killer in enumerate(parsed):
        for killed in parsed[:i]:
            # Containment -- the only structural requirement.
            if not (
                killer["Period From"] <= killed["Period From"]
                and killer["Period To"] >= killed["Period To"]
            ):
                continue
            triggers: list[str] = []
            if killer["Days Billed"] >= 365:
                triggers.append("killer period \u2265 365d")
            admitted = killer["Admitted"]
            if admitted:
                triggers.append("admit-phrase on killer")
            reversal_match = _reversal_match(
                evidence_df,
                killed["Invoice #"],
                killed["Amount"],
                killed["Period From"],
                killed["Period To"],
            )
            if reversal_match:
                triggers.append("reversal credit row matches killed")
            if not triggers:
                continue
            trigger_reason = "; ".join(triggers)
            overlap_d = max(
                0,
                (
                    min(killer["Period To"], killed["Period To"])
                    - max(killer["Period From"], killed["Period From"])
                ).days,
            )
            jumpback_d = (killed["Period From"] - killer["Period From"]).days
            rows.append(
                {
                    "Killer Invoice": killer["Invoice #"],
                    "Killed Invoice": killed["Invoice #"],
                    "Killer Date": killer["Date_raw"],
                    "Killed Date": killed["Date_raw"],
                    "Period Overlap (days)": overlap_d,
                    "Jump-back (days)": max(0, jumpback_d),
                    "Trigger Reason": trigger_reason,
                    "Cancel/Rebill Admitted (Killer)": admitted,
                }
            )
    if not rows:
        return pd.DataFrame(columns=columns)
    out = pd.DataFrame(rows, columns=columns)
    out["_k_sort"] = _safe_to_datetime(out["Killer Date"])
    out["_d_sort"] = _safe_to_datetime(out["Killed Date"])
    sort_idx = out.sort_values(["_k_sort", "_d_sort"]).index
    out = out.loc[sort_idx].drop(columns=["_k_sort", "_d_sort"]).reset_index(drop=True)
    return out[columns]


def compute_transitive_domination(
    rebilling_df: pd.DataFrame,
    back_billing_rows: pd.DataFrame,
) -> dict[str, tuple[str, bool]]:
    """Compute the transitive closure of the kill-chain edges restricted to back-billing rows.

    Consumes:
        rebilling_df: output of detect_rebilling with columns
            'Killer Invoice', 'Killed Invoice' (and others).
        back_billing_rows: output of detect_back_billing with
            'Invoice #' key and 'Period From'/'Period To'.

    Returns a mapping {superseded_invoice_id: (survivor_invoice_id, partial_overlap_flag)}.
    A back-billing row is superseded iff there exists a later invoice that
    transitively dominates it. The survivor is the transitive root (the live row
    from which the superseded row is reachable). The survivor MAY be an invoice
    that is NOT itself a back-billing row (e.g. a regular monthly rebill that
    cancels/replaces a back-billing invoice); in that case it appears in the
    map as the superseding ID even though it has no row in the back-billing
    sheet.

    partial_overlap_flag is True when the killer's period does NOT fully contain the
    killed's period (strict containment guard failed but a K*-edge still fired).
    When the survivor is not a back-billing row (no period in period_map),
    partial_overlap defaults to False.
    """
    if rebilling_df.empty:
        edges: list[tuple[str, str]] = []
    else:
        edges = list(
            zip(
                rebilling_df["Killer Invoice"].astype(str),
                rebilling_df["Killed Invoice"].astype(str),
                strict=True,
            )
        )

    # Build the set of back-billing invoice IDs.
    bb_ids = {str(row["Invoice #"]) for _, row in back_billing_rows.iterrows()}

    period_map: dict[str, tuple[pd.Timestamp, pd.Timestamp]] = {}
    for _, row in back_billing_rows.iterrows():
        inv_id = str(row["Invoice #"])
        pf = _safe_to_datetime(row.get("Period From"))
        pt = _safe_to_datetime(row.get("Period To"))
        if pd.notna(pf) and pd.notna(pt):
            period_map[inv_id] = (pf, pt)

    # Widen the edge filter to include any edge where the KILLED endpoint
    # is a back-billing row, regardless of whether the killer is. A
    # non-back-billing rebill (e.g. a regular monthly bill) CAN supersede
    # a back-billing invoice; the survivor then carries that killer's ID
    # even though it has no row in the back-billing sheet.
    bb_edges = [(u, v) for u, v in edges if v in bb_ids]

    adj: dict[str, list[str]] = defaultdict(list)
    for u, v in bb_edges:
        adj[u].append(v)

    # Sources for BFS include every back-billing row PLUS every non-bb
    # killer that appears as the 'u' endpoint of a widened edge — a
    # non-bb rebill can be the ultimate survivor that supersedes a bb
    # invoice, so its reachable set must be computed too.
    sources = set(bb_ids) | {u for u, _ in bb_edges}

    reachable_from: dict[str, set[str]] = defaultdict(set)
    for source in sources:
        visited: set[str] = set()
        queue = [source]
        while queue:
            node = queue.pop(0)
            if node in visited:
                continue
            visited.add(node)
            for neighbor in adj.get(node, []):
                if neighbor not in visited:
                    queue.append(neighbor)
        reachable_from[source] = visited - {source}

    def sort_key(inv_id: str) -> tuple[pd.Timestamp, str]:
        if inv_id in period_map:
            return (period_map[inv_id][0], inv_id)
        return (pd.Timestamp.max, inv_id)

    # First pass: identify which sources are themselves superseded by
    # any other source. A source is superseded if some other source can
    # reach it transitively. This lets the second pass pick the ultimate
    # root (a source not itself superseded) as the survivor.
    superseded_sources: set[str] = set()
    for source in sources:
        for other in sources:
            if other != source and source in reachable_from[other]:
                superseded_sources.add(source)
                break

    domination_map: dict[str, tuple[str, bool]] = {}
    for target in bb_ids:
        superseded_by = [
            source for source in sources if source != target and target in reachable_from[source]
        ]
        if not superseded_by:
            continue
        # The survivor is the ultimate transitive root — a source that is
        # not itself superseded by any other candidate. Only fall back to
        # earliest-period-start tiebreak if every candidate is superseded
        # (a cycle, which shouldn't happen in well-formed kill chains).
        roots = [s for s in superseded_by if s not in superseded_sources]
        survivor = max(roots if roots else superseded_by, key=sort_key)
        partial_overlap = False
        if survivor in period_map and target in period_map:
            survivor_start, survivor_end = period_map[survivor]
            target_start, target_end = period_map[target]
            if not (survivor_start <= target_start and survivor_end >= target_end):
                partial_overlap = True
        domination_map[target] = (survivor, partial_overlap)
    return domination_map


# Default 99,999 - 5,000 rollover threshold per spec \u00a73.3.
_DEFAULT_ROLLOVER_THRESHOLD = 99999 - 5000


def detect_meter_rollover(
    df: pd.DataFrame, rollover_threshold: int = _DEFAULT_ROLLOVER_THRESHOLD
) -> pd.DataFrame:
    """Return meter-rollover candidate events (spec \u00a73.3).

    Walks the rows of *df* keeping only ones tagged ``Actual'' or
    ``Smart'' in the ``Reading`` column (supplier-confirmed readings
    only -- ``Estimated``/``Unknown`` rows don't count). For each
    consecutive (actual-or-smart, actual-or-smart) pair, computes
    delta = (curr Units (kWh)) - (prev Units (kWh)) -- i.e. the
    change in per-period kWh consumption -- and emits a row when the
    delta is negative AND its magnitude exceeds
    ``rollover_threshold`` (default 99,999 - 5,000 = 94,999).

    Output columns:
        Date, Invoice #, Prev Units (kWh), Curr Units (kWh),
        Delta, Reading Type, Notes.

    Rows with unparseable ``Units (kWh)`` or ``Date`` are skipped
    silently.
    """
    columns = [
        "Date",
        "Invoice #",
        "Prev Units (kWh)",
        "Curr Units (kWh)",
        "Delta",
        "Reading Type",
        "Notes",
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=columns)
    # Restrict to Actual/Smart only.
    mask = df.get("Reading", pd.Series(dtype=str)).isin(["Actual", "Smart"])
    candidates = df[mask].copy()
    if candidates.empty:
        return pd.DataFrame(columns=columns)
    # Parse dates so we can sort.
    candidates["_date_dt"] = _safe_to_datetime(candidates["Date"])
    candidates = candidates.dropna(subset=["_date_dt"])
    candidates = candidates.sort_values("_date_dt")
    rows = []
    prev_units: float | None = None
    prev_invoice = ""
    prev_date_raw = ""
    for _, r in candidates.iterrows():
        u_raw = r.get("Units (kWh)", "N/A")
        try:
            u = float(u_raw)
        except (TypeError, ValueError):
            prev_units = None
            continue
        if prev_units is not None:
            delta = u - prev_units
            if delta < 0 and abs(delta) > rollover_threshold:
                rows.append(
                    {
                        "Date": r.get("Date", ""),
                        "Invoice #": r.get("Invoice #", ""),
                        "Prev Units (kWh)": prev_units,
                        "Curr Units (kWh)": u,
                        "Delta": int(delta),
                        "Reading Type": r.get("Reading", ""),
                        "Notes": (
                            f"Negative jump of {abs(int(delta))} kWh between "
                            f"{prev_invoice} ({prev_date_raw}) and "
                            f"{r.get('Invoice #', '')} ({r.get('Date', '')}) -- "
                            "consistent with a meter rollover near the "
                            f"{rollover_threshold + 5000}-rollover cap."
                        ),
                    }
                )
        prev_units = u
        prev_invoice = r.get("Invoice #", "")
        prev_date_raw = r.get("Date", "")
    if not rows:
        return pd.DataFrame(columns=columns)
    out = pd.DataFrame(rows, columns=columns)
    sort_idx = _safe_to_datetime(out["Date"]).sort_values().index
    out = out.loc[sort_idx].reset_index(drop=True)
    return out[columns]


__all__ = [
    "detect_pdf_format",
    "detect_back_billing",
    "detect_rebilling",
    "compute_transitive_domination",
    "detect_meter_rollover",
    "detect_reconciliation_statement",
    "extract_reconciliation_statement_rows",
]
