"""Fallback extractor functions and PST/OST helpers extracted from.

``edf_collector.py``.

This module is the single source of truth for:

- ``_fallback_inv_num`` — multi-regex invoice-number fallback chain
  (canonical → cover-body → loose bare-token).
- ``_fallback_period_from`` / ``_fallback_period_to`` — billing-period
  fallback chain (canonical → cover-body).
- ``_fallback_amount`` — amount fallback chain (period-charge →
  credit-total → pound-amount).
- ``_pst_attachment_filename`` — walks the MAPI record-sets of a
  ``pypff.attachment`` and returns its long filename.
- ``_extract_sender_email`` — extracts the sender email from a
  ``pypff`` message via transport headers or sender name.
- ``_matches_domain_filter`` — checks whether a sender email matches
  a comma-separated domain filter string.

Dependency regexes live in :mod:`edf_bill_fetcher.processors.patterns`
so the package is self-contained (no circular import back into
``edf_collector``).

Compat re-exports live in ``edf_collector.py`` so callers using
``from edf_collector import _fallback_amount`` continue to work.
"""

from __future__ import annotations

from edf_bill_fetcher.processors.patterns import (
    _BILLING_PERIOD_RE,
    _COVER_BLOCK_INV_RE,
    _COVER_BLOCK_PERIOD_RE,
    _CREDIT_NUMBER_RE,
    _CREDIT_TOTAL_RE,
    _EMAIL_ADDR_RE,
    _FALLBACK_INV_RE,
    _FROM_HEADER_RE,
    _INV_NUMBER_RE,
    _PERIOD_CHARGE_RE,
    _POUND_AMOUNT_FALLBACK_RE,
    _PST_PR_ATTACH_LONG_FILENAME,
)


def _fallback_inv_num(text: str) -> tuple[str | None, str]:
    """Try the canonical invoice-number regex, then the cover-body regex,.

    then a loose bare-token regex. Returns (value, regex_name) or (None, "").
    """
    for label, pat in (
        ("_INV_NUMBER_RE", _INV_NUMBER_RE),
        ("_CREDIT_NUMBER_RE", _CREDIT_NUMBER_RE),
        ("_COVER_BLOCK_INV_RE", _COVER_BLOCK_INV_RE),
        ("_FALLBACK_INV_RE", _FALLBACK_INV_RE),
    ):
        m = pat.search(text[:3000])
        if m:
            val = m.group(1).strip() if m.lastindex else m.group(0)
            return val, label
    return None, ""


def _fallback_period_from(text: str) -> tuple[str | None, str]:
    """Return (period_from_str, regex_name)."""
    m = _BILLING_PERIOD_RE.search(text[:3000])
    if m:
        return m.group(1).strip(), "_BILLING_PERIOD_RE"
    m = _COVER_BLOCK_PERIOD_RE.search(text[:3000])
    if m:
        return m.group(1).strip(), "_COVER_BLOCK_PERIOD_RE"
    return None, ""


def _fallback_period_to(text: str) -> tuple[str | None, str]:
    """Return (period_to_str, regex_name)."""
    m = _BILLING_PERIOD_RE.search(text[:3000])
    if m:
        return m.group(2).strip(), "_BILLING_PERIOD_RE"
    m = _COVER_BLOCK_PERIOD_RE.search(text[:3000])
    if m:
        return m.group(2).strip(), "_COVER_BLOCK_PERIOD_RE"
    return None, ""


def _fallback_amount(text: str) -> tuple[float | None, str]:
    """Return (amount, regex_name) or (None, "")."""
    m = _PERIOD_CHARGE_RE.search(text[:3000])
    if m:
        return float(m.group(1).replace(",", "")), "_PERIOD_CHARGE_RE"
    m = _CREDIT_TOTAL_RE.search(text[:3000])
    if m:
        return float(m.group(1).replace(",", "")), "_CREDIT_TOTAL_RE"
    m = _POUND_AMOUNT_FALLBACK_RE.search(text[:3000])
    if m:
        return float(m.group(1).replace(",", "")), "_POUND_AMOUNT_FALLBACK_RE"
    return None, ""


def _pst_attachment_filename(att: object) -> str | None:
    """Walk the MAPI record-sets of a ``pypff.attachment`` and return its filename.

    Returns the filename string (``str``) when the ``PR_ATTACH_LONG_FILENAME``
    entry is found, else ``None``.  The caller is expected to fall back to
    ``Attachment_N.pdf`` (or whatever synthetic name) when this returns
    ``None``.

    Designed to tolerate malformed record-sets: a missing record entry,
    broken record collection, or zero-record attachment produce a clean
    ``None`` rather than propagating ``AttributeError`` / ``IndexError``
    out to the caller.
    """
    if att is None:
        return None
    # ``get_number_of_record_sets`` / ``get_record_set`` are the public methods
    # on ``pypff.attachment``; the legacy code never reached them.
    getter_count = getattr(att, "get_number_of_record_sets", None)
    if getter_count is None:
        return None
    try:
        n = int(getter_count())
    except Exception:
        return None
    for i in range(n):
        try:
            rs = att.get_record_set(i)  # type: ignore[attr-defined]
        except Exception:
            continue
        entries_getter = getattr(rs, "get_number_of_entries", None)
        if entries_getter is None:
            continue
        try:
            m = int(entries_getter())
        except Exception:
            continue
        for j in range(m):
            try:
                entry = rs.get_entry(j)  # type: ignore[attr-defined]
            except Exception:
                continue
            try:
                entry_type = int(entry.entry_type)  # type: ignore[attr-defined]
            except Exception:
                continue
            if entry_type != _PST_PR_ATTACH_LONG_FILENAME:
                continue
            # ``get_data_as_string()`` returns an already-decoded Python
            # str (verified on the real PST). Keep a fallback to manual
            # UTF-16LE decode for the rare PT_UNICODE raw-bytes edge case
            # so the helper never crashes on a pypff version mismatch.
            try:
                val = entry.get_data_as_string()  # type: ignore[attr-defined]
            except Exception:
                continue
            if isinstance(val, str) and val:
                return val
            # Some legacy builds return raw bytes; decode them safely.
            try:
                raw_data = entry.get_data()  # type: ignore[attr-defined]
            except Exception:
                continue
            if isinstance(raw_data, bytes | bytearray) and raw_data:
                try:
                    decoded = bytes(raw_data).decode("utf-16-le", errors="replace")
                except Exception:
                    continue
                if decoded.strip("\x00"):
                    return decoded.strip("\x00")
    return None


def _extract_sender_email(msg: object) -> str:
    """Extract sender email address from a pypff message, trying multiple methods."""
    sender: str | None = None
    # Try transport headers first (most reliable for SMTP email address)
    try:
        headers = msg.get_transport_headers()  # type: ignore[attr-defined]
        if headers:
            headers_str = (
                headers if isinstance(headers, str) else headers.decode("utf-8", errors="replace")
            )
            m = _FROM_HEADER_RE.search(headers_str)
            if m:
                sender = m.group(1).lower()
    except Exception:
        pass
    # Fallback: try sender name field (sometimes contains email)
    if not sender:
        try:
            name = msg.get_sender_name() or ""  # type: ignore[attr-defined]
            m = _EMAIL_ADDR_RE.search(name)
            if m:
                sender = m.group(1).lower()
        except Exception:
            pass
    return sender or ""


def _matches_domain_filter(sender_email: str, filter_str: str) -> bool:
    """Check if sender_email matches the domain filter string.

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


__all__ = [
    "_fallback_amount",
    "_fallback_inv_num",
    "_fallback_period_from",
    "_fallback_period_to",
    "_extract_sender_email",
    "_matches_domain_filter",
    "_pst_attachment_filename",
]
