"""MSG reading adapter — ``extract_msg`` parser for Outlook ``.msg`` message files.

Parses a single Outlook ``.msg`` message into the flat dict shape the
engine's folder-ingestion surface consumes::

    {"sender": str, "subject": str, "date_str": str,
     "body_html": str | None, "body_text": str | None}

``extract-msg`` is an optional dependency: this module imports cleanly
without it (``HAS_EXTRACT_MSG`` is ``False``) and
:func:`parse_msg_message` raises an informative ``ImportError`` when
called in that environment.  The library import happens inside the
function body so that top-level imports of this module never fail.
"""

from __future__ import annotations

from pathlib import Path

from edf_bill_fetcher.helpers.date_utils import parse_to_display_date

__all__ = ["HAS_EXTRACT_MSG", "parse_msg_message"]

try:
    import extract_msg  # noqa: F401

    HAS_EXTRACT_MSG = True
except ImportError:
    HAS_EXTRACT_MSG = False


def _decode_html_body(raw: bytes | None) -> str:
    """Decode the raw HTML body bytes to text, or ``""`` when absent."""
    if not raw:
        return ""
    return raw.decode("utf-8", errors="replace")


def parse_msg_message(path: str | Path) -> dict[str, str | None]:
    """Parse a ``.msg`` message file into the flat 5-key record dict.

    Opens the file with ``extract_msg.Message`` and reads ``sender``,
    ``subject``, ``date``, ``body`` and ``htmlBody``.  A missing date or
    body field yields an empty string; the date is formatted through
    ``parse_to_display_date``.

    Raises ``ImportError`` with an install hint when ``extract-msg`` is
    not available in the current environment.
    """
    try:
        import extract_msg
    except ImportError as err:
        raise ImportError(
            "extract-msg is required to parse .msg files; "
            "install it with `pip install 'edf-bill-fetcher[msg]'`"
        ) from err

    msg = extract_msg.Message(str(path))
    try:
        sender = msg.sender or ""
        subject = msg.subject or ""
        date_str = parse_to_display_date(msg.date.strftime("%d/%m/%Y %H:%M:%S")) if msg.date else ""
        body_html = _decode_html_body(msg.htmlBody)
        body_text = msg.body or ""
    finally:
        msg.close()

    return {
        "sender": sender,
        "subject": subject,
        "date_str": date_str,
        "body_html": body_html,
        "body_text": body_text,
    }
