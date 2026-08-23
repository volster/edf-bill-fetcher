"""EML reading adapter — stdlib ``email`` parser for ``.eml`` message files.

Parses a single RFC-5322 ``.eml`` message into the flat dict shape the
engine's folder-ingestion surface consumes::

    {"sender": str, "subject": str, "date_str": str,
     "body_html": str | None, "body_text": str | None}

Only the Python standard library is used for the MIME work
(``email.message`` / ``email.parser`` / ``email.header``).  Body parts
are walked in order and the first ``text/html`` and first ``text/plain``
part are captured; a missing body yields empty strings rather than
exceptions.  The ``Subject`` header is decoded with
:func:`email.header.decode_header` and the ``Date`` header is formatted
through :func:`edf_bill_fetcher.helpers.date_utils.parse_to_display_date`.
"""

from __future__ import annotations

from email import policy
from email.header import decode_header
from email.message import Message
from email.parser import BytesParser
from pathlib import Path

from edf_bill_fetcher.helpers.date_utils import parse_to_display_date

__all__ = ["parse_eml_message"]


def _decode_header_value(raw: str | None) -> str:
    """Decode an RFC 2047-encoded header value into plain text.

    Handles the mixed chunk list returned by
    :func:`email.header.decode_header` — plain ``str`` chunks for
    unencoded runs and ``bytes`` chunks for encoded words.  Bytes
    chunks are decoded with the chunk's charset, falling back to UTF-8
    with replacement characters for unknown charsets.
    """
    if not raw:
        return ""
    chunks: list[str] = []
    for chunk, charset in decode_header(raw):
        if isinstance(chunk, bytes):
            chunks.append(chunk.decode(charset or "utf-8", errors="replace"))
        else:
            chunks.append(chunk)
    return "".join(chunks)


def _decode_payload(part: Message) -> str | None:
    """Decode a non-multipart part's payload to text, or ``None``."""
    payload = part.get_payload(decode=True)
    if not isinstance(payload, bytes):
        return None
    charset = part.get_content_charset() or "utf-8"
    try:
        return payload.decode(charset, errors="replace")
    except LookupError:
        return payload.decode("utf-8", errors="replace")


def parse_eml_message(path: str | Path) -> dict[str, str | None]:
    """Parse a ``.eml`` message file into the flat 5-key record dict.

    Reads the file as bytes and parses it with the stdlib ``email``
    package, returning ``sender`` (From), ``subject`` (RFC 2047
    decoded), ``date_str`` (Date header formatted via
    ``parse_to_display_date``), and ``body_html`` / ``body_text`` from
    the first ``text/html`` and ``text/plain`` parts respectively.
    Body fields are empty strings when the corresponding part is
    missing.
    """
    data = Path(path).read_bytes()
    msg = BytesParser(policy=policy.default).parsebytes(data)

    sender = _decode_header_value(str(msg.get("From") or ""))
    subject = _decode_header_value(str(msg.get("Subject") or ""))
    raw_date = str(msg.get("Date") or "")
    date_str = parse_to_display_date(raw_date) if raw_date else ""

    body_html = ""
    body_text = ""
    for part in msg.walk():
        if part.is_multipart():
            continue
        text = _decode_payload(part)
        if text is None:
            continue
        content_type = part.get_content_type()
        if content_type == "text/html" and not body_html:
            body_html = text
        elif content_type == "text/plain" and not body_text:
            body_text = text

    return {
        "sender": sender,
        "subject": subject,
        "date_str": date_str,
        "body_html": body_html,
        "body_text": body_text,
    }
