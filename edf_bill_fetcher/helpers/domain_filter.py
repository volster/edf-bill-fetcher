"""Shared sender-domain matching rules."""

from __future__ import annotations


def matches_domain_filter(sender_email: str, filter_str: str) -> bool:
    """Return whether a sender matches any configured address or domain."""
    if not sender_email or not filter_str:
        return False
    normalized_sender = sender_email.lower().strip()
    for pattern in (part.strip().lower() for part in filter_str.split(",")):
        if not pattern:
            continue
        if "@" in pattern:
            if normalized_sender == pattern:
                return True
            continue
        domain = pattern.lstrip("*").lstrip(".")
        sender_domain = normalized_sender.split("@")[-1] if "@" in normalized_sender else ""
        if sender_domain == domain or sender_domain.endswith("." + domain):
            return True
    return False
