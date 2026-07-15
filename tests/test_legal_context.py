from __future__ import annotations

from edf_collector import legal_context


def test_legal_context_returns_non_empty_string():
    out = legal_context()
    assert isinstance(out, str)
    assert len(out) > 100


def test_legal_context_mentions_back_billing():
    out = legal_context()
    assert "back-billing" in out.lower() or "back billing" in out.lower()


def test_legal_context_cites_ofgem_or_statute():
    out = legal_context()
    low = out.lower()
    assert "ofgem" in low or "electricity act" in low or "1989" in low
