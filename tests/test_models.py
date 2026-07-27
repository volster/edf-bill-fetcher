"""Tests for edf_bill_fetcher.models.events dataclasses."""
from __future__ import annotations

from edf_bill_fetcher.models.events import SapBackBillingEvent, SapEdfMatch


def test_sap_back_billing_event_has_required_fields():
    event = SapBackBillingEvent(
        clearing_doc="DOC001",
        clearing_date="2024-01-01",
        clearing_reason="Test reason",
        net_amount=100.0,
        has_credit_for_consum_billing=False,
    )
    assert event.clearing_doc == "DOC001"
    assert event.clearing_date == "2024-01-01"
    assert event.clearing_reason == "Test reason"
    assert event.net_amount == 100.0
    assert event.has_credit_for_consum_billing is False


def test_sap_edf_match_has_required_fields():
    event = SapBackBillingEvent(
        clearing_doc="DOC002",
        clearing_date="2024-01-01",
        clearing_reason="Test",
    )
    match = SapEdfMatch(
        event=event,
        edf_record={"Invoice #": "INV-001", "Amount (£)": 100.0},
        confidence_band="Low",
        confidence_score=25,
        amount_delta=0.0,
        date_delta_days=0,
        notes="Exact match",
    )
    assert match.event.clearing_doc == "DOC002"
    assert match.edf_record["Invoice #"] == "INV-001"
    assert match.confidence_band == "Low"


def test_sap_back_billing_event_default_values():
    event = SapBackBillingEvent(
        clearing_doc="DOC003",
        clearing_date="2024-01-01",
        clearing_reason="Test",
    )
    assert event.net_amount == 0.0
    assert event.has_credit_for_consum_billing is False
    assert event.rows == []


def test_models_submodule_importable():
    from edf_bill_fetcher.models import SapBackBillingEvent, SapEdfMatch

    assert SapBackBillingEvent is not None
    assert SapEdfMatch is not None
