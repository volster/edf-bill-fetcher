"""Uncertain identities and contradictory originals must never erase financial evidence."""

import copy
from decimal import Decimal

import pytest

from edf_bill_fetcher.processors.compensation import estimate_compensation
from edf_bill_fetcher.processors.ledger_context import fuel_key, resolve_ledger_context
from edf_bill_fetcher.processors.posting_identity import link_posting_sequences
from edf_bill_fetcher.processors.preparation import (
    _link_account_history_formats,
    money,
    normalize_record,
    prepare_evidence,
    scope,
)
from edf_bill_fetcher.processors.rebilling_ledger import build_ledger, components, validate_schedule
from edf_bill_fetcher.processors.rebilling_ledger import money as ledger_money
from edf_bill_fetcher.processors.revisions import build_revisions
from tests.test_claim_reconciliation import evidence
from tests.test_rebilling_ledger import position


def bill(**changes):
    return {
        "Account": "A-EXAMPLE",
        "Fuel": "Electricity",
        "Meter": "M1",
        "Register": "R1",
        "Invoice #": "INV-1",
        "Date": "01/02/2024",
        "Period From": "01/01/2024",
        "Period To": "31/01/2024",
        "Entry Type": "New Bill",
        "Period Charge (£)": 100,
        "Transaction Amount (£)": 100,
        "Balance (£)": 150,
        "Source": "Local PDF Folder",
        "Source ID": "original",
        **changes,
    }


@pytest.mark.parametrize("convert", [money, ledger_money])
def test_unrepresentable_money_is_unavailable_instead_of_zero(convert):
    assert convert(Decimal("1e100")) is None


@pytest.mark.parametrize("label", ["na", "not applicable"])
def test_non_fuel_labels_do_not_create_a_ledger(label):
    assert fuel_key(label) == ""


def test_chained_account_aliases_remain_unmerged_and_originals_unchanged():
    records = [bill(Account=value) for value in ("111", "222", "333")]
    sources = [
        {"SAP Account Number": "111", "Kraken ID": "222", "Source ID": "header-1"},
        {"SAP Account Number": "222", "Kraken ID": "333", "Source ID": "header-2"},
    ]
    before = copy.deepcopy((records, sources))
    resolved, _ = resolve_ledger_context(records, sources)
    assert len({scope(row)[0] for row in resolved}) == 3
    assert (records, sources) == before


@pytest.mark.parametrize("ids", ['["first", "second"]', "unreadable JSON"])
def test_saved_source_id_strings_preserve_only_valid_lists(ids):
    original = bill(**{"Source IDs": ids})
    row = normalize_record(original)
    assert row["Source IDs"] == (["first", "second"] if ids.startswith("[") else [])
    assert original["Source IDs"] == ids


def test_invalid_posting_date_cannot_create_a_shared_payment_identity():
    rows = [bill(**{"Date": "unreadable", "Entry Type": "Payment", "Transaction Amount (£)": -100})]
    link_posting_sequences(rows)
    assert "Reconciled Event ID" not in rows[0]


@pytest.mark.parametrize(
    ("field", "value"),
    [
        ("Period From", "02/01/2024"),
        ("Period To", "30/01/2024"),
        ("Period Charge (£)", 101),
        ("Balance (£)", 151),
    ],
)
def test_contradictory_copies_with_same_invoice_number_remain_visible(field, value):
    prepared = prepare_evidence([bill(), bill(**{field: value, "Source ID": "other"})])
    assert len(prepared.records) == 2
    assert prepared.duplicates == []


@pytest.mark.parametrize(
    ("field", "value"),
    [
        ("Source", "HTM Account History"),
        ("Entry Type", "Payment"),
        ("Account", "OTHER"),
        ("Invoice #", ""),
        ("Date", "unreadable"),
        ("Date", "01/05/2024"),
        ("Meter", "OTHER-METER"),
        ("Fuel", "Gas"),
        ("Register", "OTHER-REGISTER"),
        ("Period To", "30/01/2024"),
        ("Transaction Amount (£)", 101),
    ],
)
def test_history_line_requires_a_compatible_dated_original(field, value):
    original = bill(**{field: value})
    historical = bill(
        **{
            "Entry Type": "Charge",
            "Source": "Statement Reconciliation",
            "Source ID": "statement",
            "Document Role": "Historical account statement",
            "Date": "N/A",
            "Document Date": "01/04/2024",
        }
    )
    prepared = prepare_evidence([original, historical], {"use_dedup": False})
    row = next(row for row in prepared.records if row["Source ID"] == "statement")
    assert "Original Invoice #" not in row
    assert "Reconciled Event ID" not in row


def test_history_credit_requires_its_own_matching_date():
    original = bill(**{"Entry Type": "Credit", "Transaction Amount (£)": -100})
    historical = {
        **original,
        "Source ID": "statement",
        "Source": "Statement Reconciliation",
        "Document Role": "Historical account statement",
        "Date": "02/02/2024",
    }
    prepared = prepare_evidence([original, historical], {"use_dedup": False})
    assert all("Original Invoice #" not in row for row in prepared.records)
    assert len(prepared.records) == 2


def test_repeated_events_in_one_history_source_are_flagged_without_merging():
    row = bill(**{"Source": "HTM Account History", "Posting Date": "01/02/2024"})
    prepared = prepare_evidence([row, copy.deepcopy(row)], {"use_dedup": False})
    assert len(prepared.records) == 2
    assert all(
        "Repeated identical account-history events" in r["Reconciliation Review"]
        for r in prepared.records
    )
    assert all("Reconciled Event ID" not in r for r in prepared.records)


def test_conflicting_corroborated_event_ids_are_not_overwritten_by_history_matching():
    rows = [
        normalize_record(
            bill(
                **{
                    "Source": source,
                    "Source ID": source,
                    "Posting Date": "01/02/2024",
                    "Reconciled Event ID": event,
                }
            )
        )
        for source, event in (
            ("HTM Account History", "event-one"),
            ("PDF Account History", "event-two"),
        )
    ]
    _link_account_history_formats(rows)
    assert [row["Reconciled Event ID"] for row in rows] == ["event-one", "event-two"]


def test_prepared_dataframe_is_detached_from_accepted_records():
    prepared = prepare_evidence([bill()])
    frame = prepared.frame
    frame.loc[0, "Balance (£)"] = 999
    assert prepared.records[0]["Balance (£)"] == 150
    assert prepared.frame.loc[0, "Balance (£)"] == 150


@pytest.mark.parametrize(
    ("field", "value"),
    [
        ("as_of", "unreadable"),
        ("min_amount", -1),
        ("analysis_min", "NaN"),
        ("credit_interest_rate", "Infinity"),
    ],
)
def test_invalid_financial_configuration_fails_before_analysis(field, value):
    with pytest.raises(ValueError, match=field):
        prepare_evidence([bill()], {field: value})


def test_invalid_interest_configuration_cannot_publish_compensation():
    import pandas as pd

    with pytest.raises(ValueError, match="finite and non-negative"):
        estimate_compensation(pd.DataFrame([bill()]), {"credit_interest_rate": -1})


def test_extracted_energy_is_not_duplicated_by_component_facts():
    row = bill(
        **{
            "Sub Periods": "01/01/2024|31/01/2024|100|100|100",
            "Rebilling Components": [
                {"kind": "energy", "amount": 999},
                {"kind": "vat", "amount": 5},
            ],
        }
    )
    original = copy.deepcopy(row)
    parts = components(row)
    assert [part["amount"] for part in parts if part["kind"] == "energy"] == ["100.00"]
    assert [part["amount"] for part in parts if part["kind"] == "vat"] == ["5.00"]
    assert row == original


@pytest.mark.parametrize("facts", [{"kind": "energy"}, [None, {"kind": "unknown"}]])
def test_unusable_component_facts_do_not_invent_an_allocation(facts):
    assert components(bill(**{"Rebilling Components": facts})) == []


@pytest.mark.parametrize(
    ("decisions", "message"),
    [
        ({}, "Invalid dispute schedule"),
        ([{}], "Invalid dispute schedule fields"),
        ([position("a", first_demand_date="0001-01-01")], "preceding 12-month"),
    ],
)
def test_invalid_dispute_schedule_is_rejected(decisions, message):
    items, _ = evidence()
    with pytest.raises(ValueError, match=message):
        validate_schedule(items, decisions)


def test_unreviewed_position_can_remain_explicitly_unquantified():
    items, _ = evidence()
    value = position(
        "a",
        status="unreviewed",
        challenged_amount="",
        proposed_reduction=None,
        first_demand_date="",
        basis="",
    )
    validated = validate_schedule(items, [value])[0]
    assert validated["challenged_amount"] is None
    assert validated["proposed_reduction"] is None
    assert validated["first_demand_date"] is None
    assert value["challenged_amount"] == ""


@pytest.mark.parametrize("missing", [2, 3, 4])
def test_component_comparison_requires_all_original_quantities(missing):
    items, _ = evidence()
    fields = ["01/01/2024", "31/01/2024", "100", "100", "100"]
    fields[missing] = "unreadable"
    items[0]["record"]["Sub Periods"] = "|".join(fields)
    items[1]["record"]["Sub Periods"] = "01/01/2024|31/01/2024|120|100|120"
    model = build_ledger(items, build_revisions(items))
    part = model["comparisons"][0]["components"][0]
    assert part["status"] == "unallocated"
    assert all(
        part[field] is None
        for field in ("charge_difference", "units_effect", "rate_effect", "residual")
    )
    assert [item["record"]["Period Charge (£)"] for item in items[:2]] == [100, 120]
