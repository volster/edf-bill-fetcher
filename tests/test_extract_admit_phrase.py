from __future__ import annotations

import pytest

from edf_collector import extract_admit_phrase


@pytest.mark.parametrize(
    ("text", "expected_start"),
    [
        (
            "We've recently cancelled some charges for you. This credit is "
            "included in your balance and is shown on page 2.",
            "We've recently cancelled",
        ),
        ("We have reversed some charges on your account.", "We have reversed"),
        ("We are reversing charges for you.", "We are reversing"),
    ],
)
def test_extract_admit_phrase_returns_matched_substring(text: str, expected_start: str) -> None:
    out = extract_admit_phrase(text)
    assert out is not None
    assert out.lower().startswith(expected_start.lower())


@pytest.mark.parametrize(
    "text",
    [
        "",
        "You can cancel your direct debit at any time.",
        "Please contact us to cancel your tariff renewal.",
        "Your charges for this period (including VAT) \u00a31,525.13",
    ],
)
def test_extract_admit_phrase_returns_none_when_no_admission(text: str) -> None:
    assert extract_admit_phrase(text) is None
