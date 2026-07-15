from __future__ import annotations

import pytest

from edf_collector import _ADMIT_RE


@pytest.mark.parametrize(
    "text",
    [
        "We've recently cancelled some charges for you. This credit is "
        "included in your balance and is shown on page 2.",
        "We've recently cancelled charges for you.",
        "We've previously cancelled charges for your account.",
        "We have reversed some charges on your account.",
        "We are reversing charges for you.",
        "We've credited your account for cancelled charges.",
        "We have canceled some charges for you.",
    ],
)
def test_admit_regex_matches_real_phrases(text: str) -> None:
    assert _ADMIT_RE.search(text) is not None


@pytest.mark.parametrize(
    "text",
    [
        "You can cancel your direct debit at any time.",
        "Please contact us to cancel your tariff renewal.",
        "Thank you \u2014 your account has been credited.",
        "Your charges for this period (including VAT) \u00a31,525.13",
    ],
)
def test_admit_regex_rejects_non_admission_phrases(text: str) -> None:
    assert _ADMIT_RE.search(text) is None
