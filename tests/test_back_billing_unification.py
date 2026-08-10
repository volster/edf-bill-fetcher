"""Structural guard: the writer module must re-export the canonical detector.

``edf_bill_fetcher.io.writers.back_billing`` used to carry byte-identical
copies of ``_assess_reason`` / ``_pull_period_charge`` / ``detect_back_billing``
from ``processors.detection``. They are now re-exported so there is exactly
one definition in the codebase. If a local copy is ever re-introduced, the
``__module__`` assertions below fail.
"""

from __future__ import annotations


def test_back_billing_writer_reexports_canonical_detect() -> None:
    from edf_bill_fetcher.io.writers.back_billing import detect_back_billing

    assert detect_back_billing.__module__ == "edf_bill_fetcher.processors.detection"


def test_back_billing_writer_reexports_canonical_helpers() -> None:
    from edf_bill_fetcher.io.writers.back_billing import (
        _assess_reason,
        _pull_period_charge,
    )

    assert _assess_reason.__module__ == "edf_bill_fetcher.processors.detection"
    assert _pull_period_charge.__module__ == "edf_bill_fetcher.processors.detection"
