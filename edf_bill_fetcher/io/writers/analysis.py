"""Phase-2 analysis orchestrator.

Extracted from ``edf_bill_fetcher/writers/__init__.py`` during Phase 5G
of the modularization refactor (Task 6).  Hosts ``run_analysers`` —
the canonical location is now this module, not ``edf_bill_fetcher.writers``.
"""

from __future__ import annotations

from typing import Any

import pandas as pd


def run_analysers(df: pd.DataFrame) -> dict[str, Any]:
    """Run all Phase-2 detection analyses on the deduplicated

    DataFrame and return their outputs in a dict.

    The orchestrator is a thin wrapper so :func:`export_to_excel` can
    call four detectors with one line and downstream tests can
    inspect the full set without re-running each individually.

    Detectors are imported lazily inside the function body (rather than
    at module load) so callers that mock-patch the canonical-home
    module attribute (e.g. ``edf_bill_fetcher.processors.detection.detect_rebilling``)
    via :mod:`unittest.mock.patch` continue to intercept the calls.
    Binding the names at module scope would freeze the reference and
    prevent the patch from taking effect — see
    ``tests/test_integration_sap_and_recon.py::test_run_analysers_passes_evidence_df_to_detect_rebilling``
    for the contract.

    Returns:
        dict with keys ``back_billing``, ``rebilling``,
        ``meter_rollover``, ``contracts``, ``evidence_index``. The
        first four are tidy DataFrames; ``evidence_index`` is a
        ``dict[str, int]`` mapping per-row signatures to the Excel row
        on the ``EDF Evidence Report`` sheet so the analyser tabs can
        emit a ``View on Evidence Report`` hotlink.

    """
    # Lazy import — preserves the test contract for mock.patch paths.
    # The canonical homes are the detection/matching submodules, so
    # `mock.patch("edf_bill_fetcher.processors.detection.detect_rebilling")`
    # (and friends) intercept the call below.
    import edf_bill_fetcher.processors.detection as _det
    import edf_bill_fetcher.processors.matching as _matching

    return {
        "back_billing": _det.detect_back_billing(df),
        "rebilling": _det.detect_rebilling(df, evidence_df=df),
        "meter_rollover": _det.detect_meter_rollover(df),
        "contracts": _matching.infer_contracts(df),
        "evidence_index": _matching.build_evidence_index(df, header_row_offset=1),
    }
