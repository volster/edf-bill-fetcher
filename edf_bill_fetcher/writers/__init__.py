"""Writers submodule — Excel sheet writers.

Placeholder for the modularization refactor (Task 5). The full writer
function extraction was deferred because each writer is hundreds of
lines and references module-level constants in ``edf_collector.py``
that would need to move with it for standalone operation.

For now, this module re-exports the writer functions from
``edf_collector`` so callers can use
``from edf_bill_fetcher.writers import write_reconciliation_sheet``
without changing their code. The compat shim is removed by Task 7.
"""

from edf_collector import (
    _write_sap_bb_matches_sheet,
    export_to_excel,
    write_back_billing_sheet,
    write_contract_history_sheet,
    write_evidence_sheet,
    write_meter_readings_sheet,
    write_rebilling_sheet,
    write_reconciliation_sheet,
    write_sap_contract_history_sheet,
    write_summary_sheet,
)

__all__ = [
    "_write_sap_bb_matches_sheet",
    "export_to_excel",
    "write_back_billing_sheet",
    "write_contract_history_sheet",
    "write_evidence_sheet",
    "write_meter_readings_sheet",
    "write_rebilling_sheet",
    "write_reconciliation_sheet",
    "write_sap_contract_history_sheet",
    "write_summary_sheet",
]
