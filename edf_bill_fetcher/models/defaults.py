"""Shared extraction defaults for every interactive and automated entry point."""

from edf_bill_fetcher.models.config import ConfigDict

EXTRACTION_DEFAULTS: ConfigDict = {
    "use_anchors": True,
    "use_large": True,
    "use_reading_classification": True,
    "use_pdf_fields": True,
    "use_acc_filter": False,
    "filter_below": False,
    "save_filtered": True,
    "use_dedup": True,
    "save_dups": True,
    "amalgamate_duplicates": False,
    "use_domain_filter": True,
    "save_evidence_files": True,
    "scan_sap_dumps": True,
    "generate_reconciliation_sheet": True,
    "acc_num": "",
    "domain_filter": "edfenergy.com",
    "min_amount": 50.0,
    "analysis_min": 0.0,
}
