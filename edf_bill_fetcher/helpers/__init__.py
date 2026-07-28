"""Utility helpers for the EDF evidence workbook."""

from __future__ import annotations

from edf_bill_fetcher.helpers.date_utils import (
    _safe_to_datetime,
    build_evidence_trail,
    completeness_score,
    compute_ema,
    compute_momentum,
    compute_rolling_stats,
    parse_to_display_date,
    parse_to_sort_date,
    to_excel_date,
)
from edf_bill_fetcher.helpers.excel_utils import (
    _TEXT_SUPPRESSION_QUEUE,
    CELL_BORDER,
    build_sap_row_index_map,
    hcell,
    money,
    num,
    open_pdf_hyperlink_cell,
    section_hdr,
    set_column_widths_from_spec,
    suppress_text_warning,
    suppress_text_warnings_post_save,
    text,
)
from edf_bill_fetcher.helpers.formatting import (
    _amalgamate_cluster,
    _apply_amalgamate_to_kept_frame,
    _is_populated,
    account_number_matches,
    apply_currency_format,
    apply_int_format,
)
from edf_bill_fetcher.helpers.pdf_utils import (
    _htm_excerpt,
    extract_admit_phrase,
    legal_context,
    parse_htm_account_history,
    slice_pdf_pages,
)

__all__ = [
    "_is_populated",
    "_amalgamate_cluster",
    "_apply_amalgamate_to_kept_frame",
    "account_number_matches",
    "apply_currency_format",
    "apply_int_format",
    "build_evidence_trail",
    "completeness_score",
    "compute_ema",
    "compute_momentum",
    "compute_rolling_stats",
    "parse_to_sort_date",
    "_safe_to_datetime",
    "parse_to_display_date",
    "to_excel_date",
    "_TEXT_SUPPRESSION_QUEUE",
    "CELL_BORDER",
    "build_sap_row_index_map",
    "hcell",
    "money",
    "num",
    "open_pdf_hyperlink_cell",
    "section_hdr",
    "set_column_widths_from_spec",
    "suppress_text_warning",
    "suppress_text_warnings_post_save",
    "text",
]
