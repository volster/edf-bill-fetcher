"""Tests that color constants are importable from the helpers.theme submodule.

All tests are RED at Phase 0 because ``edf_bill_fetcher.helpers.theme``
does not yet exist.  They will turn GREEN once the submodule is created
during modularization.
"""

from __future__ import annotations


def test_edf_navy_importable() -> None:
    from edf_bill_fetcher.helpers.theme import EDF_NAVY

    assert EDF_NAVY == "#10367A"


def test_edf_orange_importable() -> None:
    from edf_bill_fetcher.helpers.theme import EDF_ORANGE

    assert EDF_ORANGE == "#FE5716"


def test_medium_grey_importable() -> None:
    from edf_bill_fetcher.helpers.theme import MEDIUM_GREY

    assert MEDIUM_GREY == "#666666"


def test_dup_grey_importable() -> None:
    from edf_bill_fetcher.helpers.theme import DUP_GREY

    assert DUP_GREY is not None


def test_edf_offwhite_importable() -> None:
    from edf_bill_fetcher.helpers.theme import EDF_OFFWHITE

    assert EDF_OFFWHITE == "#F5F5F5"


def test_orange_importable() -> None:
    from edf_bill_fetcher.helpers.theme import ORANGE

    assert ORANGE == "FE5716"


def test_navy_blue_importable() -> None:
    from edf_bill_fetcher.helpers.theme import NAVY_BLUE

    assert NAVY_BLUE == "10367A"


def test_sap_bb_summary_fill_pair_importable() -> None:
    from edf_bill_fetcher.helpers.theme import SAP_BB_SUMMARY_FILL_PAIR

    assert SAP_BB_SUMMARY_FILL_PAIR == ("EFF4FB", "ffffff")


def test_sap_bb_detail_fill_pair_importable() -> None:
    from edf_bill_fetcher.helpers.theme import SAP_BB_DETAIL_FILL_PAIR

    assert SAP_BB_DETAIL_FILL_PAIR == ("F8FAFC", "ffffff")


def test_sap_bb_medium_border_importable() -> None:
    from edf_bill_fetcher.helpers.theme import SAP_BB_MEDIUM_BORDER

    assert SAP_BB_MEDIUM_BORDER is not None


def test_cell_border_importable() -> None:
    from edf_bill_fetcher.helpers.theme import CELL_BORDER

    assert CELL_BORDER is not None
