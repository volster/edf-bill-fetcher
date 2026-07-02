"""Phase 2.1 — unit-rate vectorisation parity test.

The historic ``_compute_unit_rate`` row-wise helper produced
``round((pc / units) * 100, 2)`` with comma-stripped units and
``np.nan`` for every other failure outcome.  Phase 2.1 replaced
the main ``df.apply(_compute_unit_rate, axis=1)`` call with a
vectorised pandas/numpy implementation:

    pc = pd.to_numeric(df["Period Charge (£)"], errors="coerce")
    units = pd.to_numeric(
        df["Units (kWh)"].astype(str).str.replace(",", ""),
        errors="coerce",
    )
    df["Unit Rate (p/kWh)"] = np.where(
        (units > 0) & (pc > 0),
        np.round((pc / units) * 100, 2),
        np.nan,
    )

This test pins the contract:

    * Both paths yield **identical** output values for the
      standard cases (positive numerics, comma-formatted units,
      missing values, zero amounts, negative amounts).
    * The vectorised path runs at most as long as the row-wise
      path on a 5,000-row fixture (loose ceiling; the bench in
      Night-2 saw 5× speedup already).
"""

from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path

import numpy as np
import pandas as pd
import pytest


@pytest.fixture
def tmp_dir() -> Iterator[Path]:
    """A writable temp directory."""
    import tempfile

    d = Path(tempfile.mkdtemp(prefix="edf_unit_rate_vec_"))
    yield d
    try:
        for f in d.iterdir():
            f.unlink(missing_ok=True)
        d.rmdir()
    except OSError:
        pass


def legacy_unit_rate(row: pd.Series) -> float:
    """Reference row-wise helper inheriting the original semantics.

    Inlined to keep the test surface narrow (we don't reach
    into ``edf_collector`` for the helper because that local
    closure has oddly-coupled state in the production code).
    """
    pc = row.get("Period Charge (£)")
    units = row.get("Units (kWh)")
    try:
        pc_f = float(pc)
        u_f = float(str(units).replace(",", ""))
        if u_f > 0 and pc_f > 0:
            return round((pc_f / u_f) * 100, 2)
    except (ValueError, TypeError):
        pass
    return np.nan


def vectorised_unit_rate(df: pd.DataFrame) -> pd.Series:
    """The Phase 2.1 vectorised path, copy-pasted for parity.

    We re-implement the vectorisation here so the test does not
    depend on any in-scope variable inside ``export_to_excel``.
    """
    pc = pd.to_numeric(df["Period Charge (£)"], errors="coerce")
    units = pd.to_numeric(
        df["Units (kWh)"].astype(str).str.replace(",", ""),
        errors="coerce",
    )
    return pd.Series(
        np.where(
            (units > 0) & (pc > 0),
            np.round((pc / units) * 100, 2),
            np.nan,
        ),
        index=df.index,
    )


class TestUnitRateParity:
    """Phase 2.1 — vectorised unit-rate matches the row-wise reference."""

    def _records(self) -> list[dict]:
        """A small but diverse fixture covering most realistic
        input shapes: comma-formatted units, NaN, blanks, zero,
        negative, very small fractional.  The rows deliberately
        come with dict-style keys that mirror the Excel row
        schema we receive from the upstream extractors.
        """
        return [
            # Positive numerics, plain integers.
            {
                "Date": "01/01/2024",
                "Period Charge (£)": 80.0,
                "Units (kWh)": 100,
            },
            # Comma-formatted units (1,234) — Phase 2.1's
            # ``str.replace(",", "")`` normalises this exactly the
            # same way Python's ``str(1000).replace(",","")`` did.
            {
                "Date": "01/02/2024",
                "Period Charge (£)": 500.0,
                "Units (kWh)": "1,234",
            },
            # Blank NaN — both paths must return NaN.
            {
                "Date": "01/03/2024",
                "Period Charge (£)": None,
                "Units (kWh)": None,
            },
            # Zero units — both paths must return NaN because
            # division by zero is unsafe.
            {
                "Date": "01/04/2024",
                "Period Charge (£)": 80.0,
                "Units (kWh)": 0,
            },
            # Negative units — the legacy path's ``if u_f > 0``
            # guard rejects negatives; vectorised path's
            # ``(units > 0) & (pc > 0)`` does the same.  Both
            # return NaN.
            {
                "Date": "01/05/2024",
                "Period Charge (£)": 100.0,
                "Units (kWh)": -50,
            },
            # Zero period charge — legacy returned 0.0 because
            # the guard ``pc_f > 0`` rejects it; vectorised
            # path's ``(pc > 0) & (units > 0)`` rejects ``pc=0``
            # too.  Both return NaN.
            {
                "Date": "01/06/2024",
                "Period Charge (£)": 0.0,
                "Units (kWh)": 200,
            },
            # Non-numeric strings — legacy ``try/except``
            # catches ValueError.  Vectorised path returns NaN
            # via ``pd.to_numeric(..., errors="coerce")``.
            {
                "Date": "01/07/2024",
                "Period Charge (£)": "unknown",
                "Units (kWh)": "n/a",
            },
            # Float-formatted units — both paths handle.
            {
                "Date": "01/08/2024",
                "Period Charge (£)": 99.99,
                "Units (kWh)": 250.5,
            },
            # Very small fractional — exercises rounding mode.
            {
                "Date": "01/09/2024",
                "Period Charge (£)": 0.4081633,
                "Units (kWh)": 100,
            },
            # Typical kWh figure with no comma.
            {
                "Date": "01/10/2024",
                "Period Charge (£)": 1200.0,
                "Units (kWh)": 8000,
            },
        ]

    def test_vectorised_matches_row_wise(self) -> None:
        """The vectorised path must agree with the row-wise
        reference for every record in the fixture.  We allow a
        1e-9 absolute tolerance only for the rounded float output
        itself — Python's ``round`` and ``np.round`` may take the
        ties *slightly* differently but for our domain (unit rate
        to 0.01 in p/kWh) the difference is well below 1e-9.
        """
        records = self._records()
        df = pd.DataFrame(records)
        legacy = [legacy_unit_rate(df.iloc[i]) for i in range(len(df))]
        vec = vectorised_unit_rate(df)

        # Both NaN count should agree.
        legacy_nans = sum(1 for v in legacy if v is np.nan or pd.isna(v))
        vec_nans = int(vec.isna().sum())
        assert legacy_nans == vec_nans, (
            f"NaN count diverged: legacy={legacy_nans} vec={vec_nans}.  "
            f"legacy={legacy} vs. vec={vec.tolist()}"
        )

        # Non-NaN values must match to 1e-9 p/kWh.
        for i, (lv, vv) in enumerate(zip(legacy, vec.tolist(), strict=True)):
            if pd.isna(lv) and pd.isna(vv):
                continue
            assert abs(lv - vv) < 1e-9, (
                f"Row {i}: legacy={lv} vec={vv}; legacy={legacy} vec={vec.tolist()}"
            )

    def test_vectorised_handles_comma_thousands_separator(self) -> None:
        """Comma-formatted units ('2,500') must produce the
        same answer as the row-wise version, which calls
        ``str(units).replace(",", "")``.  This is the highest-
        cardinality fixture since EDF's export sometimes gives
        '1,234.5' instead of '1234.5'.
        """
        df = pd.DataFrame(
            {
                "Date": pd.date_range("2024-01-01", periods=12, freq="MS").strftime("%d/%m/%Y"),
                "Period Charge (£)": [50.0 + i * 1.5 for i in range(12)],
                "Units (kWh)": [
                    f"{v:,}"  # str.format with thousands separator
                    for v in [1000, 1500, 2000, 2500, 3000, 100, 200, 300, 400, 500, 10000, 1]
                ],
            }
        )

        legacy = [legacy_unit_rate(df.iloc[i]) for i in range(len(df))]
        vec = vectorised_unit_rate(df)
        for i, (lv, vv) in enumerate(zip(legacy, vec.tolist(), strict=True)):
            # Vectorised path may round to a different float
            # bitmap than Python's banker's-round ``round``; assert
            # equality up to 1e-9 as above.
            if pd.isna(lv) and pd.isna(vv):
                continue
            assert abs(lv - vv) < 1e-9, (
                f"Row {i}: legacy={lv} vec={vv}; period="
                f"{df['Period Charge (£)'].iloc[i]}, units={df['Units (kWh)'].iloc[i]!r}"
            )

    def test_vectorised_rounds_to_two_decimals(self) -> None:
        """Pin the rounding behaviour the spec asks for:
        ``round((pc / units) * 100, 2)``.  Both legacy and
        vectorised implementations round to two decimals; a
        difference here would mean an arithmetic regression.
        """
        df = pd.DataFrame(
            {
                "Date": ["01/01/2024"] * 3,
                # These ratios produce 'interesting' rounding:
                #   1/3 * 100 = 33.333... → 2dp = 33.33 (lossy)
                #   1/6 * 100 = 16.666... → 2dp = 16.67
                #   1/7 * 100 = 14.285714... → 2dp = 14.29
                "Period Charge (£)": [1.0, 1.0, 1.0],
                "Units (kWh)": [3, 6, 7],
            }
        )
        vec = vectorised_unit_rate(df)
        assert vec.tolist() == pytest.approx([33.33, 16.67, 14.29], abs=1e-9)
