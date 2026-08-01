from __future__ import annotations

import pandas as pd

from edf_bill_fetcher.processors.detection import detect_meter_rollover


def _row(
    invoice: str = "T-001",
    date: str = "01 Jan 2023",
    reading_type: str = "Actual",
    units_kwh: float | str = 100.0,
) -> dict:
    return {
        "Invoice #": invoice,
        "Date": date,
        "Reading": reading_type,
        "Units (kWh)": units_kwh,
        "Attachment Name": f"{invoice}.pdf",
    }


def test_empty_df_returns_empty_df() -> None:
    out = detect_meter_rollover(pd.DataFrame())
    assert out.empty
    expected_cols = {
        "Date",
        "Invoice #",
        "Prev Units (kWh)",
        "Curr Units (kWh)",
        "Delta",
        "Reading Type",
        "Notes",
    }
    assert set(out.columns) == expected_cols


def test_normal_actual_readings_not_flagged() -> None:
    df = pd.DataFrame(
        [
            _row(date="01 Jan 2023", units_kwh=300.0),
            _row(date="01 Feb 2023", units_kwh=350.0),
            _row(date="01 Mar 2023", units_kwh=400.0),
        ]
    )
    assert detect_meter_rollover(df).empty


def test_estimated_readings_skipped() -> None:
    df = pd.DataFrame(
        [
            _row(date="01 Jan 2023", reading_type="Estimated", units_kwh=300.0),
            _row(date="01 Feb 2023", reading_type="Estimated", units_kwh=-120000.0),
        ]
    )
    # All rows are Estimated -> rollover rule only trusts Actual/Smart.
    assert detect_meter_rollover(df).empty


def test_negative_unit_delta_above_threshold_emits_event() -> None:
    # Spec: rollover threshold is 99999 - 5000 = 94999. A delta below
    # -94999 indicates the meter likely rolled over near its 99,999-cap.
    df = pd.DataFrame(
        [
            _row(invoice="T-A", date="01 Jan 2023", units_kwh=95000.0),
            _row(invoice="T-B", date="01 Feb 2023", units_kwh=-120000.0),
        ]
    )
    out = detect_meter_rollover(df)
    assert len(out) == 1
    row = out.iloc[0]
    assert row["Invoice #"] == "T-B"
    assert int(row["Delta"]) < 0
    assert abs(int(row["Delta"])) > 94999


def test_small_negative_delta_not_flagged() -> None:
    # A small negative delta (-200) is a normal correction / boundary
    # adjustment -- not a rollover.
    df = pd.DataFrame(
        [
            _row(date="01 Jan 2023", units_kwh=300.0),
            _row(date="01 Feb 2023", units_kwh=-200.0),
        ]
    )
    assert detect_meter_rollover(df).empty


def test_smart_reading_triggers_rollover() -> None:
    # Smart readings count as Actual for the spec's algorithm.
    df = pd.DataFrame(
        [
            _row(invoice="SM-A", date="01 Jan 2023", reading_type="Smart", units_kwh=95000.0),
            _row(invoice="SM-B", date="01 Feb 2023", reading_type="Smart", units_kwh=-150000.0),
        ]
    )
    out = detect_meter_rollover(df)
    assert len(out) == 1
    assert out.iloc[0]["Invoice #"] == "SM-B"


def test_unknown_reading_type_skipped() -> None:
    df = pd.DataFrame(
        [
            _row(date="01 Jan 2023", reading_type="Unknown", units_kwh=300.0),
            _row(date="01 Feb 2023", reading_type="Unknown", units_kwh=-120000.0),
        ]
    )
    assert detect_meter_rollover(df).empty


def test_unparseable_units_silently_skipped() -> None:
    df = pd.DataFrame(
        [
            _row(date="01 Jan 2023", units_kwh="N/A"),
            _row(date="01 Feb 2023", units_kwh=-120000.0),
        ]
    )
    # Only the second row has a numeric Units value; no prior row to
    # pair against for delta, so nothing flagged.
    assert detect_meter_rollover(df).empty


def test_output_sorted_by_date() -> None:
    # Only MID2 has a negative delta larger than the threshold, so only
    # one event fires.
    df = pd.DataFrame(
        [
            _row(invoice="EARLY", date="01 Jan 2023", units_kwh=95000.0),
            _row(invoice="MID", date="01 Jun 2023", units_kwh=50000.0),
            _row(invoice="MID2", date="01 Aug 2023", units_kwh=-160000.0),
        ]
    )
    out = detect_meter_rollover(df)
    assert list(out["Invoice #"]) == ["MID2"]


def test_custom_threshold_param() -> None:
    # delta = curr - prev = 120000 - (-60000) = 180000.
    df = pd.DataFrame(
        [
            _row(date="01 Jan 2023", units_kwh=120000.0),
            _row(date="01 Feb 2023", units_kwh=-60000.0),
        ]
    )
    out_loose = detect_meter_rollover(df, rollover_threshold=150_000)
    assert len(out_loose) == 1
    out_strict = detect_meter_rollover(df, rollover_threshold=250_000)
    assert out_strict.empty
