# Option C Unlawful Charges + SAP Backbilling Position Tab Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Compute each back-billing invoice's unlawful charge from its per-sub-period `kWh × rate` slices (Option C), sum a no-double-count union total across consumption days, and add a cross-referenced "Backbilling According to SAP" tab.

**Architecture:** Per-sub-period rows are mined from new-format invoice PDFs at extraction time, serialized onto `BillingRecord` and the evidence frame, parsed back in `detect_back_billing` where the unlawful charge is recomputed per invoice and a union total is derived day-by-day. A new analyser + writer build the SAP position tab from the existing financial-transactions and reconciliation-statement rows.

**Tech Stack:** Python 3.10+, pandas, openpyxl, pdfplumber, pytest.

## Global Constraints

- Use the conda env `edf-bill-fetcher` for every command: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/python -m pytest ...`
- Run `ruff check` and `mypy edf_bill_fetcher` before committing; both must be clean.
- Do not change the 365-day legal gate (`Date − Period From > 365`) or `_prepare_analysis_frame` eligibility.
- Serialized sub-period date format is `DD/MM/YYYY` (matches `parse_to_display_date`); parse with `_safe_to_datetime` (dayfirst=True).
- The sub-period serialized column name is exactly `Sub Periods`; the per-row basis column is exactly `Sub-Period Basis`.
- Existing `TOTAL RETROSPECTIVE CHARGES — SURVIVING INVOICES` row is retained; the union row is added alongside it.
- Follow existing patterns: regexes in `processors/patterns.py`, detection in `processors/detection.py`, writers in `io/writers/`, tests in `tests/`.

---

### Task 1: `SUB_PERIOD_RE` + `extract_sub_periods` in patterns.py

**Files:**
- Modify: `edf_bill_fetcher/processors/patterns.py`
- Create: `tests/test_sub_period_extraction.py`

**Interfaces:**
- Produces: `extract_sub_periods(text: str) -> list[dict]` — each dict has keys `period_from`, `period_to` (display `DD/MM/YYYY` strings), `units_kwh` (float), `rate_p` (float), `charge` (float). Returns `[]` when no rows match.

- [ ] **Step 1: Write the failing test**

```python
from __future__ import annotations

from edf_bill_fetcher.processors.patterns import extract_sub_periods

T68_TEXT = (
    "About your charges Page 2 of 4\n"
    "02 Oct 20 - 24 Mar 21 39386YOUR READ 59129 ESTIMATED 19743 kWh 16.42p £3,241.80\n"
    "25 Mar 21 - 06 Apr 21 59129 ESTIMATED 60583 ESTIMATED 1454 kWh 16.42p £238.75\n"
    "07 Apr 21 - 31 Mar 22 60583 ESTIMATED 97767 ESTIMATED 37184 kWh 16.42p £6,105.61\n"
    "01 Apr 22 - 12 May 22 97767 ESTIMATED 1503 ESTIMATED 3736 kWh 52.00p £1,942.72\n"
    "13 May 22 - 31 Mar 23 1503 ESTIMATED 32178 ESTIMATED 30675 kWh 52.00p £15,951.00\n"
    "01 Apr 23 - 09 Aug 23 32178 ESTIMATED 42785 ESTIMATED 10607 kWh 45.92p £4,870.73\n"
)

T34_TEXT = (
    "10 Mar 17 - 30 Sep 17 72551 OUR READ 98875 YOUR READ 26324 kWh 10.88p £2,864.05\n"
    "01 Oct 17 - 08 May 18 98875YOUR READ 33348 ESTIMATED 34473 kWh 20.20p £6,963.55\n"
    "09 May 18 - 31 Dec 18 33348 ESTIMATED 64543 ESTIMATED 31195 kWh 23.50p £7,330.83\n"
    "01 Jan 19 - 03 Sep 19 64543 ESTIMATED 97262 ESTIMATED 32719 kWh 16.42p £5,372.46\n"
    "04 Sep 19 - 04 Sep 19 97262 ESTIMATED 97375 YOUR READ 113 kWh 16.42p £18.55\n"
)


def test_extract_t68_all_six_rows() -> None:
    rows = extract_sub_periods(T68_TEXT)
    assert len(rows) == 6
    assert rows[0] == {
        "period_from": "02/10/2020",
        "period_to": "24/03/2021",
        "units_kwh": 19743.0,
        "rate_p": 16.42,
        "charge": 3241.80,
    }
    assert rows[4]["units_kwh"] == 30675.0
    assert rows[4]["rate_p"] == 52.00
    assert rows[4]["charge"] == 15951.00


def test_extract_t34_one_day_row() -> None:
    rows = extract_sub_periods(T34_TEXT)
    assert len(rows) == 5
    assert rows[4]["period_from"] == "04/09/2019"
    assert rows[4]["period_to"] == "04/09/2019"
    assert rows[4]["units_kwh"] == 113.0


def test_extract_no_match_returns_empty() -> None:
    assert extract_sub_periods("no table here") == []
```

- [ ] **Step 2: Run test to verify it fails**

Run: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/python -m pytest tests/test_sub_period_extraction.py -v`
Expected: FAIL with `ImportError: cannot import name 'extract_sub_periods'`

- [ ] **Step 3: Write the implementation** — append to `edf_bill_fetcher/processors/patterns.py`:

```python
# Per-sub-period "About your charges" rows on new-format invoice PDFs.
# The middle reading tokens vary in length ("38535 OUR READ 64543 ESTIMATED"
# vs "98875YOUR READ 33348 ESTIMATED"), so the middle is captured non-greedily
# up to the trailing "<units> kWh <rate>p £<charge>" anchor.
SUB_PERIOD_RE = re.compile(
    r"(?P<pf>\d{1,2}\s+\w{3}\s+\d{2,4})\s+-\s+"
    r"(?P<pt>\d{1,2}\s+\w{3}\s+\d{2,4})\s+"
    r"(?P<mid>.*?)\s+"
    r"(?P<units>[\d,]+)\s*kWh\s+"
    r"(?P<rate>[\d.]+)p\s+"
    r"£(?P<charge>[\d,]+\.\d{2})",
    re.IGNORECASE,
)


def extract_sub_periods(text: str) -> list[dict]:
    """Return the per-sub-period charge rows from a new-format invoice body.

    Each dict carries ``period_from`` / ``period_to`` (DD/MM/YYYY display
    strings), ``units_kwh``, ``rate_p``, and ``charge`` (floats).  Rows whose
    dates fail to parse are skipped individually; returns ``[]`` when no
    rows match.
    """
    from edf_bill_fetcher.helpers.date_utils import parse_to_display_date

    out: list[dict] = []
    if not text:
        return out
    for m in SUB_PERIOD_RE.finditer(text):
        pf = parse_to_display_date(m.group("pf"))
        pt = parse_to_display_date(m.group("pt"))
        if pf == m.group("pf") or pt == m.group("pt"):
            continue  # unparseable date — skip the row
        out.append(
            {
                "period_from": pf,
                "period_to": pt,
                "units_kwh": float(m.group("units").replace(",", "")),
                "rate_p": float(m.group("rate")),
                "charge": float(m.group("charge").replace(",", "")),
            }
        )
    return out
```

Note: `parse_to_display_date` returns the input string unchanged when it cannot parse, so the `pf == m.group("pf")` guard detects unparseable dates.

- [ ] **Step 4: Run test to verify it passes**

Run: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/python -m pytest tests/test_sub_period_extraction.py -v`
Expected: PASS (3 passed)

- [ ] **Step 5: Commit**

```bash
git add edf_bill_fetcher/processors/patterns.py tests/test_sub_period_extraction.py
git commit -m "feat: add SUB_PERIOD_RE and extract_sub_periods"
```

---

### Task 2: `BillingRecord.sub_periods` + serialized `Sub Periods` column

**Files:**
- Modify: `edf_bill_fetcher/models/records.py`
- Test: `tests/test_billing_record.py`

**Interfaces:**
- Consumes: `extract_sub_periods(text) -> list[dict]` from Task 1.
- Produces: `BillingRecord.sub_periods: list[dict]` field (default `[]`); `to_dict()` emits `"Sub Periods"` as a `"; "`-joined string of `DD/MM/YYYY|DD/MM/YYYY|units|rate|charge` tokens (empty string when the list is empty).

- [ ] **Step 1: Write the failing test** — append to `tests/test_billing_record.py`:

```python
def test_sub_periods_serialized_in_to_dict() -> None:
    rec = BillingRecord(
        source="PDF",
        entry_type="New Bill",
        logic_used="New Invoice Format",
        invoice_num="T-68",
        sub_periods=[
            {
                "period_from": "02/10/2020",
                "period_to": "24/03/2021",
                "units_kwh": 19743.0,
                "rate_p": 16.42,
                "charge": 3241.80,
            },
            {
                "period_from": "25/03/2021",
                "period_to": "06/04/2021",
                "units_kwh": 1454.0,
                "rate_p": 16.42,
                "charge": 238.75,
            },
        ],
    )
    out = rec.to_dict()
    assert out["Sub Periods"] == (
        "02/10/2020|24/03/2021|19743.0|16.42|3241.8; "
        "25/03/2021|06/04/2021|1454.0|16.42|238.75"
    )


def test_sub_periods_default_empty_serialized() -> None:
    rec = BillingRecord(
        source="PDF", entry_type="New Bill", logic_used="New Invoice Format"
    )
    assert rec.to_dict()["Sub Periods"] == ""
```

- [ ] **Step 2: Run test to verify it fails**

Run: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/python -m pytest tests/test_billing_record.py -v`
Expected: FAIL (`TypeError: unexpected keyword argument 'sub_periods'`)

- [ ] **Step 3: Implement** — edit `edf_bill_fetcher/models/records.py`:

```python
@dataclass
class BillingRecord:
    ...
    cancel_rebill_admitted: bool | None = None
    sub_periods: list[dict[str, Any]] = field(default_factory=list)
```

Add `from dataclasses import dataclass, field` to the imports, and in `to_dict` add a `Sub Periods` key:

```python
    def _serialize_sub_periods(self) -> str:
        return "; ".join(
            f"{s.get('period_from', '')}|{s.get('period_to', '')}|"
            f"{s.get('units_kwh', '')}|{s.get('rate_p', '')}|"
            f"{s.get('charge', '')}"
            for s in self.sub_periods
        )

    def to_dict(self) -> dict[str, Any]:
        d = {
            ...existing keys...,
        }
        d["Sub Periods"] = self._serialize_sub_periods()
        return d
```

- [ ] **Step 4: Run test to verify it passes**

Run: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/python -m pytest tests/test_billing_record.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add edf_bill_fetcher/models/records.py tests/test_billing_record.py
git commit -m "feat: add sub_periods field and serialized Sub Periods column to BillingRecord"
```

---

### Task 3: Wire sub-period extraction into `_process_new_invoice` + evidence frame

**Files:**
- Modify: `edf_bill_fetcher/collectors/engine.py:294-382`
- Modify: `edf_bill_fetcher/io/writers/export.py:681` (col_order), `edf_bill_fetcher/io/writers/evidence.py:46` (EVIDENCE_HEADERS)
- Test: `tests/test_cancel_rebill_admitted_wiring.py` (add a case), `tests/test_io_writers_extraction.py`

**Interfaces:**
- Consumes: `extract_sub_periods` (Task 1), `BillingRecord(sub_periods=...)` (Task 2).
- Produces: new-format invoice records carry a non-empty `Sub Periods` value; the evidence frame and EDF Evidence Report sheet expose a `Sub Periods` column.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_cancel_rebill_admitted_wiring.py` (uses the file's `_engine()` helper at line 30 and mirrors `test_process_new_invoice_path_populates_cancel_rebill_admitted` at line 52):

```python
def test_process_new_invoice_carries_sub_periods() -> None:
    engine = _engine()
    body = (
        "About your charges\n"
        "02 Oct 20 - 24 Mar 21 39386YOUR READ 59129 ESTIMATED 19743 kWh 16.42p £3,241.80\n"
    )
    ok = engine._process_new_invoice(
        body,
        "PDF",
        "Test invoice",
        "09 Aug 2023",
        attachment_name="t68.pdf",
    )
    assert ok
    assert engine.records
    rec = engine.records[-1]
    assert "02/10/2020|24/03/2021|19743.0|16.42|3241.8" in rec["Sub Periods"]
```

(Verify the exact `_process_new_invoice` call signature from the existing test at line 52-90 and mirror it — it takes `(text, source_label, detail_label, fallback_date, sender="", attachment_name="")`.)

Append to `tests/test_io_writers_extraction.py` — use `write_evidence_sheet` with `_fixture_df()` extended by a `Sub Periods` column (mirror `test_evidence_writer_importable` at line 38):

```python
def test_evidence_sheet_has_sub_periods_column() -> None:
    from edf_bill_fetcher.io.writers.evidence import write_evidence_sheet

    df = _fixture_df()
    df["Sub Periods"] = ""
    df.loc[0, "Sub Periods"] = "02/10/2020|24/03/2021|19743.0|16.42|3241.8"
    ws = Workbook()
    write_evidence_sheet(ws.active, df)
    headers = [ws.active.cell(row=1, column=c).value for c in range(1, ws.active.max_column + 1)]
    assert "Sub Periods" in headers
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/python -m pytest tests/test_cancel_rebill_admitted_wiring.py tests/test_io_writers_extraction.py -v`
Expected: FAIL (missing `Sub Periods` key / missing column)

- [ ] **Step 3: Implement**

In `edf_bill_fetcher/collectors/engine.py` `_process_new_invoice`, after `fields = extract_new_invoice_fields(text)`:

```python
from edf_bill_fetcher.processors.patterns import extract_sub_periods  # extend existing import block
...
        sub_periods = extract_sub_periods(text)
```

and pass `sub_periods=sub_periods` to the `BillingRecord(...)` constructor (before `.to_dict()`).

In `edf_bill_fetcher/io/writers/evidence.py`, add `"Sub Periods"` at the end of `EVIDENCE_HEADERS`.

In `edf_bill_fetcher/io/writers/export.py`, add `"Sub Periods"` at the end of `col_order` (line ~681). Do NOT add it to `_allowed_extras` — it is a real column.

- [ ] **Step 4: Run tests to verify they pass**

Run: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/python -m pytest tests/test_cancel_rebill_admitted_wiring.py tests/test_io_writers_extraction.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add edf_bill_fetcher/collectors/engine.py edf_bill_fetcher/io/writers/evidence.py edf_bill_fetcher/io/writers/export.py tests/
git commit -m "feat: wire sub-period extraction into new-invoice pipeline and evidence report"
```

---

### Task 4: Option C unlawful charge in `detect_back_billing`

**Files:**
- Modify: `edf_bill_fetcher/processors/detection.py:187-332`
- Test: `tests/test_detect_back_billing.py` (+ `tests/test_option_c_unlawful.py` for the T68 figure)

**Interfaces:**
- Consumes: `Sub Periods` column on the input frame (Task 3), `_safe_to_datetime` (helpers/date_utils.py:114).
- Produces: `detect_back_billing(df)` output gains a `Sub-Period Basis` column (`"Sub-period × rate"` | `"Day-ratio fallback"`) and a per-row internal `_unlawful_slices: list[tuple[pd.Timestamp, pd.Timestamp, float, float]]` column (slice start, slice end, rate_p, kwh_per_day). `Unlawful Charge (£)` uses the sub-period computation when available.
- Helper produced: `_parse_sub_periods(raw: str) -> list[dict]`.

- [ ] **Step 1: Write the failing tests**

Create `tests/test_option_c_unlawful.py`:

```python
from __future__ import annotations

import pandas as pd

from edf_bill_fetcher.processors.detection import detect_back_billing

T68_SUB_PERIODS = (
    "02/10/2020|24/03/2021|19743.0|16.42|3241.8; "
    "25/03/2021|06/04/2021|1454.0|16.42|238.75; "
    "07/04/2021|31/03/2022|37184.0|16.42|6105.61; "
    "01/04/2022|12/05/2022|3736.0|52.00|1942.72; "
    "13/05/2022|31/03/2023|30675.0|52.00|15951.0; "
    "01/04/2023|09/08/2023|10607.0|45.92|4870.73"
)


def _t68_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Invoice #": "T78701920068",
                "Date": "09 Aug 2023",
                "Period From": "02 Oct 2020",
                "Period To": "09 Aug 2023",
                "Amount (£)": 32876.86,
                "Period Charge (£)": 1525.13,
                "Cancel/Rebill Admitted": True,
                "Sub Periods": T68_SUB_PERIODS,
            }
        ]
    )


def test_t68_unlawful_from_sub_periods() -> None:
    out = detect_back_billing(_t68_df())
    row = out.iloc[0]
    # fully-unlawful 02 Oct 20 -> 12 May 22 sub-periods + straddling
    # 13 May 22 -> 31 Mar 23 slice prorated at the 09/08/2022 cutoff.
    # 3241.80 + 238.75 + 6105.61 + 1942.72 + 15951.00 * 88/322
    expected = round(3241.80 + 238.75 + 6105.61 + 1942.72 + 15951.00 * (88 / 322), 2)
    assert abs(row["Unlawful Charge (£)"] - expected) < 0.01
    assert row["Sub-Period Basis"] == "Sub-period × rate"


def test_no_sub_periods_uses_day_ratio_fallback() -> None:
    df = _t68_df().drop(columns=["Sub Periods"])
    out = detect_back_billing(df)
    row = out.iloc[0]
    assert row["Sub-Period Basis"] == "Day-ratio fallback"
    assert row["Unlawful Charge (£)"] == round(1525.13 * (676 / 1041), 2)
```

Append to `tests/test_detect_back_billing.py`:

```python
def test_sub_period_basis_column_present() -> None:
    out = detect_back_billing(_row())
    assert "Sub-Period Basis" in out.columns
    assert out.iloc[0]["Sub-Period Basis"] == "Day-ratio fallback"
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/python -m pytest tests/test_option_c_unlawful.py tests/test_detect_back_billing.py -v`
Expected: FAIL (missing `Sub-Period Basis` column; unlawful uses old value)

- [ ] **Step 3: Implement**

In `edf_bill_fetcher/processors/detection.py`, add `_parse_sub_periods` and `_sub_period_unlawful_charge`, and rework `detect_back_billing`:

```python
def _parse_sub_periods(raw: object) -> list[dict]:
    """Parse the serialized ``Sub Periods`` column back into row dicts."""
    if not isinstance(raw, str) or not raw.strip():
        return []
    rows: list[dict] = []
    for token in raw.split(";"):
        token = token.strip()
        if not token:
            continue
        parts = [p.strip() for p in token.split("|")]
        if len(parts) != 5:
            continue
        pf, pt = _safe_to_datetime(parts[0]), _safe_to_datetime(parts[1])
        if pd.isna(pf) or pd.isna(pt):
            continue
        try:
            rows.append(
                {
                    "period_from": pf,
                    "period_to": pt,
                    "units_kwh": float(parts[2]),
                    "rate_p": float(parts[3]),
                    "charge": float(parts[4]),
                }
            )
        except (TypeError, ValueError):
            continue
    return rows


def _sub_period_unlawful(
    sub_periods: list[dict], cutoff: pd.Timestamp
) -> tuple[float, list[tuple]]:
    """Return (unlawful_charge, unlawful_slices).

    Day intervals are HALF-OPEN like the existing Excess Days computation
    (``(bill_date - 365 - pf).days``): a day at exactly ``cutoff`` is NOT
    unlawful, so unlawful days = ``(unlawful_to - pf).days``.  Each unlawful
    slice is ``(slice_from, slice_to, rate_p, kwh_per_day)``.  ``kwh_per_day``
    uses ``max(1, span_days)`` to survive 1-day sub-periods (e.g. T34's
    ``04 Sep 19 - 04 Sep 19`` row, span 0).
    """
    total = 0.0
    slices: list[tuple] = []
    for s in sub_periods:
        pf: pd.Timestamp = s["period_from"]
        pt: pd.Timestamp = s["period_to"]
        span_days = (pt - pf).days
        if span_days < 0:
            continue
        unlawful_to = min(pt, cutoff)
        if unlawful_to <= pf:
            continue
        rate = s["rate_p"]
        units = s["units_kwh"]
        kwh_per_day = units / max(1, span_days)
        if unlawful_to >= pt:
            charge = rate / 100.0 * units
            slices.append((pf, pt, rate, kwh_per_day))
        else:
            frac = (unlawful_to - pf).days / max(1, span_days)
            charge = rate / 100.0 * units * frac
            slices.append((pf, unlawful_to, rate, kwh_per_day))
        total += charge
    return round(total, 2), slices
```

In `detect_back_billing`, replace the `unlawful_charge` line with:

```python
        sub_periods = _parse_sub_periods(r.get("Sub Periods"))
        cutoff = bill_date_dt - pd.Timedelta(days=365)
        if sub_periods:
            unlawful_charge, unlawful_slices = _sub_period_unlawful(sub_periods, cutoff)
            sub_basis = "Sub-period × rate"
        else:
            # Day-ratio fallback (no sub-period table, e.g. KI-0014).
            # Synthesize a single unlawful slice spanning the first
            # ``excess_eff`` days of the period so the union total covers
            # the fallback row too.  The slice reproduces the day-ratio
            # unlawful charge exactly:
            #   * units present: rate = charge/units*100 p/kWh,
            #     kwh_per_day = units/days
            #   * units absent:  rate = 100.0 p/kWh (i.e. £1/kWh),
            #     kwh_per_day = charge/days (money-per-day)
            # In both cases rate/100 * kwh_per_day * excess == unlawful.
            excess_eff = min(excess, days)
            if excess_eff > 0:
                try:
                    units_f = float(str(r.get("Units (kWh)", "")).replace(",", "").strip())
                except (TypeError, ValueError):
                    units_f = 0.0
                if units_f > 0 and days > 0:
                    rate_p = charge / units_f * 100.0 if units_f else 100.0
                    kwh_per_day = units_f / days if days else 0.0
                else:
                    rate_p = 100.0
                    kwh_per_day = charge / days if days else 0.0
                unlawful_slices = [(pf, pf + pd.Timedelta(days=excess_eff), rate_p, kwh_per_day)]
            else:
                unlawful_slices = []
            unlawful_charge = round(charge * (excess_eff / days), 2) if days > 0 else 0.0
            sub_basis = "Day-ratio fallback"
```

Add `"Sub-Period Basis"` to the output `columns` list and to each row dict; add `"_unlawful_slices": unlawful_slices` to each row dict (internal column; do NOT add to `columns`).

IMPORTANT — the function's final line is currently `return out[columns]` (detection.py:332), which would DROP `_unlawful_slices`. Change it to preserve the internal column for the downstream union-total consumers:

```python
    return out[columns + ["_unlawful_slices"]]
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/python -m pytest tests/test_option_c_unlawful.py tests/test_detect_back_billing.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add edf_bill_fetcher/processors/detection.py tests/test_option_c_unlawful.py tests/test_detect_back_billing.py
git commit -m "feat: compute Option C unlawful charges from per-sub-period slices"
```

---

### Task 5: `compute_unlawful_union_total`

**Files:**
- Modify: `edf_bill_fetcher/processors/detection.py`
- Create: `tests/test_union_total.py`

**Interfaces:**
- Consumes: `detect_back_billing` output with `_unlawful_slices` per row (Task 4).
- Produces: `compute_unlawful_union_total(bb: pd.DataFrame) -> float` — the no-double-count total over the union of unlawful consumption days.

- [ ] **Step 1: Write the failing test**

```python
from __future__ import annotations

import pandas as pd

from edf_bill_fetcher.processors.detection import (
    compute_unlawful_union_total,
    detect_back_billing,
)

# T67 (bill 13 Jul 2023) recovers 15 Apr 22 - 03 Jul 23; T68 (bill 09 Aug 2023)
# recovers 02 Oct 20 - 09 Aug 23.  Their unlawful windows overlap on the days
# both invoices first recovered before each invoice's own 365-day cutoff.
T67_SUB = (
    "15/04/2022|12/05/2022|2468.0|52.00|1283.36; "
    "13/05/2022|31/03/2023|30675.0|52.00|15951.0; "
    "01/04/2023|03/07/2023|7547.0|45.92|3465.58"
)


def _df(invoice, date, pf, pt, sub_periods) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Invoice #": invoice,
                "Date": date,
                "Period From": pf,
                "Period To": pt,
                "Amount (£)": 1000.0,
                "Period Charge (£)": 100.0,
                "Cancel/Rebill Admitted": True,
                "Sub Periods": sub_periods,
            }
        ]
    )


def test_union_total_equals_sum_when_no_overlap() -> None:
    a = detect_back_billing(_df("A", "01 Mar 2023", "01 Jan 2020", "31 Dec 2020", ""))
    b = detect_back_billing(_df("B", "01 Mar 2024", "01 Jan 2023", "31 Dec 2023", ""))
    bb = pd.concat([a, b], ignore_index=True)
    total = compute_unlawful_union_total(bb)
    assert total == round(bb["Unlawful Charge (£)"].sum(), 2)


def test_union_total_does_not_double_count_overlap() -> None:
    # T67 and T68 overlap; the union must be <= the naive per-row sum and
    # strictly less when overlap exists.
    bb = pd.concat(
        [
            detect_back_billing(_df("T67", "13 Jul 2023", "15 Apr 2022", "03 Jul 2023", T67_SUB)),
            detect_back_billing(
                _df(
                    "T68",
                    "09 Aug 2023",
                    "02 Oct 2020",
                    "09 Aug 2023",
                    (
                        "02/10/2020|24/03/2021|19743.0|16.42|3241.8; "
                        "25/03/2021|06/04/2021|1454.0|16.42|238.75; "
                        "07/04/2021|31/03/2022|37184.0|16.42|6105.61; "
                        "01/04/2022|12/05/2022|3736.0|52.00|1942.72; "
                        "13/05/2022|31/03/2023|30675.0|52.00|15951.0; "
                        "01/04/2023|09/08/2023|10607.0|45.92|4870.73"
                    ),
                )
            ),
        ],
        ignore_index=True,
    )
    naive = round(bb["Unlawful Charge (£)"].sum(), 2)
    union = compute_unlawful_union_total(bb)
    assert union <= naive
    assert union < naive  # the overlapping days are counted once
```

- [ ] **Step 2: Run test to verify it fails**

Run: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/python -m pytest tests/test_union_total.py -v`
Expected: FAIL (`ImportError: cannot import name 'compute_unlawful_union_total'`)

- [ ] **Step 3: Implement** — append to `edf_bill_fetcher/processors/detection.py`:

```python
def compute_unlawful_union_total(bb: pd.DataFrame) -> float:
    """Sum unlawful consumption across all events without double counting.

    Iterates rows in the detector's bill-date order (ascending).  Each
    consumption day is claimed once, by the earliest-bill-date invoice that
    recovers it, at that invoice's rate and kWh/day.  Returns the sum over
    claimed days of ``rate_p / 100 * kwh_per_day``.
    """
    if bb is None or bb.empty:
        return 0.0
    claimed: dict[pd.Timestamp, tuple[float, float]] = {}  # day -> (rate_p, kwh/day)
    for _, row in bb.iterrows():
        for (slice_from, slice_to, rate_p, kwh_per_day) in row.get("_unlawful_slices", []):
            if not isinstance(slice_from, pd.Timestamp):
                continue
            day = slice_from
            while day < slice_to:  # half-open, matches _sub_period_unlawful
                if day not in claimed:
                    claimed[day] = (rate_p, kwh_per_day)
                day += pd.Timedelta(days=1)
    return round(sum(rate_p / 100.0 * kwh for rate_p, kwh in claimed.values()), 2)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/python -m pytest tests/test_union_total.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add edf_bill_fetcher/processors/detection.py tests/test_union_total.py
git commit -m "feat: add no-double-count union total for unlawful consumption"
```

---

### Task 6: Writer trailing union row in Back-billing Analysis

**Files:**
- Modify: `edf_bill_fetcher/io/writers/back_billing.py:297-317`
- Test: `tests/test_back_billing_sheet.py`

**Interfaces:**
- Consumes: `compute_unlawful_union_total` (Task 5).
- Produces: Back-billing Analysis tab gains a second trailing row `TOTAL UNLAWFUL CHARGES — UNION OF CONSUMPTION DAYS (no double count)` (col 10).

- [ ] **Step 1: Write the failing test** — append to `tests/test_back_billing_sheet.py` (use the file's `_open_ws()` and `_sample_df()` helpers):

```python
def test_trailing_union_total_row_written() -> None:
    from edf_bill_fetcher.io.writers.back_billing import write_back_billing_sheet
    from edf_bill_fetcher.processors.detection import detect_back_billing

    ws = _open_ws()
    df = _sample_df()
    # Give both rows real sub-period slices so the union is non-zero.
    df = df.copy()
    df["Sub Periods"] = ""
    df.loc[0, "Sub Periods"] = (
        "02/10/2020|24/03/2021|19743.0|16.42|3241.8; "
        "25/03/2021|06/04/2021|1454.0|16.42|238.75; "
        "07/04/2021|31/03/2022|37184.0|16.42|6105.61; "
        "01/04/2022|12/05/2022|3736.0|52.00|1942.72; "
        "13/05/2022|31/03/2023|30675.0|52.00|15951.0; "
        "01/04/2023|09/08/2023|10607.0|45.92|4870.73"
    )
    bb = detect_back_billing(df)
    write_back_billing_sheet(ws, bb, evidence_df=df)
    labels = [ws.cell(row=r, column=5).value for r in range(1, ws.max_row + 1)]
    assert any(
        l is not None and "UNION OF CONSUMPTION DAYS" in str(l) for l in labels
    )
```

(The existing `test_write_back_billing_sheet_total_excludes_superseded` at line 253 shows the writer-call pattern with `domination_map`.)

- [ ] **Step 2: Run test to verify it fails**

Run: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/python -m pytest tests/test_back_billing_sheet.py -v`
Expected: FAIL

- [ ] **Step 3: Implement** — edit `write_back_billing_sheet` trailing-totals block in `edf_bill_fetcher/io/writers/back_billing.py`:

```python
    if not bb.empty:
        from edf_bill_fetcher.processors.detection import compute_unlawful_union_total

        total_label = "TOTAL RETROSPECTIVE CHARGES — SURVIVING INVOICES"
        unlawful_total = 0.0
        for _, _row in bb.iterrows():
            _inv = str(_row.get("Invoice #", ""))
            if domination_map is not None and _inv in domination_map:
                continue
            unlawful_total += float(_row.get("Unlawful Charge (£)", 0.0) or 0.0)
        write_trailing_total(
            ws,
            r,
            total_label,
            [(6, total), (10, round(unlawful_total, 2))],
            5,
            17,
        )
        r += 1
        union_label = "TOTAL UNLAWFUL CHARGES — UNION OF CONSUMPTION DAYS (no double count)"
        union_total = compute_unlawful_union_total(bb)
        write_trailing_total(
            ws,
            r,
            union_label,
            [(10, round(union_total, 2))],
            5,
            17,
        )
```

- [ ] **Step 4: Run test to verify it passes**

Run: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/python -m pytest tests/test_back_billing_sheet.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add edf_bill_fetcher/io/writers/back_billing.py tests/test_back_billing_sheet.py
git commit -m "feat: add union no-double-count total row to Back-billing Analysis"
```

---

### Task 7: `analyse_sap_back_billing` analyser

**Files:**
- Modify: `edf_bill_fetcher/processors/matching.py` (add function; file already hosts SAP↔EDF matching)
- Create: `tests/test_sap_bb_position.py`

**Interfaces:**
- Consumes: `parse_sap_financial_transactions` rows (list of dicts), evidence dataframe (contains `Source == "Statement Reconciliation"` rows), and the Back-billing Analysis output df.
- Produces:
  - `analyse_sap_back_billing(sap_financial: list[dict], evidence_df: pd.DataFrame, back_billing_df: pd.DataFrame) -> dict`
  - Returns `{"events": list[dict], "reconciliation": list[dict], "summary": dict}`.
  - `events` items: `Clearing Doc #`, `Clearing Date`, `Clearing Reason`, `Net Amount (£)`, `# Rows`, `Has Credit for Consum Billing`, `Period(s)`, `Matched EDF Invoice #`.
  - `reconciliation` items: `SAP Event`, `EDF Invoice #`, `EDF Unlawful Charge (£)`, `SAP Net (£)`, `Verdict` (`Reconciled` | `Partial` | `SAP-only` | `Ours-only` | `Δ £X.XX`).

- [ ] **Step 1: Write the failing test**

```python
from __future__ import annotations

import pandas as pd

from edf_bill_fetcher.processors.matching import analyse_sap_back_billing


def _sap_row(doc, cd, reason, amount, txt) -> dict:
    return {
        "Document No.": doc,
        "Posting Date": "2023-07-13",
        "Amount": amount,
        "Transaction Text": txt,
        "Clearing Document": cd,
        "Clearing Date": "2023-08-01",
        "Clearing Reason": reason,
        "Clearing Status": "Cleared Item",
        "Statistical Key Flag": "",
    }


def _fixture() -> dict:
    sap = [
        # A real back-billing cluster: reversal credit + rebill debit.
        _sap_row("DOC-1", "CLR-100", "Reversal", -436.0, "Cr- Credit for Consum Billing"),
        _sap_row("DOC-2", "CLR-100", "Reversal", 436.0, "Dr- Consum Billing Receivable"),
        # An unrelated cluster (installment) — must be excluded.
        _sap_row("DOC-3", "CLR-999", "Automatic Clearing", 565.0, "Dr- Installment Receivable"),
    ]
    ev = pd.DataFrame(
        [
            {
                "Invoice #": "T-001",
                "Bill Date": "2023-07-13",
                "Period From": "2022-04-15",
                "Period To": "2023-07-03",
                "Period Charge (£)": 436.0,
                "Unlawful Charge (£)": 200.0,
                "_unlawful_slices": [],
            }
        ]
    )
    return {"sap": sap, "evidence": pd.DataFrame(), "bb": ev}


def test_sap_events_restricted_to_reversal_clusters() -> None:
    fx = _fixture()
    out = analyse_sap_back_billing(fx["sap"], fx["evidence"], fx["bb"])
    docs = {e["Clearing Doc #"] for e in out["events"]}
    assert docs == {"CLR-100"}  # CLR-999 excluded


def test_sap_bb_summary_totals() -> None:
    fx = _fixture()
    out = analyse_sap_back_billing(fx["sap"], fx["evidence"], fx["bb"])
    assert out["summary"]["sap_events"] == 1
    assert out["summary"]["sap_net_total"] == 0.0  # -436 + 436
```

- [ ] **Step 2: Run test to verify it fails**

Run: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/python -m pytest tests/test_sap_bb_position.py -v`
Expected: FAIL (`ImportError`)

- [ ] **Step 3: Implement** — append to `edf_bill_fetcher/processors/matching.py`:

```python
def analyse_sap_back_billing(
    sap_financial: list[dict],
    evidence_df: pd.DataFrame,
    back_billing_df: pd.DataFrame,
) -> dict:
    """Build the cross-referenced SAP back-billing position.

    SAP events are clearing-doc clusters that contain a
    ``Cr- Credit for Consum Billing`` reversal (real back-billing money
    movement).  Periods are recovered from reconciliation-statement rows
    (Source == "Statement Reconciliation") when present; otherwise blank.
    Reconciliation rows compare each SAP event's net against the matched
    EDF invoice's unlawful charge.
    """
    from collections import Counter
    from edf_bill_fetcher.writers._helpers import detect_sap_back_billing_events

    events: list[dict] = []
    reconciliation: list[dict] = []
    all_events = detect_sap_back_billing_events(sap_financial)
    back_billing_rows = (
        back_billing_df.to_dict(orient="records") if back_billing_df is not None else []
    )

    for ev in all_events:
        if not ev.has_credit_for_consum_billing:
            continue
        periods: list[str] = []
        if evidence_df is not None and not evidence_df.empty:
            recon = evidence_df[
                evidence_df.get("Source", "") == "Statement Reconciliation"
            ]
            for _, r in recon.iterrows():
                det = str(r.get("Details", ""))
                if str(r.get("Invoice #", "")) == str(ev.matched_edf_invoice or "") and det:
                    periods.append(det)
        events.append(
            {
                "Clearing Doc #": ev.clearing_doc,
                "Clearing Date": (
                    pd.Timestamp(ev.clearing_date).strftime("%Y-%m-%d")
                    if not pd.isna(ev.clearing_date)
                    else ""
                ),
                "Clearing Reason": ev.clearing_reason,
                "# Rows": len(ev.rows),
                "Net Amount (£)": ev.net_amount,
                "Has Credit for Consum Billing": "Yes" if ev.has_credit_for_consum_billing else "No",
                "Period(s)": "; ".join(dict.fromkeys(periods)) if periods else "—",
                "Matched EDF Invoice #": ev.matched_edf_invoice or "—",
            }
        )
        # Reconcile against our PDF-derived back-billing row (if any).
        match = next(
            (r for r in back_billing_rows if str(r.get("Invoice #", "")) == str(ev.matched_edf_invoice or "")),
            None,
        )
        if match is not None:
            ours = float(match.get("Unlawful Charge (£)", 0.0) or 0.0)
            if abs(ev.net_amount) < 0.01 and ours < 0.01:
                verdict = "Reconciled"
            elif abs(ours - ev.net_amount) < 0.01:
                verdict = "Reconciled"
            elif abs(ours) < 0.01:
                verdict = "SAP-only"
            elif abs(ev.net_amount) < 0.01:
                verdict = "Ours-only"
            else:
                verdict = f"Δ £{ours - ev.net_amount:,.2f}"
            reconciliation.append(
                {
                    "SAP Event": ev.clearing_doc,
                    "EDF Invoice #": match.get("Invoice #", ""),
                    "EDF Unlawful Charge (£)": ours,
                    "SAP Net (£)": ev.net_amount,
                    "Verdict": verdict,
                }
            )
        else:
            reconciliation.append(
                {
                    "SAP Event": ev.clearing_doc,
                    "EDF Invoice #": ev.matched_edf_invoice or "—",
                    "EDF Unlawful Charge (£)": 0.0,
                    "SAP Net (£)": ev.net_amount,
                    "Verdict": "Ours-only" if ev.matched_edf_invoice is None else "Partial",
                }
            )

    summary = {
        "sap_events": len(events),
        "sap_net_total": round(sum(e["Net Amount (£)"] for e in events), 2),
        "reconciled": sum(1 for r in reconciliation if r["Verdict"] == "Reconciled"),
    }
    return {"events": events, "reconciliation": reconciliation, "summary": summary}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/python -m pytest tests/test_sap_bb_position.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add edf_bill_fetcher/processors/matching.py tests/test_sap_bb_position.py
git commit -m "feat: add cross-referenced SAP back-billing position analyser"
```

---

### Task 8: `write_sap_back_billing_position_sheet` writer

**Files:**
- Modify: `edf_bill_fetcher/io/writers/sap.py`
- Test: `tests/test_export_sap_back_billing_sheets.py`

**Interfaces:**
- Consumes: `analyse_sap_back_billing` output dict (Task 7).
- Produces: `write_sap_back_billing_position_sheet(wb, result: dict, account: str = "") -> Worksheet` — sheet titled `Backbilling According to SAP` with three sections: summary banner, events table, reconciliation table.

- [ ] **Step 1: Write the failing test** — append to `tests/test_export_sap_back_billing_sheets.py`:

```python
def test_write_sap_back_billing_position_sheet() -> None:
    from edf_bill_fetcher.io.writers.sap import write_sap_back_billing_position_sheet

    wb = Workbook()
    result = {
        "events": [
            {
                "Clearing Doc #": "CLR-100",
                "Clearing Date": "2023-08-01",
                "Clearing Reason": "Reversal",
                "# Rows": 2,
                "Net Amount (£)": 0.0,
                "Has Credit for Consum Billing": "Yes",
                "Period(s)": "—",
                "Matched EDF Invoice #": "T-001",
            }
        ],
        "reconciliation": [
            {
                "SAP Event": "CLR-100",
                "EDF Invoice #": "T-001",
                "EDF Unlawful Charge (£)": 0.0,
                "SAP Net (£)": 0.0,
                "Verdict": "Reconciled",
            }
        ],
        "summary": {"sap_events": 1, "sap_net_total": 0.0, "reconciled": 1},
    }
    ws = write_sap_back_billing_position_sheet(wb, result)
    assert ws.title == "Backbilling According to SAP"
    col_a = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
    assert any("CLR-100" in str(v) for v in col_a)
    assert any("Reconciled" in str(v) for v in col_a)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/python -m pytest tests/test_export_sap_back_billing_sheets.py -v`
Expected: FAIL (`ImportError`)

- [ ] **Step 3: Implement** — append to `edf_bill_fetcher/io/writers/sap.py`:

```python
def write_sap_back_billing_position_sheet(
    wb: openpyxl.Workbook,
    result: dict,
    account: str = "",
) -> Worksheet:
    """Render the 'Backbilling According to SAP' cross-referenced position.

    Three sections: title banner (with event-count summary), the SAP
    back-billing events table (reversal-containing clusters), and the
    reconciliation table against our PDF-derived Back-billing Analysis.
    """
    ws = wb.create_sheet(title="Backbilling According to SAP")
    ORANGE = "FE5716"
    NAVY = "10367A"

    summary = result.get("summary", {})
    title = (
        "BACKBILLING ACCORDING TO SAP  |  Account {acc}  |  "
        "{n} event(s)  |  SAP net total £{net:,.2f}  |  "
        "{rec} reconciled"
    ).format(
        acc=account or "(no account)",
        n=summary.get("sap_events", 0),
        net=summary.get("sap_net_total", 0.0),
        rec=summary.get("reconciled", 0),
    )
    _write_sap_header_row(ws, 1, [title])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    c1 = ws.cell(row=1, column=1)
    c1.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    c1.fill = PatternFill("solid", start_color=ORANGE)

    events = result.get("events", [])
    ev_cols = [
        "Clearing Doc #",
        "Clearing Date",
        "Clearing Reason",
        "# Rows",
        "Net Amount (£)",
        "Has Credit for Consum Billing",
        "Period(s)",
        "Matched EDF Invoice #",
    ]
    _write_sap_header_row(ws, 3, ev_cols)
    r = 4
    for i, ev in enumerate(events):
        for j, col in enumerate(ev_cols, start=1):
            cell = ws.cell(row=r, column=j, value=ev.get(col, ""))
            cell.font = Font(name="Calibri", size=10)
            cell.border = CELL_BORDER
            if i % 2 == 0:
                cell.fill = PatternFill("solid", start_color=SAP_BB_SUMMARY_FILL_PAIR[0])
        r += 1

    r += 1
    rec_cols = [
        "SAP Event",
        "EDF Invoice #",
        "EDF Unlawful Charge (£)",
        "SAP Net (£)",
        "Verdict",
    ]
    _write_sap_header_row(ws, r, rec_cols)
    r += 1
    for i, rec in enumerate(result.get("reconciliation", [])):
        for j, col in enumerate(rec_cols, start=1):
            cell = ws.cell(row=r, column=j, value=rec.get(col, ""))
            cell.font = Font(name="Calibri", size=10)
            cell.border = CELL_BORDER
            if i % 2 == 0:
                cell.fill = PatternFill("solid", start_color=SAP_BB_DETAIL_FILL_PAIR[0])
        r += 1

    for idx, width in enumerate((18, 14, 22, 10, 16, 26, 40, 20), start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = width
    return ws
```

- [ ] **Step 4: Run test to verify it passes**

Run: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/python -m pytest tests/test_export_sap_back_billing_sheets.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add edf_bill_fetcher/io/writers/sap.py tests/test_export_sap_back_billing_sheets.py
git commit -m "feat: add Backbilling According to SAP position sheet writer"
```

---

### Task 9: Wire the new tab into `export_to_excel` + reorder

**Files:**
- Modify: `edf_bill_fetcher/io/writers/export.py:1761-1832` (SAP block) and `:295` (`_reorder_sheets`)
- Test: `tests/test_export_sap_back_billing_sheets.py` / `tests/test_io_writers_extraction.py`

**Interfaces:**
- Consumes: `analyse_sap_back_billing` (Task 7), `write_sap_back_billing_position_sheet` (Task 8), `analyses["back_billing"]` (existing wiring).
- Produces: workbook contains a `Backbilling According to SAP` tab, placed after `SAP ↔ EDF Matched Events` in the severity-led order.

- [ ] **Step 1: Write the failing test** — append to `tests/test_export_sap_back_billing_sheets.py` (or the existing full-export fixture test):

```python
def test_full_export_includes_sap_bb_position_tab(export_fixture) -> None:
    assert "Backbilling According to SAP" in export_fixture.wb.sheetnames
```

(Reuse the file's existing full-export fixture; if none exists, extend the fixture used by `test_io_writers_extraction.py`.)

- [ ] **Step 2: Run test to verify it fails**

Run: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/python -m pytest tests/test_export_sap_back_billing_sheets.py -v`
Expected: FAIL (tab missing)

- [ ] **Step 3: Implement**

In `export.py`, inside the `if sap_financial:` block, after `write_sap_back_billing_sheets(...)`:

```python
            if config.get("generate_reconciliation_sheet", True):
                from edf_bill_fetcher.io.writers.sap import write_sap_back_billing_position_sheet
                from edf_bill_fetcher.processors.matching import analyse_sap_back_billing

                sap_bb_position = analyse_sap_back_billing(
                    sap_financial, dfc, analyses.get("back_billing")
                )
                write_sap_back_billing_position_sheet(
                    wb, sap_bb_position, account=account_label
                )
```

Add `"Backbilling According to SAP"` to `_SEVERITY_LED_ORDER` right after `"SAP ↔ EDF Matched Events"` (export.py:295 list).

- [ ] **Step 4: Run test to verify it passes**

Run: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/python -m pytest tests/test_export_sap_back_billing_sheets.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add edf_bill_fetcher/io/writers/export.py tests/
git commit -m "feat: wire Backbilling According to SAP tab into export pipeline"
```

---

### Task 10: Full CI gate + regenerate workbook

**Files:**
- All modified/created files.

- [ ] **Step 1: Run the full test suite**

Run: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/python -m pytest -q`
Expected: all tests pass (1411 + new tests; existing 9 skips unchanged). If an existing test asserts the old Back-billing Analysis trailing-total layout or old unlawful numbers, update it to the new union-row layout.

- [ ] **Step 2: Run linters and type checks**

Run: `cd /home/matthew/ai/opencode/edf-bill-fetcher && /home/matthew/miniconda3/envs/edf-bill-fetcher/bin/ruff check edf_bill_fetcher tests`
Run: `/home/matthew/miniconda3/envs/edf-bill-fetcher/bin/mypy edf_bill_fetcher`
Expected: both clean.

- [ ] **Step 3: Regenerate the workbook and verify**

Run the export from `scratch/input` (same command used for the v4 workbook) into `scratch/output/refactor 2/EDF_Dispute_Evidence_refactor_2026-08-13_1.xlsx` (or a new dated file). Then verify:
- Back-billing Analysis trailing block has BOTH total rows (surviving + union).
- `Unlawful Charge (£)` per row reflects sub-period computation (T68 ≈ £15,888, not £990.38).
- `Backbilling According to SAP` tab exists with events + reconciliation.
- Union total is < naive per-row sum and compares against Claude's £35,884.73 (document the delta in the tab's notes or the PR description).

Use a short script to dump the two total rows and the union figure for confirmation.

- [ ] **Step 4: Commit**

```bash
git add -A
git commit -m "chore: regenerate workbook with Option C unlawful charges and SAP position tab"
```

---

## Self-Review Notes

- **Spec coverage:** §3.1→Tasks 1-3; §3.2→Task 4; §3.3→Tasks 5-6; §3.4→Tasks 7-9; §3.5 error handling folded into Task 1 (date guard), Task 4 (div-by-zero, inverted periods), Task 5 (missing slices); §4 testing→Task 10.
- **Type consistency:** `_unlawful_slices` produced in Task 4 as `list[tuple[pd.Timestamp, pd.Timestamp, float, float]]`, consumed in Tasks 5-6 with the same shape; `Sub Periods` serialization format (`; `-joined `|` tokens) identical across Tasks 2-4; `analyse_sap_back_billing` returns `{"events", "reconciliation", "summary"}` consumed in Tasks 8-9.
- **Placeholder scan:** every step has concrete code or an exact command; no TBDs.
