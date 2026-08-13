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
