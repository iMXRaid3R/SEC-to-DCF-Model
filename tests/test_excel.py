"""Full recalculation of the built workbooks in real Excel (Windows + pywin32). Skipped elsewhere.
Live-price tests also need Microsoft 365 (STOCKHISTORY) and internet; they skip when the price comes back blank."""
from datetime import date, datetime

import pytest

from test_model import TICKERS, fin_for

xl_com = pytest.importorskip("xl_com", reason="needs pywin32 + Excel")
from openpyxl import Workbook  # noqa: E402

from edgar.comps import _metrics, write_comps  # noqa: E402
from edgar.xl_dcf import R  # noqa: E402
from edgar.xl_model import build_workbook  # noqa: E402


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    try:
        import win32com.client
        win32com.client.DispatchEx("Excel.Application").Quit()
    except Exception as e:
        pytest.skip(f"Excel not available: {e}")
    out = {}
    for t in TICKERS:
        path = tmp_path_factory.mktemp(t) / f"{t}.xlsx"
        build_workbook(fin_for(t), 5)[0].save(path)
        out[t] = path
    return out


@pytest.mark.parametrize("ticker", TICKERS)
def test_all_checks_ok_and_no_formula_errors(built, ticker):
    with xl_com.excel_workbook(built[ticker]) as wb:
        s = xl_com.summary(wb, price=100)
    assert s["status"] == "ALL CHECKS OK", s
    assert s["error_sheets"] == []
    per_share_gordon, per_share_exit = s["dcf"]["Value per share"]
    assert per_share_gordon > 0 and per_share_exit > 0
    assert s["sens_center"][0] == pytest.approx(per_share_gordon)
    assert s["sens_center"][1] == pytest.approx(per_share_exit)


def test_assumption_change_flows_through(built):
    """Raise ISRG year-1 revenue growth: revenue and the DCF move, and the model still balances."""
    with xl_com.excel_workbook(built["ISRG"]) as wb:
        base = xl_com.summary(wb, price=100)
        a, inc = wb.Worksheets("Drivers"), wb.Worksheets("Income Statement")
        wb.Worksheets("Inputs").Range("B7").Value = "Off"       # forecast the whole of year 1 from last year
        r_growth, r_rev = xl_com.find_row(a, "Revenue growth"), xl_com.find_row(inc, "Revenue")
        fy1 = next(j for j in range(2, 20) if inc.Cells(4, j).Value == "FY2026E")
        a.Cells(r_growth, fy1).Value = 0.30
        after = xl_com.summary(wb)
        last_fy_rev = inc.Cells(r_rev, fy1 - 3).Value      # FY2025A sits three columns left (LTM and YTD in between)
        assert inc.Cells(r_rev, fy1).Value == pytest.approx(last_fy_rev * 1.30)
    assert after["status"] == "ALL CHECKS OK"
    assert after["dcf"]["Value per share"][0] > base["dcf"]["Value per share"][0]


def test_scenario_switch_changes_the_valuation(built):
    """Bear < Base < Bull on the default spreads, and the model balances under each."""
    with xl_com.excel_workbook(built["ISRG"]) as wb:
        a, value = wb.Worksheets("Inputs"), {}
        for name in ("Bear", "Base", "Bull"):
            a.Range("B5").Value = name
            s = xl_com.summary(wb, price=100)
            assert s["status"] == "ALL CHECKS OK"
            value[name] = s["dcf"]["Value per share"][0]
    assert value["Bear"] < value["Base"] < value["Bull"]


def test_revenue_build_drives_revenue_when_on(built):
    with xl_com.excel_workbook(built["ISRG"]) as wb:
        rb, inc = wb.Worksheets("Revenue Build"), wb.Worksheets("Income Statement")
        fy1 = next(j for j in range(2, 20) if inc.Cells(4, j).Value == "FY2026E")
        rb.Cells(5, fy1 - 3).Value = 6000            # FY2025A segment revenue, $ millions
        rb.Cells(6, fy1 - 3).Value = 4000
        rb.Cells(15, fy1).Value = 0.10               # segment growth, forecast year 1
        rb.Cells(16, fy1).Value = 0.20
        rb.Range("F1").Value = "On"
        s = xl_com.summary(wb, price=100)
        assert inc.Cells(xl_com.find_row(inc, "Revenue"), fy1).Value == pytest.approx((6600 + 4800) * 1e6)
    assert s["status"] == "ALL CHECKS OK"


def _days_old(d):
    d = d.date() if isinstance(d, datetime) else d
    return (date.today() - d).days


def test_live_price_feeds_share_price_and_override_wins(built):
    with xl_com.excel_workbook(built["ISRG"]) as wb:
        live = xl_com.summary(wb)                      # no override: live close should be used
        if not isinstance(live["live_close"], float):
            pytest.skip(f"STOCKHISTORY returned {live['live_close']!r} (offline or not signed in to Microsoft 365)")
        overridden = xl_com.summary(wb, price=123.45)
        dcf_price = wb.Worksheets("DCF").Range(f"B{R['price']}").Value
    assert live["live_close"] > 0 and _days_old(live["price_date"]) <= 10
    assert live["share_price"] == live["live_close"]
    assert live["status"] == "ALL CHECKS OK"
    assert overridden["share_price"] == 123.45 and dcf_price == 123.45


def test_comps_prices_and_multiples_fill_in(tmp_path):
    path = tmp_path / "comps.xlsx"
    book = Workbook()
    write_comps([_metrics(fin_for(t)) for t in TICKERS], book.active)
    book.save(path)
    with xl_com.excel_workbook(path) as wb:
        ws = wb.Worksheets("Comps")
        hdr = {ws.Cells(4, j).Value: j for j in range(1, ws.UsedRange.Columns.Count + 1)}
        rows = range(5, 5 + len(TICKERS))
        xl_com.recalc(wb, watch=[ws.Cells(r, hdr["Live close"]) for r in rows])
        live = [ws.Cells(r, hdr["Live close"]).Value for r in rows]
        if not all(isinstance(v, float) for v in live):
            pytest.skip(f"STOCKHISTORY returned {live} (offline or not signed in to Microsoft 365)")
        for r in rows:
            assert ws.Cells(r, hdr["Share price"]).Value == ws.Cells(r, hdr["Live close"]).Value
            assert ws.Cells(r, hdr["Market cap"]).Value > 0
            assert ws.Cells(r, hdr["P / E"]).Value > 0
        median_ev_ebitda = ws.Cells(5 + len(TICKERS) + 1, hdr["EV / EBITDA"]).Value
    assert isinstance(median_ev_ebitda, float) and median_ev_ebitda > 0
