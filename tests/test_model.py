"""
Offline tests on saved SEC data (tests/fixtures/*.json.gz, refreshed by copying from ./cache).
The Excel tests recalculate the built workbooks in real Excel and are skipped where Excel isn't available.

Run:  uv run --with pytest --with requests --with pandas --with openpyxl --with pywin32 pytest tests
"""
from pathlib import Path

import pandas as pd
import pytest

from edgar.sec import load_json
from edgar.transform import load_financials
from edgar.xl_model import build_workbook

FIXTURES = Path(__file__).parent / "fixtures"
TICKERS = ["AAPL", "COST", "KO", "ISRG"]


def fin_for(ticker, years=5):
    facts = load_json(FIXTURES / f"{ticker}.json.gz")
    return load_financials(facts, ticker, facts.get("entityName", ticker), years)


@pytest.fixture(scope="module")
def fins():
    return {t: fin_for(t) for t in TICKERS}


def fy_value(fin, label, year_end):
    return fin.values[label].get(pd.Timestamp(year_end))


# ---- historical values tie to the 10-Ks ----------------------------------
def test_aapl_revenue_ties(fins):
    f = fins["AAPL"]
    assert fy_value(f, "Revenue", "2022-09-24") == 394_328e6
    assert fy_value(f, "Revenue", "2024-09-28") == 391_035e6
    assert fy_value(f, "Revenue", "2025-09-27") == 416_161e6


def test_isrg_revenue_and_da_tie(fins):
    f = fins["ISRG"]
    assert fy_value(f, "Revenue", "2024-12-31") == 8_352.1e6
    assert fy_value(f, "D&A", "2025-12-31") == 613.2e6          # Depreciation + amortization (summed tags)


def test_ltm_only_when_a_newer_10q_exists(fins):
    assert fins["ISRG"].ltm_end == pd.Timestamp("2026-09-30")
    assert fins["KO"].ltm_end is not None
    assert fins["AAPL"].ltm_end is None and fins["AAPL"].ltm == {}


def test_isrg_ltm_revenue_is_fy_plus_ytd_minus_prior_ytd(fins):
    ltm = fins["ISRG"].ltm["Revenue"]
    assert fy_value(fins["ISRG"], "Revenue", "2025-12-31") < ltm < 1.3 * fy_value(fins["ISRG"], "Revenue", "2025-12-31")


@pytest.mark.parametrize("ticker", TICKERS)
def test_reported_balance_sheet_balances(fins, ticker):
    f = fins[ticker]
    for end in f.cols:
        ta, tle = fy_value(f, "Total Assets", end), fy_value(f, "Total Liabilities & Equity", end)
        if ta is not None and tle is not None:
            assert ta == tle


@pytest.mark.parametrize("ticker", TICKERS)
def test_five_fiscal_years_and_raw_tab_limited_to_them(fins, ticker):
    f = fins[ticker]
    assert len(f.cols) == 5
    shown = {c.date() for c in f.cols} | ({f.ltm_end.date()} if f.ltm_end is not None else set())
    assert set(f.sources["Period End"]) <= shown


def test_shares_outstanding_found(fins):
    for t in TICKERS:
        shares, _, how = fins[t].shares_out
        assert shares and shares > 1e8, (t, how)


# ---- workbook structure ---------------------------------------------------
def test_workbook_sheets_and_columns(fins):
    wb, m = build_workbook(fins["ISRG"], 5)
    assert wb.sheetnames == ["Summary", "Inputs", "Drivers", "Revenue Build", "Income Statement", "Balance Sheet",
                             "Cash Flow", "DCF", "Checks & Metrics", "Notes", "Mapping", "Raw", "Calc", "_inputs"]
    assert [ws.title for ws in wb if ws.sheet_state == "hidden"] == ["Calc", "_inputs"]
    ws = wb["Income Statement"]
    headers = [ws.cell(row=4, column=j).value for j in range(2, 14)]
    assert headers[:5] == ["FY2021A", "FY2022A", "FY2023A", "FY2024A", "FY2025A"]
    assert headers[5:7] == ["LTM Sep-26", "YTD 9M"]
    assert headers[7:] == ["FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E"]


def test_year_one_uses_year_to_date_when_a_newer_10q_exists(fins):
    f = fins["ISRG"]
    assert f.ytd_months == 9 and f.values["Revenue"][pd.Timestamp("2025-12-31")] < f.ltm["Revenue"]
    assert 0 < f.ytd["Revenue"] < f.ltm["Revenue"]
    wb, m = build_workbook(f, 5)
    g = m.grid
    rev = wb["Income Statement"].cell(row=g.rows["Revenue"][1], column=g.fcst[0].j).value
    assert f"{g.ytd.L}{g.rows['Revenue'][1]}" in rev and "Use" not in rev and 'Inputs!$B$7="On"' in rev
    later = wb["Income Statement"].cell(row=g.rows["Revenue"][1], column=g.fcst[1].j).value
    assert g.ytd.L + str(g.rows["Revenue"][1]) not in later            # only year 1 is anchored to actuals
    assert fins["AAPL"].ytd == {} and build_workbook(fins["AAPL"], 5)[1].grid.ytd is None


def test_peers_add_a_comps_tab_linked_to_the_exit_multiple(fins):
    from edgar.comps import _metrics
    from edgar.xl_dcf import R
    wb, m = build_workbook(fins["ISRG"], 5, [_metrics(fins[t]) for t in ("KO", "COST")])
    assert "Comps" in wb.sheetnames
    comps = wb["Comps"]
    assert [comps.cell(row=r, column=1).value for r in (5, 6, 7)] == ["ISRG", "KO", "COST"]
    from edgar.layout import INPUT_ROWS
    assert "Comps!" in wb["Inputs"][f"B{INPUT_ROWS['multiple']}"].value
    assert "Inputs!" in wb["DCF"][f"B{R['multiple']}"].value            # the DCF tab only shows the setting


# ---- inputs survive a rebuild ---------------------------------------------
def test_inputs_carry_over_to_a_rebuild_with_a_different_layout(fins, tmp_path):
    from edgar.inputs import carry_over
    from edgar.layout import INPUT_ROWS as I
    path = tmp_path / "ISRG.xlsx"
    wb, m = build_workbook(fins["ISRG"], 5)
    inp, drv, rb = wb["Inputs"], wb["Drivers"], wb["Revenue Build"]
    g = m.grid
    growth = g.rows["Revenue growth"][1]
    inp[f"B{I['Scenario']}"] = "Bull"
    drv.cell(row=g.scen_rows["Gross margin"], column=8).value = 0.70        # Bull, year 1
    drv.cell(row=growth, column=g.fcst[1].j).value = 0.25                   # FY2027E override
    drv.cell(row=growth, column=g.fcst[0].j).value = 0.11                   # FY2026E override
    rb.cell(row=5, column=g.last_fy.j).value = 6000                         # FY2025A segment revenue
    rb["A5"] = "Instruments & accessories"
    inp[f"B{I['beta']}"] = 1.3
    inp[f"B{I['RSUs']}"] = 4_000_000
    wb["Notes"]["A5"] = "Procedure growth is the thesis."
    wb.save(path)

    wb2, m2 = build_workbook(fin_for("ISRG", years=3), 7)                   # fewer history columns, more forecast
    kept, dropped = carry_over(path, wb2, m2.reg)
    g2 = m2.grid
    assert wb2["Inputs"][f"B{I['Scenario']}"].value == "Bull"
    assert wb2["Drivers"].cell(row=g2.scen_rows["Gross margin"], column=8).value == 0.70
    assert wb2["Drivers"].cell(row=g2.rows["Revenue growth"][1], column=g2.fcst[1].j).value == 0.25
    assert wb2["Revenue Build"].cell(row=5, column=g2.last_fy.j).value == 6000
    assert wb2["Revenue Build"]["A5"].value == "Instruments & accessories"
    assert wb2["Inputs"][f"B{I['beta']}"].value == 1.3 and wb2["Inputs"][f"B{I['RSUs']}"].value == 4_000_000
    assert wb2["Notes"]["A5"].value == "Procedure growth is the thesis."
    assert len(kept) == 9 and dropped == []
    # untouched defaults are regenerated, not copied: only what the user changed is in the list
    assert not any(k.startswith("Scenario|Revenue growth") for k, _ in kept)


def test_rebuild_of_an_unedited_workbook_carries_nothing(fins, tmp_path):
    from edgar.inputs import carry_over
    path = tmp_path / "KO.xlsx"
    build_workbook(fins["KO"], 5)[0].save(path)
    wb2, m2 = build_workbook(fins["KO"], 5)
    assert carry_over(path, wb2, m2.reg) == ([], [])


def test_price_symbol_adds_the_exchange_prefix(tmp_path, monkeypatch):
    import json
    from edgar import sec
    (tmp_path / "company_tickers_exchange.json").write_text(json.dumps({
        "fields": ["cik", "name", "ticker", "exchange"],
        "data": [[1035267, "INTUITIVE SURGICAL INC", "ISRG", "Nasdaq"], [1067983, "BERKSHIRE", "BRK-B", "NYSE"],
                 [1, "Some OTC Co", "OTCX", "OTC"]]}))
    monkeypatch.setattr(sec, "CACHE_DIR", tmp_path)
    monkeypatch.setattr(sec, "OFFLINE", True)
    assert sec.price_symbol("ISRG", "x@example.com") == "XNAS:ISRG"
    assert sec.price_symbol("BRK.B", "x@example.com") == "XNYS:BRK.B"
    assert sec.price_symbol("OTCX", "x@example.com") == "OTCX"
    assert sec.price_symbol("NOPE", "x@example.com") == "NOPE"


# ---- foreign filers (20-F, IFRS, non-dollar) -------------------------------
def _as_ifrs_filer(facts, currency="EUR"):
    """Turn the ISRG fixture into a look-alike foreign filer: annual 20-F only, IFRS tag names, euros."""
    rename = {"RevenueFromContractWithCustomerExcludingAssessedTax": "Revenue", "CostOfGoodsAndServicesSold": "CostOfSales",
              "OperatingIncomeLoss": "ProfitLossFromOperatingActivities", "NetIncomeLoss": "ProfitLossAttributableToOwnersOfParent",
              "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest": "ProfitLossBeforeTax",
              "IncomeTaxExpenseBenefit": "IncomeTaxExpenseContinuingOperations",
              "EarningsPerShareDiluted": "DilutedEarningsLossPerShare",
              "WeightedAverageNumberOfDilutedSharesOutstanding": "AdjustedWeightedAverageShares",
              "CashAndCashEquivalentsAtCarryingValue": "CashAndCashEquivalents", "AssetsCurrent": "CurrentAssets",
              "LiabilitiesCurrent": "CurrentLiabilities", "StockholdersEquity": "EquityAttributableToOwnersOfParent",
              "LiabilitiesAndStockholdersEquity": "EquityAndLiabilities", "InventoryNet": "Inventories",
              "NetCashProvidedByUsedInOperatingActivities": "CashFlowsFromUsedInOperatingActivities",
              "NetCashProvidedByUsedInInvestingActivities": "CashFlowsFromUsedInInvestingActivities",
              "NetCashProvidedByUsedInFinancingActivities": "CashFlowsFromUsedInFinancingActivities",
              "PaymentsToAcquirePropertyPlantAndEquipment": "PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities"}
    out = {"entityName": "Foreign Look-alike NV", "facts": {"ifrs-full": {}, "dei": {}}}
    for ns, target in (("us-gaap", "ifrs-full"), ("dei", "dei")):
        for tag, body in facts["facts"][ns].items():
            units = {}
            for unit, entries in body["units"].items():
                keep = [dict(e, form="20-F") for e in entries if e.get("form") == "10-K"]
                if keep:
                    units[unit.replace("USD", currency)] = keep
            if units:
                out["facts"][target][rename.get(tag, tag)] = dict(body, units=units)
    return out


def test_foreign_ifrs_filer_builds_in_its_own_currency():
    from edgar.xl_dcf import R
    us = fin_for("ISRG")
    facts = _as_ifrs_filer(load_json(FIXTURES / "ISRG.json.gz"))
    fin = load_financials(facts, "FORN", "Foreign Look-alike NV", 5)
    fin.adr_ratio = 5
    assert fin.currency == "EUR" and fin.annual_form == "20-F"
    assert fin.ltm_end is None                                   # interim reports aren't in the SEC data
    for label in ("Revenue", "Operating Income", "Net Income", "Total Assets", "Cash from Operations", "Capex",
                  "Diluted EPS"):
        assert fy_value(fin, label, "2025-12-31") == fy_value(us, label, "2025-12-31"), label
    wb, m = build_workbook(fin, 5)
    a = wb["Inputs"]
    labels = {a.cell(row=r, column=1).value: r for r in range(4, 70)}
    assert "USD:EUR" in a.cell(row=labels["EUR per US dollar (live)"], column=2).value
    assert a.cell(row=labels["Ordinary shares per US-listed share"], column=2).value == 5
    assert "Share price (EUR per ordinary share)" in labels
    assert wb["Income Statement"]["A4"].value == "(EUR millions)"
    assert wb["Summary"].cell(row=12, column=6).value == "$ per US-listed share"
    from edgar import xl_style
    xl_style.set_currency("USD")                                 # don't leak the euro format into other tests


def test_overrides_put_company_tags_first(fins):
    facts = load_json(FIXTURES / "ISRG.json.gz")
    assert fy_value(fins["ISRG"], "Accrued Liabilities", "2025-12-31") is None
    fin = load_financials(facts, "ISRG", "ISRG", 5,
                          {"Accrued Liabilities": [["EmployeeRelatedLiabilitiesCurrent", "OtherAccruedLiabilitiesCurrent"]]})
    assert fy_value(fin, "Accrued Liabilities", "2025-12-31") == 648.4e6 + 412.3e6
    tags = {label: t for mp in fin.maps for label, _, _, t in mp}
    assert tags["Accrued Liabilities"][0] == ("EmployeeRelatedLiabilitiesCurrent", "OtherAccruedLiabilitiesCurrent")


def test_find_tags_lists_matching_tags_with_latest_value():
    from edgar.transform import find_tags
    hits = find_tags(load_json(FIXTURES / "ISRG.json.gz"), ["accrued", "current"])
    assert ("2025-12-31", "OtherAccruedLiabilitiesCurrent") in [(h[0], h[1]) for h in hits]


def test_risk_free_rate_reads_the_latest_fred_value(tmp_path, monkeypatch):
    from edgar import sec
    (tmp_path / "DGS10.csv").write_text("observation_date,DGS10\n2026-11-11,4.11\n2026-11-12,4.07\n2026-11-13,.\n")
    monkeypatch.setattr(sec, "CACHE_DIR", tmp_path)
    monkeypatch.setattr(sec, "OFFLINE", True)
    assert sec.risk_free_rate("") == (0.0407, "2026-11-12, FRED")


# ---- DCF method choices, gap report, roll-up --------------------------------
def test_dcf_has_method_choices_and_ten_year_default(fins):
    from edgar.layout import INPUT_ROWS as I
    from edgar.xl_dcf import CAPEX_MODES, LEASE_MODES, R, SBC_MODES, TV_MODES
    wb, m = build_workbook(fins["ISRG"])
    assert len(m.grid.fcst) == 10 and m.grid.fcst[-1].header == "FY2035E"
    d, inp = wb["DCF"], wb["Inputs"]
    assert [inp[f"B{I[k]}"].value for k in ("tv_mode", "sbc_mode", "capex_mode", "lease_mode")] == \
        [TV_MODES[0], SBC_MODES[0], CAPEX_MODES[0], LEASE_MODES[0]]
    assert TV_MODES[1] in d[f"B{R['tv']}"].value and "SQRT" in d[f"B{R['imp_terminal']}"].value
    assert "Calc!" in d[f"B{R['imp_cagr']}"].value                       # solver steps live on the hidden sheet
    assert "Diluted" in d[f"A{R['shares']}"].value and "claims" in d[f"A{R['claims']}"].value.lower()


def test_old_layout_inputs_still_carry_into_the_new_one(tmp_path):
    """A workbook saved before the Inputs tab existed keeps its typed settings: they are tracked by name."""
    from openpyxl import Workbook
    from edgar.inputs import SHEET, carry_over
    from edgar.layout import INPUT_ROWS as I
    old = Workbook()
    dcf, a = old.active, old.create_sheet("Assumptions")
    dcf.title = "DCF"
    dcf["B5"], dcf["B21"], a["F1"], a["B57"] = 0.0515, 0.04, "Bear", "XNAS:ISRG"
    reg = old.create_sheet(SHEET)
    reg.append(["key", "sheet", "cell", "default (as generated)"])
    for key, sheet, cell, default in (("DCF|Risk-free rate", "DCF", "B5", "0.0425"), ("DCF|Terminal growth", "DCF", "B21", "0.025"),
                                      ("Control|Scenario", "Assumptions", "F1", '"Base"'), ("Market|Ticker", "Assumptions", "B57", '"ISRG"')):
        reg.append([key, sheet, cell, default])
    old.save(tmp_path / "old.xlsx")
    wb, m = build_workbook(fin_for("ISRG"), 5)
    kept, dropped = carry_over(tmp_path / "old.xlsx", wb, m.reg)
    inp = wb["Inputs"]
    assert (inp[f"B{I['rf']}"].value, inp[f"B{I['g']}"].value, inp[f"B{I['Scenario']}"].value, inp[f"B{I['Ticker']}"].value) == \
        (0.0515, 0.04, "Bear", "XNAS:ISRG")
    assert len(kept) == 4 and dropped == []


def test_gap_report_flags_empty_valuation_lines(fins, capsys):
    from edgar.xl_model import gap_report
    wb, m = build_workbook(fins["ISRG"], 5)
    assert gap_report(m, load_json(FIXTURES / "ISRG.json.gz")) == 0
    del m.data["Capex"]
    assert gap_report(m, load_json(FIXTURES / "ISRG.json.gz")) == 1
    out = capsys.readouterr().out
    assert "Capex:" in out and "PaymentsToAcquireProductiveAssets" in out


def test_rollup_row_and_workbook(tmp_path):
    from openpyxl import load_workbook
    from edgar import rollup as ru

    def summary(perp, exit_):
        cells = {(1, 1): "ACME CORP (ACME) - Summary", (5, 1): "Share price (US listing)", (5, 2): 100.0,
                 (8, 1): "Scenario", (8, 2): "Base", (6, 4): "WACC", (6, 5): 0.09, (7, 4): "Terminal growth",
                 (7, 5): 0.025, (8, 4): "Model checks", (8, 5): "ALL CHECKS OK", (12, 1): "Method",
                 (13, 1): "DCF: perpetuity growth", (13, 3): perp, (14, 1): "DCF: exit multiple", (14, 3): exit_,
                 (22, 1): "Revenue growth per year in the model", (22, 2): 0.08, (22, 3): 0.08,
                 (23, 1): "Revenue growth per year the price needs", (23, 2): 0.12, (23, 3): -2146826281}
        return ru.parse_summary(lambda r, c: cells.get((r, c)))

    row = ru.to_row("ACME", {"Bear": summary(80.0, 100.0), "Base": summary(110.0, 130.0), "Bull": summary(150.0, 170.0)},
                    "Base", "2026-11-16 16:30")
    assert row["midpoint"] == 120.0 and row["upside"] == pytest.approx(0.2)
    assert (row["bear"], row["bull"]) == (90.0, 160.0)
    assert row["growth_needed_exit"] is None                     # an Excel error code is not a number
    ru.append_log(tmp_path, [row])
    ru.append_log(tmp_path, [row])
    history = ru.read_log(tmp_path)
    assert len(history) == 2 and history[0]["midpoint"] == 120.0
    ws = load_workbook(ru.write_book(tmp_path, [row], history))["Portfolio"]
    assert ws["A5"].value == "ACME" and ws["G5"].value == 120.0


# ---- financials template, write-up ------------------------------------------
def test_financial_template_builds_an_equity_model(fins):
    from edgar import xl_fin
    from edgar.layout import INPUT_ROWS as I
    wb, m = build_workbook(fins["KO"], 10, template="financial")
    assert m.kind == "fin" and wb.sheetnames == ["Summary", "Inputs", "Equity Model", "Notes", "Mapping", "Raw", "_inputs"]
    ws, g = wb["Equity Model"], m.grid
    assert g.ytd is None and len(g.fcst) == 10
    assert ws.cell(row=xl_fin.ROW["Net income"], column=g.last_fy.j).value == fy_value(fins["KO"], "Net Income", fins["KO"].cols[-1])
    ni = ws.cell(row=xl_fin.ROW["Net income"], column=g.fcst[0].j).value            # ROE x opening equity
    assert str(xl_fin.ROW["Return on equity"]) in ni and g.last_fy.L in ni
    assert wb["Summary"]["A13"].value == "Equity model: residual income"
    assert wb["Inputs"][f"A{I['multiple']}"].value == "Exit price-to-book multiple"
    assert wb["Inputs"][f"B{I['tv_mode']}"].value is None                           # no DCF-only settings
    # a normal company is not mistaken for a financial one
    assert not fins["KO"].looks_financial() and build_workbook(fins["KO"], 5)[1].kind == "standard"


def test_writeup_renders_the_page():
    from datetime import date
    from edgar import writeup as wu
    s = {"company": "ACME CORP (ACME)", "currency": "USD", "price": 100.0, "perp": 110.0, "exit": 130.0, "peers": 95.0,
         "wacc": 0.09, "terminal_growth": 0.025, "checks": "ALL CHECKS OK", "scenario": "Base",
         "growth_model": (0.08, 0.08), "growth_needed": (0.12, 0.10)}
    by = {"Bear": dict(s, perp=80.0, exit=100.0), "Base": s, "Bull": dict(s, perp=150.0, exit=170.0)}
    summary = [[None] * 7 for _ in range(40)]
    summary[4][3:5], summary[5][1], summary[5][3:5] = ["Valuation date", date(2026, 11, 16)], date(2026, 11, 13), ["WACC", 0.09]
    summary[20][0], summary[21][1:3] = "Key figures", ["FY2025A", "FY2026E"]
    summary[22][:3], summary[23][:3] = ["Revenue", 5e9, 6e9], ["Revenue growth", 0.1, 0.2]
    notes = [[None]] * 3 + [["Thesis"], ["It compounds."], [None], ["Risks and what would change my mind"], ["Competition."]]
    md = wu.render("ACME", by, "Base", (0.25, 0.5, 0.25), {"summary": summary, "inputs": [[None, None]] * 70,
                                                           "drivers": [], "notes": notes}, date(2026, 11, 16))
    assert "| **Midpoint of the two** | **$90.00** | **$120.00** | **$160.00** |" in md
    assert "Probability-weighted value: $122.50" in md and "| Revenue | 5,000 | 6,000 |" in md
    assert "## Thesis" in md and "- It compounds." in md and "## Risks and what would change my mind" in md
