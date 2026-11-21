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
                             "Cash Flow", "WACC", "DCF", "Multiples", "Checks & Metrics", "Notes", "Mapping", "Raw", "Calc",
                             "_inputs"]
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
    assert [comps.cell(row=4, column=j).value for j in (2, 3, 4, 5, 6)] == ["ISRG", "KO", "COST", "Peer median", "Peer mean"]
    labels = {}                                                                      # tall: one row per metric
    for r in range(5, 60):
        labels.setdefault(comps.cell(row=r, column=1).value, r)
    assert {"Revenue", "EBITDA margin", "Share price", "EV / EBITDA", "P / E"} <= set(labels)
    med = comps.cell(row=labels["EV / EBITDA"], column=5).value
    assert med == f'=IFERROR(MEDIAN(C{labels["EV / EBITDA"]}:D{labels["EV / EBITDA"]}),"")'   # peers only, not ISRG
    assert comps.cell(row=labels["Price override"], column=3).fill.fgColor.rgb.endswith("FFF2CC")
    assert "Inputs!" in comps.cell(row=labels["Share price"], column=2).value     # the subject's own price
    from edgar.layout import INPUT_ROWS
    assert "Comps!" in wb["DCF"][f"C{R['multiple']}"].value             # default on the DCF tab: the peer median
    assert f"DCF!$B${R['multiple']}" in wb["Inputs"][f"B{INPUT_ROWS['multiple']}"].value   # Inputs links to it
    assert wb["Inputs"][f"A{INPUT_ROWS['multiple']}"].hyperlink.location == f"'DCF'!D{R['multiple']}"


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
    drv.cell(row=g.scen_ovr_rows["Gross margin"], column=8).value = 0.70    # Bull, year 1 override
    yr = g.year_ovr_rows["Revenue growth"]
    drv.cell(row=yr, column=g.fcst[1].j).value = 0.25                       # FY2027E override
    drv.cell(row=yr, column=g.fcst[0].j).value = 0.11                       # FY2026E override
    rb.cell(row=5, column=g.last_fy.j).value = 6000                         # FY2025A segment revenue
    rb["A5"] = "Instruments & accessories"
    from edgar.xl_wacc import W
    wb["WACC"][f"D{W['beta']}"] = 1.3
    inp[f"B{I['RSUs']}"] = 4_000_000
    wb["Notes"]["A5"] = "Procedure growth is the thesis."
    wb.save(path)

    wb2, m2 = build_workbook(fin_for("ISRG", years=3), 7)                   # fewer history columns, more forecast
    kept, dropped, _ = carry_over(path, wb2, m2.reg)
    g2 = m2.grid
    assert wb2["Inputs"][f"B{I['Scenario']}"].value == "Bull"
    assert wb2["Drivers"].cell(row=g2.scen_ovr_rows["Gross margin"], column=8).value == 0.70
    assert wb2["Drivers"].cell(row=g2.year_ovr_rows["Revenue growth"], column=g2.fcst[1].j).value == 0.25
    live = wb2["Drivers"].cell(row=g2.rows["Revenue growth"][1], column=g2.fcst[1].j).value
    assert live.startswith(f'=IF({g2.fcst[1].L}${g2.year_ovr_rows["Revenue growth"]}=""')   # the formula is kept
    assert wb2["Revenue Build"].cell(row=5, column=g2.last_fy.j).value == 6000
    assert wb2["Revenue Build"]["A5"].value == "Instruments & accessories"
    assert wb2["WACC"][f"D{W['beta']}"].value == 1.3 and wb2["Inputs"][f"B{I['RSUs']}"].value == 4_000_000
    assert f"$B${W['adj']}" in wb2["WACC"][f"C{W['beta']}"].value                 # default: the adjusted beta
    assert wb2["Notes"]["A5"].value == "Procedure growth is the thesis."
    assert len(kept) == 9 and dropped == []
    # untouched defaults are regenerated, not copied: only what the user changed is in the list
    assert not any(k.startswith("Scenario|Revenue growth") for k, _ in kept)


def test_rebuild_of_an_unedited_workbook_carries_nothing(fins, tmp_path):
    from edgar.inputs import carry_over
    path = tmp_path / "KO.xlsx"
    build_workbook(fins["KO"], 5)[0].save(path)
    wb2, m2 = build_workbook(fins["KO"], 5)
    assert carry_over(path, wb2, m2.reg) == ([], [], [])


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
    s = wb["Summary"]
    head = next(r for r in range(1, 40) if s.cell(row=r, column=1).value == "Method")
    assert s.cell(row=head, column=6).value == "$ per US-listed share"
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
    assert [d[f"C{R[k]}"].value for k in ("tv_mode", "sbc_mode", "capex_mode", "lease_mode")] == \
        [TV_MODES[0], SBC_MODES[0], CAPEX_MODES[0], LEASE_MODES[0]]
    assert [inp[f"B{I[k]}"].value for k in ("tv_mode", "g")] == [f'=IF(DCF!$B${R[k]}="","",DCF!$B${R[k]})'
                                                                for k in ("tv_mode", "g")]
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
    kept, dropped, _ = carry_over(tmp_path / "old.xlsx", wb, m.reg)
    from edgar.xl_dcf import R as D
    from edgar.xl_wacc import W
    inp, dcf = wb["Inputs"], wb["DCF"]
    assert (wb["WACC"][f"D{W['rf']}"].value, dcf[f"D{D['g']}"].value, inp[f"B{I['Scenario']}"].value, inp[f"B{I['Ticker']}"].value) == \
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
    heads = {ws.cell(row=4, column=j).value: j for j in range(1, ws.max_column + 1)}
    assert ws["A5"].value == "ACME" and ws.cell(row=5, column=heads["DCF midpoint"]).value == 120.0


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
    assert "Probability-weighted DCF midpoint: $122.50" in md and "| Revenue | 5,000 | 6,000 |" in md
    assert "## Thesis" in md and "- It compounds." in md and "## Risks and what would change my mind" in md


# ---- what changed since the last build ------------------------------------------
def _filed_by(facts, cutoff):
    """The company facts as they stood on `cutoff`: everything filed later is dropped."""
    import copy
    f = copy.deepcopy(facts)
    for ns in f["facts"].values():
        for tag in ns.values():
            for unit in tag["units"]:
                tag["units"][unit] = [x for x in tag["units"][unit] if x["filed"] <= cutoff]
    return f


def test_change_report_after_a_new_quarter(fins, tmp_path, capsys):
    from datetime import date
    from edgar import changes as ch
    new = fins["ISRG"]
    filed = max(r for r in new.sources["Filed"])
    facts = load_json(FIXTURES / "ISRG.json.gz")
    old = load_financials(_filed_by(facts, str(filed - pd.Timedelta(days=1))[:10]), "ISRG", "Intuitive Surgical")
    assert old.ltm_end < new.ltm_end
    path = tmp_path / "ISRG_financials.xlsx"
    build_workbook(old, 5)[0].save(path)

    a, b = ch.Snapshot(ch.read_raw(path)), ch.Snapshot(new.sources)
    assert (a.q_end, b.q_end) == (old.ltm_end, new.ltm_end) and a.ltm["Revenue"] == old.ltm["Revenue"]
    sections = ch.compare(a, b)
    title, header, rows = sections[0]
    assert title.startswith("Quarter ended") and rows[0][0] == "Revenue"
    quarter, year_ago = (float(x.replace(",", "")) for x in rows[0][1:3])     # $ millions
    assert quarter == pytest.approx((new.ytd["Revenue"] - old.ytd["Revenue"]) / 1e6, abs=1)
    assert quarter - year_ago == pytest.approx((new.ltm["Revenue"] - old.ltm["Revenue"]) / 1e6, abs=1)

    assert ch.report(new, path, tmp_path, date(2026, 11, 16)) == "new"
    text = (tmp_path / "changes" / "ISRG_2026-11-16.md").read_text(encoding="utf-8")
    assert "New quarterly report" in text and "| Revenue |" in text and "Balance sheet" in text
    assert "What changed since the last build" in capsys.readouterr().out

    build_workbook(new, 5)[0].save(path)                                      # same data: nothing to report
    assert ch.report(new, path, tmp_path) is None
    assert "No change in the reported figures" in capsys.readouterr().out


def test_change_report_lists_restated_years(fins):
    from edgar import changes as ch
    src = fins["ISRG"].sources
    before = src.copy()
    row = before[(before["Line Item"] == "Revenue") & (before["Form"] == "10-K")].index[0]
    before.loc[row, "Value"] = before.loc[row, "Value"] * 0.9
    sections = ch.compare(ch.Snapshot(before), ch.Snapshot(src))
    assert [s[0] for s in sections] == ["Earlier years restated"]
    assert sections[0][2][0][0].startswith("Revenue, FY to") and sections[0][2][0][3] == "+11.1%"


def test_rollup_reports_values_that_moved():
    from edgar import rollup as ru
    history = [{"ticker": "ACME", "run": "2026-08-01 09:00", "midpoint": 90.0, "price": 95.0},
               {"ticker": "ACME", "run": "2026-11-01 09:00", "midpoint": 100.0, "price": 100.0},
               {"ticker": "SAME", "run": "2026-11-01 09:00", "midpoint": 50.0, "price": 40.0}]
    rows = [{"ticker": "ACME", "midpoint": 120.0, "price": 110.0}, {"ticker": "SAME", "midpoint": 50.0, "price": 41.0},
            {"ticker": "NEW", "midpoint": 10.0, "price": 9.0}]
    lines = ru.value_changes(rows, history)
    assert len(lines) == 1 and "ACME" in lines[0] and "was $100.00 on 2026-11-01 09:00 (+20.0%)" in lines[0]
    assert rows[0]["prev_midpoint"] == 100.0 and rows[1]["midpoint_change"] == 0 and "prev_run" not in rows[2]


# ---- build options, multiples, menu ------------------------------------------
def test_parts_choose_the_tabs_and_are_remembered(fins, tmp_path):
    from edgar import parts as P
    assert P.parse("statements") == set() and P.parse("all") == set(P.PARTS)
    assert P.parse("dcf, peg,revenue") == {"dcf", "multiples", "revenue_build"}
    with pytest.raises(ValueError):
        P.parse("dfc")
    assert P.resolve(None, ["dcf"], remembered={"dcf", "comps"}) == {"comps"}
    assert P.resolve({"dcf"}, [], remembered=set()) == {"dcf"} and P.resolve(None, [], default=set()) == set()

    wb, m = build_workbook(fins["ISRG"], 5, parts=set())
    assert wb.sheetnames == ["Summary", "Inputs", "Drivers", "Income Statement", "Balance Sheet", "Cash Flow",
                             "Checks & Metrics", "Notes", "Mapping", "Raw", "_inputs"]
    rev = wb["Income Statement"].cell(row=m.grid.rows["Revenue"][1], column=m.grid.fcst[1].j).value
    assert "Revenue Build" not in rev                                    # no link to a tab that isn't there
    for ws in wb.worksheets:                                             # nothing points at a missing tab
        for row in ws.iter_rows():
            for c in row:
                assert not (isinstance(c.value, str) and ("DCF!" in c.value or "Multiples!" in c.value)), c.coordinate
    path = tmp_path / "ISRG_financials.xlsx"
    wb.save(path)
    assert P.previous(path) == set()
    wb, m = build_workbook(fins["ISRG"], 5, parts={"multiples"})
    assert "Multiples" in wb.sheetnames and "DCF" not in wb.sheetnames
    s = wb["Summary"]
    col_a = [s.cell(row=r, column=1).value for r in range(1, 40)]
    assert "Forward P/E" in col_a and "DCF: exit multiple" not in col_a
    wacc = next(r for r in range(1, 20) if s.cell(row=r, column=4).value == "WACC")
    assert s.cell(row=wacc, column=6).value == "n/a (no DCF)"


def test_inputs_on_a_switched_off_tab_wait_for_it(fins, tmp_path):
    from edgar.inputs import carry_over
    from edgar.layout import INPUT_ROWS as I
    from edgar import xl_multiples as xm
    wb, m = build_workbook(fins["ISRG"], 5)
    O = xm.COL["override"]
    wb["Multiples"][f"{O}{xm.R['pe']}"] = 31.5                           # a typed target P/E
    from edgar.xl_wacc import W
    wb["WACC"][f"D{W['kd']}"] = 0.061                                    # a typed cost of debt
    pe_weight = next(r for r in range(1, 40) if wb["Summary"].cell(row=r, column=1).value == "Forward P/E")
    wb["Summary"].cell(row=pe_weight, column=7).value = 0.4                 # its weight in the price target
    wb.save(tmp_path / "a.xlsx")
    wb2, m2 = build_workbook(fins["ISRG"], 5, parts=set())               # statements only
    kept, dropped, parked = carry_over(tmp_path / "a.xlsx", wb2, m2.reg)
    assert dict(parked) == {"Multiples|Forward P/E|Target": 31.5, "DCF|Pre-tax cost of debt": 0.061,
                            "Summary|Weight|Forward P/E": 0.4} and not dropped
    wb2.save(tmp_path / "b.xlsx")
    wb3, m3 = build_workbook(fins["ISRG"], 5)                            # everything again
    kept, dropped, parked = carry_over(tmp_path / "b.xlsx", wb3, m3.reg)
    assert wb3["Multiples"][f"{O}{xm.R['pe']}"].value == 31.5 and wb3["WACC"][f"D{W['kd']}"].value == 0.061
    assert ("Summary|Weight|Forward P/E", 0.4) in kept
    assert not parked


def test_multiples_tab_formulas(fins):
    from edgar import xl_multiples as xm
    from edgar.comps import _metrics
    wb, m = build_workbook(fins["ISRG"], 10)
    ws, r, C = wb["Multiples"], xm.R, xm.COL
    assert ws[f"B{r['eps']}"].value == f"=C{r['eps']}*(1-$B${r['elapsed']})+D{r['eps']}*$B${r['elapsed']}"
    assert "Income Statement" in ws[f"E{r['eps']}"].value                 # reported (LTM) figures beside them
    peg = ws[f"{C['value']}{r['peg']}"].value                            # EPS x PEG x growth in %
    assert f"B{r['eps']}*({C['target']}{r['peg']})*B{r['peg']}*100" in peg and f'{C["use"]}{r["peg"]}="On"' in peg
    ev = ws[f"{C['value']}{r['ev_ebitda']}"].value
    assert f"(B{r['ev_ebitda']}*({C['target']}{r['ev_ebitda']})-$B${r['claims']})/$B${r['shares']}" in ev
    t = r["pe"]                                                          # target used = your target, else default
    D, O, T, RG = (f"{C[k]}{t}" for k in ("default", "override", "target", "range"))
    assert ws[T].value == f'=IF({O}="",IF({D}="","",{D}),{O})' and ws[O].value is None
    assert f"{T}*(1-N({RG}))" in ws[f"{C['v_low']}{t}"].value and ws[RG].value == xm.SPREAD
    # today's multiple on the forecast (column B of the figures) and on reported results (column E)
    assert f"/B{r['eps']}" in ws[f"{C['now']}{t}"].value and f"/E{r['eps']}" in ws[f"{C['now_ltm']}{t}"].value
    assert "forecast" in ws[f"{C['now']}{xm.MET_HDR}"].value and "reported" in ws[f"{C['now_ltm']}{xm.MET_HDR}"].value
    assert ws[f"B{r['claims']}"].value == "=DCF!$B$" + str(__import__("edgar.xl_dcf").xl_dcf.R["claims"])
    assert ws[f"{C['default']}{t}"].value == xm.DEFAULTS["pe"] and ws[f"{C['use']}{r['ev_sales']}"].value == "Off"
    assert ws[f"{C['peer']}{t}"].value is None                          # no peers, no peer column value
    wb, _ = build_workbook(fins["ISRG"], 10, parts={"multiples"})        # without the DCF: its own bridge
    own = wb["Multiples"][f"B{r['claims']}"].value
    assert "DCF" not in own and "Long-Term Debt" not in own and "Balance Sheet" in own
    wb, m = build_workbook(fins["ISRG"], 10, [_metrics(fins[x]) for x in ("KO", "COST")])
    ws = wb["Multiples"]
    assert "Comps!" in ws[f"{C['peer']}{t}"].value and "Comps!" in ws[f"{C['default']}{t}"].value
    assert ws[f"{C['peer']}{r['p_fcf']}"].value is None                  # Comps has no P/FCF


def test_summary_parse_handles_multiples_and_statements_only():
    from edgar import rollup as ru
    cells = {(1, 1): "ACME CORP (ACME) - Summary", (5, 1): "Share price (US listing)", (5, 2): 100.0,
             (8, 1): "Scenario", (8, 2): "Base", (6, 4): "WACC", (6, 5): "n/a (no DCF)", (7, 4): "Terminal growth",
             (7, 5): 0.025, (8, 4): "Model checks", (8, 5): "ALL CHECKS OK", (12, 1): "Method",
             (13, 1): "Forward P/E", (13, 3): 90.0, (14, 1): "PEG", (14, 3): 110.0, (16, 1): "Key figures"}
    s = ru.parse_summary(lambda r, c: cells.get((r, c)))
    assert (s["fwd_pe"], s["peg"], s["perp"], s["wacc"]) == (90.0, 110.0, None, None)
    row = ru.to_row("ACME", {"Base": s}, "Base", "2026-11-20 10:00")
    assert row["mult_avg"] == 100.0 and row["midpoint"] is None and "without the DCF" in row["note"]
    del cells[(12, 1)], cells[(13, 1)], cells[(14, 1)]                   # statements only: no method table
    s = ru.parse_summary(lambda r, c: cells.get((r, c)))
    assert s["fwd_pe"] is None and s["price"] == 100.0


def test_menu_builds_with_the_chosen_parts(monkeypatch, capsys):
    from edgar import menu, sec
    monkeypatch.setattr(sec, "DEFAULT_EMAIL", "someone@example.com")
    answers = iter(["1", "acme", "1 4", "", "n", "q"])                   # build ACME without DCF and Revenue Build
    monkeypatch.setattr("builtins.input", lambda prompt="": next(answers))
    monkeypatch.setattr(menu, "_watchlist", lambda: {"peers": {"ACME": ["BETA"]}})
    calls = []

    def fake_main(args):
        calls.append(args)
    fake_main.open_file = lambda path: None
    menu.menu(fake_main)
    assert calls == [["ACME", "--parts", "comps,multiples"]]
    assert "> ./run ACME --parts comps,multiples" in capsys.readouterr().out


# ---- input highlighting, price target ------------------------------------------
def test_every_input_cell_is_highlighted(fins):
    from edgar.xl_style import INPUT_FILL
    wb, m = build_workbook(fins["ISRG"], 5)
    looked = 0
    for key, (sheet, addr, _) in m.reg.items.items():
        if key.split("|")[0] in ("Live", "Notes"):
            continue
        c = wb[sheet][addr]
        assert c.fill.fgColor.rgb.endswith(INPUT_FILL.fgColor.rgb[-6:]) and c.font.color.rgb.endswith("0000FF"), key
        looked += 1
    assert looked > 100
    # the SEC history on the statements is not an input
    ws, g = wb["Income Statement"], m.grid
    assert ws.cell(row=g.rows["Revenue"][1], column=g.last_fy.j).fill.fgColor.rgb in (None, "00000000")


def test_summary_price_target_weights_the_methods(fins):
    from edgar import xl_summary as xs
    wb, m = build_workbook(fins["ISRG"], 5)
    ws = wb["Summary"]
    rows = {ws.cell(row=r, column=1).value: r for r in range(1, 60)}
    first, last = rows["DCF: perpetuity growth"], rows["Peer multiples"] if "Peer multiples" in rows else rows["Price / free cash flow"]
    assert ws.cell(row=first, column=xs.WEIGHT_COL).value == xs.WEIGHTS["DCF: perpetuity growth"]
    pt = rows[xs.TARGET]
    f = ws.cell(row=pt, column=3).value
    assert f"SUMPRODUCT(G{first}:G{last},C{first}:C{last})" in f and f"ISNUMBER(C{first}:C{last})" in f
    rr = rows[xs.ROLL]
    roll = ws.cell(row=rr, column=2).value                                # default: cost of equity less dividends
    assert "WACC!$B$" in roll and ws.cell(row=rr, column=3).value == f'=IF(D{rr}="",IF(B{rr}="","",B{rr}),D{rr})'
    assert ws.cell(row=rows[xs.TARGET_12M], column=3).value.startswith(f'=IFERROR(IF(OR(C{pt}=""')
    assert sum(v for v in xs.WEIGHTS.values()) == pytest.approx(1.0)


def test_rollup_reads_the_price_target():
    from edgar import rollup as ru
    def summary(target):
        cells = {(1, 1): "ACME CORP (ACME) - Summary", (5, 1): "Share price (US listing)", (5, 2): 100.0,
                 (8, 1): "Scenario", (8, 2): "Base", (7, 4): "Terminal growth", (7, 5): 0.025, (8, 4): "Model checks",
                 (8, 5): "ALL CHECKS OK", (12, 1): "Method", (13, 1): "Forward P/E", (13, 3): target,
                 (15, 1): "Price target (weighted)", (15, 3): target, (16, 1): "12-month price target", (16, 3): target * 1.1}
        return ru.parse_summary(lambda r, c: cells.get((r, c)))
    row = ru.to_row("ACME", {"Bear": summary(80.0), "Base": summary(120.0), "Bull": summary(160.0)}, "Base", "x",
                    (0.25, 0.5, 0.25))
    assert (row["target"], row["target_12m"]) == (120.0, pytest.approx(132.0))
    assert row["upside_target"] == pytest.approx(0.2) and row["target_weighted"] == pytest.approx(120.0)


# ---- inputs next to what they change --------------------------------------
def test_settings_are_typed_where_they_act_and_listed_on_inputs(fins):
    from edgar.layout import INPUT_ROWS as I
    from edgar.xl_dcf import R
    wb, m = build_workbook(fins["ISRG"], 5)
    dcf, inp = wb["DCF"], wb["Inputs"]
    from edgar.xl_wacc import W
    for key, sheet, rows in (("rf", "WACC", W), ("beta", "WACC", W), ("g", "DCF", R), ("multiple", "DCF", R),
                             ("tv_mode", "DCF", R), ("other_claims", "DCF", R)):
        ws, r = wb[sheet], rows[key]
        assert m.reg.items[f"DCF|{m.settings[key]['label']}"][:2] == (sheet, f"D{r}")       # the override cell
        assert ws[f"D{r}"].fill.fgColor.rgb.endswith("FFF2CC") and ws[f"D{r}"].value is None
        assert ws[f"B{r}"].value == f'=IF(D{r}="",IF(C{r}="","",C{r}),D{r})'
        assert inp[f"A{I[key]}"].hyperlink.location == f"'{sheet}'!D{r}"                    # Inputs links to it
        assert f"{sheet}!$B${r}" in inp[f"B{I[key]}"].value                                 # ... shows the value used
        assert f"Inputs!$C${I[key]}" in ws[f"E{r}"].value                                   # its range check
    assert dcf[f"B{R['wacc']}"].value == f"=WACC!$B${W['wacc']}"                            # the DCF links to it
    # the index of inputs on other tabs, and click-throughs from the Summary's methods
    links = {inp.cell(row=r, column=1).value: inp.cell(row=r, column=1).hyperlink.location
             for r in range(max(I.values()) + 3, max(I.values()) + 12) if inp.cell(row=r, column=1).hyperlink}
    assert set(links) >= {"Bear / Base / Bull drivers", "DCF settings", "Target multiples", "Price target weights"}
    s = wb["Summary"]
    exit_row = next(r for r in range(1, 40) if s.cell(row=r, column=1).value == "DCF: exit multiple")
    assert s.cell(row=exit_row, column=1).hyperlink.location == f"'DCF'!D{R['multiple']}"   # its override
    # a results line on every tab with inputs
    for name in ("Inputs", "Drivers", "DCF", "Multiples", "Revenue Build"):
        row1 = [c.value for c in wb[name][1] if isinstance(c.value, str)]
        assert any("Price target" in v and "Scenario" in v for v in row1), name


def test_financials_template_keeps_its_settings_on_inputs(fins):
    from edgar.layout import INPUT_ROWS as I
    wb, m = build_workbook(fins["KO"], 10, template="financial")
    c = wb["Inputs"][f"B{I['g']}"]
    assert isinstance(c.value, (int, float)) and c.fill.fgColor.rgb.endswith("FFF2CC")


def test_no_input_cell_holds_a_formula(fins):
    """Every yellow cell is a plain value or blank: changing it never deletes a formula. Defaults that are
    formulas sit beside their override cell."""
    from edgar.comps import _metrics
    wb, m = build_workbook(fins["ISRG"], 10, [_metrics(fins[t]) for t in ("KO", "COST")])
    formulas = [k for k, (sheet, addr, _) in m.reg.items.items()
                if isinstance(wb[sheet][addr].value, str) and wb[sheet][addr].value.startswith("=")]
    assert formulas == []


def test_formula_defaults_match_after_excel_saves_them():
    from edgar.inputs import _same
    assert _same("=IFERROR(ROUND(Comps!$AB$11,1),20.0)", "=IFERROR(ROUND('Comps'!$AB$11,1),20)")
    assert not _same("=X*1.05", "=X*1.5") and not _same("=A10.0", "=A10")


def test_summary_shows_fcf_yield_and_leaves_p_fcf_out_of_the_target(fins):
    from edgar import xl_summary as xs
    wb, m = build_workbook(fins["ISRG"], 10)
    s = wb["Summary"]
    labels = {s.cell(row=r, column=xs.YIELD_COL).value: r for r in range(1, 20)}
    assert {"FCF yield, your forecast (NTM)", "FCF yield after stock-based comp (reported)", "Cost of equity"} <= set(labels)
    f = s.cell(row=labels["FCF yield after stock-based comp (reported)"], column=xs.YIELD_COL + 1).value
    assert "Stock-Based Comp" not in f and "Market cap" not in f and "/Inputs!" in f        # (FCF - SBC) / market cap
    assert xs.WEIGHTS["Price / free cash flow"] == 0
    wb, _ = build_workbook(fins["ISRG"], 10, parts=set())                  # statements only: yields, no rates
    labels = [wb["Summary"].cell(row=r, column=xs.YIELD_COL).value for r in range(1, 20)]
    assert "FCF yield, reported (LTM)" in labels and "Cost of equity" not in labels


def test_summary_dashboard_layout(fins):
    """Headline cards, gridlines off, charts, and the labels the roll-up and write-up look for."""
    from edgar import rollup as ru
    from edgar import xl_summary as xs
    wb, m = build_workbook(fins["ISRG"], 10)
    s = wb["Summary"]
    assert not s.sheet_view.showGridLines and len(s._charts) == 2
    cards = [s.cell(row=xs.CARD_ROW, column=c).value for c in range(1, 13)]
    assert {"SHARE PRICE", "PRICE TARGET", "UPSIDE TO TARGET", "12-MONTH TARGET", "SCENARIO"} <= set(cards)
    col_a = {s.cell(row=r, column=1).value for r in range(1, 90)}
    assert {"Share price (US listing)", "Scenario", "Method", xs.TARGET, xs.TARGET_12M, "Key figures",
            "Revenue growth per year the price needs", "52-week trading range"} <= col_a
    # the parser reads it (values come from Excel; here the cells hold formulas, so just the structure)
    get = lambda r, c: s.cell(row=r, column=c).value
    p = ru.parse_summary(get)
    assert p["company"].endswith("(ISRG)") and isinstance(p["checks"], str)
    wb, _ = build_workbook(fins["ISRG"], 10, parts=set())                  # statements only still has a layout
    assert len(wb["Summary"]._charts) == 1


# ---- house style and brand colours ------------------------------------------
def test_brand_colour_is_made_readable_and_applied(fins):
    from edgar import xl_style as st
    for raw in ("FF9900", "FFFFFF", "76B900", "000000", "0064E0", "not a colour"):
        acc = st.readable_accent(raw)
        assert st.luminance(acc) <= 0.18 and st.luminance(acc) >= 0.012      # white text reads on it; not flat black
    assert st.readable_accent("0064E0") == "0064E0"                          # a dark enough colour is kept as is
    fin = fins["ISRG"]
    fin.brand = "D6002A"
    try:
        wb, m = build_workbook(fin, 5)
    finally:
        fin.brand = None
    for name in ("Summary", "Inputs", "Drivers", "Income Statement", "DCF", "Checks & Metrics"):
        ws = wb[name]
        assert ws["A1"].fill.fgColor.rgb.endswith(st.THEME["accent"]) and not ws.sheet_view.showGridLines, name
    # meaning colours don't change with the brand: inputs stay yellow
    from edgar.xl_dcf import R
    assert wb["DCF"][f"D{R['g']}"].fill.fgColor.rgb.endswith("FFF2CC")
    wb, _ = build_workbook(fins["ISRG"], 5)                                  # no brand: navy
    assert wb["Summary"]["A1"].fill.fgColor.rgb.endswith(st.DEFAULT_ACCENT)


def test_statements_have_margin_lines_and_checks_show_status(fins):
    wb, m = build_workbook(fins["ISRG"], 5)
    g, ws = m.grid, wb["Income Statement"]
    r = g.rows["Gross margin (%)"][1]
    assert r == g.rows["Gross Profit"][1] + 1 and "Gross Profit" not in str(ws.cell(row=r, column=1).font.color)
    assert ws.cell(row=r, column=g.fcst[0].j).value.startswith("=IF(")
    chk = wb["Checks & Metrics"]
    status = chk.cell(row=g.rows["Check: Assets - Liabilities & Equity"][1], column=g.cols[-1].j + 1).value
    assert "✓" in status and "✗" in status
