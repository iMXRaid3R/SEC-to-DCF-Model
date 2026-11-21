"""WACC tab: the discount rate the DCF uses, worked out in one place.

    Cost of equity       risk-free rate + beta x equity risk premium (settings: default / your override / used)
    Beta                 measured from prices (benchmark, window, frequency), then Blume-adjusted
    Cost of debt         your pre-tax rate (last year's interest / debt beside it for reference), after tax
    Weights              market value of equity and debt
    WACC                 the blend; the DCF tab links to it

Built with the DCF part. Settings keep the registry keys they had on the DCF tab, so typed values carry over."""
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.datavalidation import DataValidation

from . import xl_dcf, xl_inputs
from . import xl_style as st
from .layout import Ctx, input_ref

SHEET = "WACC"
W = {"eq_title": 4, "eq_hdr": 5, "rf": 6, "erp": 7, "beta": 8, "ke": 9,
     "beta_title": 11, "bench": 12, "years": 13, "freq": 14, "raw": 15, "adj": 16,
     "debt_title": 18, "debt_hdr": 19, "kd": 20, "kd_hist": 21, "tax": 22, "kd_at": 23,
     "cap_title": 25, "price": 26, "shares": 27, "equity": 28, "debt": 29, "w_e": 30, "w_d": 31,
     "wacc_title": 33, "wacc": 34}


def ref(key, col="B"):
    """Absolute reference to a WACC cell from another sheet, e.g. ref("ke")."""
    return f"{SHEET}!${col}${W[key]}"


def write_wacc(ws, m):
    g, fin = m.grid, m.fin
    last, yr1 = Ctx(g, "WACC", g.last_fy), Ctx(g, "WACC", g.fcst[0])
    B = lambda k: f"$B${W[k]}"
    ws["A1"], ws["A1"].font = f"{fin.name} ({fin.ticker}) - WACC", st.TITLE_FONT
    ws["A2"] = ("The discount rate the DCF uses. Settings: Default is what the model suggests; type in the yellow "
                "Your override column to replace it (clear it to go back); Used is what the model uses.")
    ws["A2"].font = st.NOTE_FONT

    def calc(key, label, f, fmt, note="", bold=False):
        ws.cell(row=W[key], column=1, value=label).font = st.font(bold=bold)
        c = ws.cell(row=W[key], column=2, value=f)
        c.number_format, c.alignment = st.FMT[fmt], Alignment(horizontal="right")
        c.font = st.formula_font(f, bold) if isinstance(f, str) else st.font(bold=bold)
        if note:
            st.hover_note(ws.cell(row=W[key], column=6), note).font = st.NOTE_FONT
        return c

    # ---- cost of equity
    st.section(ws, W["eq_title"], 1, 6, "Cost of equity")
    xl_inputs.override_header(ws, W["eq_hdr"])
    rf, rf_day = fin.rf
    xl_inputs.setting(ws, m, W["rf"], "rf", "Risk-free rate", rf if rf else xl_dcf.DEFAULTS["rf"], "pct",
                      f"10-year Treasury yield as of {rf_day}" if rf else
                      "Default: the 10-year Treasury yield could not be downloaded for this build")
    erp, erp_source = fin.erp
    xl_inputs.setting(ws, m, W["erp"], "erp", "Equity risk premium", erp if erp else xl_dcf.DEFAULTS["erp"], "pct",
                      erp_source if erp else "Default: set \"equity_risk_premium\" in watchlist.json to change it "
                                             "for every model")
    xl_inputs.setting(ws, m, W["beta"], "beta", "Beta",
                      f"=IFERROR(ROUND({B('adj')},2),{xl_dcf.DEFAULTS['beta']:g})", "num",
                      "Defaults to the adjusted beta measured below (1.0 without live prices)")
    calc("ke", "Cost of equity", f"={B('rf')}+{B('beta')}*{B('erp')}", "pct",
         "Risk-free rate + beta x equity risk premium (CAPM)", bold=True)

    # ---- beta, measured from prices
    st.section(ws, W["beta_title"], 1, 6, "Beta, measured from prices", size=11)
    dvs = {}

    def pick(key, label, value, options, note, reg):
        ws.cell(row=W[key], column=1, value=label)
        c = ws.cell(row=W[key], column=2, value=value)
        c.alignment = Alignment(horizontal="right")
        m.reg.add(ws, c, "Market", reg)
        if options:
            dv = DataValidation(type="list", formula1='"' + ",".join(str(o) for o in options) + '"', allow_blank=False)
            ws.add_data_validation(dv)
            dv.add(c)
            dvs[key] = dv
        st.hover_note(ws.cell(row=W[key], column=6), note).font = st.NOTE_FONT

    pick("bench", "Benchmark", "SPY", None, "Index ETF the beta is measured against", "Beta benchmark")
    pick("years", "Window (years)", 5, (2, 5), "2 or 5 years of history", "Beta window (years)")
    pick("freq", "Return frequency", "Monthly", ("Weekly", "Monthly"),
         "5 years of monthly returns is the usual convention; 2 years of weekly reacts faster to recent trading",
         "Beta return frequency")
    # TAKE / DROP keep this a plain formula: Excel adds an implicit-intersection "@" to INDEX(SEQUENCE()) in
    # formulas loaded from a file, which breaks it.
    px = lambda tick: (f"_xlfn.STOCKHISTORY({tick},TODAY()-ROUND(365.25*{B('years')},0),TODAY(),"
                       f'IF({B("freq")}="Weekly",1,2),0,1)')
    ret = lambda v: f"_xlfn.DROP(_xlpm.{v},1)/_xlfn.DROP(_xlpm.{v},-1)-1"
    beta = ("=IFERROR(_xlfn.LET("
            f"_xlpm.a,{px(input_ref('Ticker'))},_xlpm.b,{px(B('bench'))},"
            "_xlpm.n,MIN(ROWS(_xlpm.a),ROWS(_xlpm.b)),"
            "_xlpm.x,_xlfn.TAKE(_xlpm.a,-_xlpm.n),_xlpm.y,_xlfn.TAKE(_xlpm.b,-_xlpm.n),"
            f'SLOPE({ret("x")},{ret("y")})),"")')
    calc("raw", "Raw beta (regression)", beta, "num",
         "Slope of the stock's returns on the benchmark's over the window above (live prices, Microsoft 365)")
    calc("adj", "Adjusted beta (Blume)", f'=IF(ISNUMBER({B("raw")}),2/3*{B("raw")}+1/3,"")', "num",
         "2/3 x raw + 1/3 x 1.0: measured betas drift toward the market's 1.0 over time")

    # ---- cost of debt
    st.section(ws, W["debt_title"], 1, 6, "Cost of debt")
    xl_inputs.override_header(ws, W["debt_hdr"])
    xl_inputs.setting(ws, m, W["kd"], "kd", "Pre-tax cost of debt", xl_dcf.DEFAULTS["kd"], "pct",
                      "What the company would pay to borrow today: its bond yields, or the risk-free rate plus a "
                      "credit spread for its rating")
    debt_fy = (f"({last.n('Short-Term Borrowings')}+{last.n('Current Portion of LT Debt')}"
               f"+{last.n('Long-Term Debt')})")
    calc("kd_hist", f"For reference: {g.last_fy.header} interest / year-end debt",
         f'=IFERROR(IF({debt_fy}<=0,"no debt",{last.n("Interest Expense")}/{debt_fy}),"")', "pct",
         "Interest expense over debt at the year end. A rough check on the rate above: it reflects older, "
         "possibly cheaper borrowing")
    calc("tax", f"Tax rate ({g.fcst[0].header})", f"={yr1.v('Tax rate')}", "pct", "From the Drivers tab")
    calc("kd_at", "After-tax cost of debt", f"={B('kd')}*(1-{B('tax')})", "pct", "Interest is tax-deductible",
         bold=True)

    # ---- weights at market value
    st.section(ws, W["cap_title"], 1, 6, "Capital structure at market value")
    calc("price", "Share price", f"=N({last.market('Share price')})", "price", "Inputs tab")
    calc("shares", "Diluted shares", f"=N({last.market('Diluted shares')})", "shares", "Bridge tab")
    calc("equity", "Equity value (market cap)", f"={B('price')}*{B('shares')}", "usd")
    from . import xl_bridge as br
    calc("debt", "Debt", f"={br.ref('debt')}", "usd",
         "From the Bridge tab: latest balance sheet, with leases when the lease choice there says so")
    calc("w_e", "Equity weight", f'=IFERROR({B("equity")}/({B("equity")}+{B("debt")}),1)', "pct")
    calc("w_d", "Debt weight", f'=IFERROR({B("debt")}/({B("equity")}+{B("debt")}),0)', "pct")

    # ---- the result
    st.section(ws, W["wacc_title"], 1, 6, "WACC")
    c = calc("wacc", "WACC", f"=IF(OR({B('price')}=0,{B('equity')}=0),{B('ke')},"
                             f"{B('w_e')}*{B('ke')}+{B('w_d')}*{B('kd_at')})", "pct",
             "Equity weight x cost of equity + debt weight x after-tax cost of debt. Uses the cost of equity alone "
             "when there is no share price", bold=True)
    c.font = Font(bold=True, size=13, color=st.THEME["accent"])
    c.fill = st.PatternFill("solid", fgColor=st.THEME["pale"])
    st.row_rules(ws, [W[k] for k in ("ke", "raw", "adj", "kd_hist", "tax", "kd_at", "price", "shares", "equity",
                                      "debt", "w_e", "w_d")], 6)
    for col, w in zip("ABCDEF", (44, 14, 14, 14, 30, 90)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "B4"
