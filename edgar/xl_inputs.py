"""Inputs tab: every setting in one place - scenario, price and currency, dilution, discount rate, terminal
value, method choices and equity-bridge adjustments. Yellow cells are yours to change; the Check column
flags values that are out of the usual range. The other tabs read from here."""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.datavalidation import DataValidation

from . import xl_dcf
from . import xl_style as st
from .layout import BASIS_OPTIONS, INPUT_HEADERS, INPUT_ROWS, SCENARIOS, SHEET_NAMES

R = INPUT_ROWS
WARN = Font(color="C00000", bold=True)


def write_inputs(ws, m):
    g, fin = m.grid, m.fin
    cur, foreign = fin.currency, fin.currency != "USD"
    ws["A1"], ws["A1"].font = f"{fin.name} ({fin.ticker}) - Inputs", st.TITLE_FONT
    ws["A2"] = ("Every setting in one place. " + st.INPUT_LEGEND + " Settings that act on one tab are typed "
                "there and listed here as green links: click a blue underlined name to go to it. The list of inputs "
                "on other tabs is at the bottom.")
    ws["A2"].font = st.NOTE_FONT
    for j, h in enumerate(("Setting", "Value", "Check", "What it is"), start=1):
        st.header_cell(ws.cell(row=3, column=j), h)
    for title, r in INPUT_HEADERS.items():
        c = ws.cell(row=r, column=1, value=title)
        c.font, c.fill = st.HDR_FONT, st.FILL["fcst"]
        for j in (2, 3, 4):
            ws.cell(row=r, column=j).fill = st.FILL["fcst"]
    B = lambda key: f"B{R[key]}"

    def put(key, label, value, fmt=None, kind="formula", note="", reg=None, options=None, check=None, bold=False):
        """kind: "input" or "optional" (normally blank): registered, so yellow; "formula" (black)."""
        r = R[key]
        ws.cell(row=r, column=1, value=label).font = st.font(bold=bold)
        c = ws.cell(row=r, column=2, value=value)
        c.alignment = Alignment(horizontal="right")
        if fmt:
            c.number_format = fmt
        if kind == "formula":
            c.font = st.formula_font(value, bold) if isinstance(value, str) else st.font(bold=bold)
        else:
            c.font = st.font(st.BLUE, bold)
            if kind == "optional":
                c.fill = st.INPUT_FILL
            m.reg.add(ws, c, *(reg or ("Inputs", label)))
        if options:
            dv = DataValidation(type="list", formula1='"' + ",".join(str(o) for o in options) + '"', allow_blank=False)
            ws.add_data_validation(dv)
            dv.add(c)
            c.fill = st.INPUT_FILL
        if check:
            ws.cell(row=r, column=3, value=check).font = WARN
        if note:
            st.hover_note(ws.cell(row=r, column=4), note).font = st.NOTE_FONT
        return c

    # ---- Scenario
    put("Scenario", "Scenario", "Base", kind="input", options=SCENARIOS, reg=("Control", "Scenario"), bold=True,
        note="Which case drives the whole model. Bear / Base / Bull inputs are on the Scenarios tab")
    put("Basis", "Base case starts from", BASIS_OPTIONS[0], kind="input", options=BASIS_OPTIONS,
        reg=("Control", "Basis"), note="History the Base year-1 drivers are taken from: 3-year average, last "
                                       "fiscal year, or last twelve months")
    put("Use YTD", "Year 1 uses reported year-to-date", "On" if g.ytd is not None else "Off", kind="input",
        options=("On", "Off"),
        note=(f"On: {g.fcst[0].header} = the {fin.ytd_months} months already reported + a forecast for the rest of "
              "the year. Off: the whole year is forecast from last year") if g.ytd is not None else
        "No quarterly report newer than the last annual report, so the whole of year 1 is forecast")
    for key, name, w in (("w_bear", "Bear", 0.25), ("w_base", "Base", 0.5), ("w_bull", "Bull", 0.25)):
        if m.kind == "fin":                         # the financials template has no Scenarios tab
            put(key, f"Probability: {name}", w, st.FMT["pct"], "input",
                note="Used by the portfolio roll-up for a probability-weighted value (the workbook itself shows "
                     "one scenario at a time)" if key == "w_bear" else "")
        else:                                       # set on the Scenarios tab, next to the cases
            r = m.grid.prob_rows[key]
            src = f"{SHEET_NAMES['SCN']}!$B${r}"
            c = put(key, f"Probability: {name}", f"={src}", st.FMT["pct"],
                    note="Set on the Scenarios tab (click the name)" if key == "w_bear" else "")
            st.link_to(ws.cell(row=R[key], column=1), SHEET_NAMES["SCN"], f"B{r}")
    put("w_total", "Probabilities add up to", f"=SUM({B('w_bear')}:{B('w_bull')})", st.FMT["pct"],
        check=f'=IF(ABS({B("w_total")}-1)>0.0001,"Should be 100%","")')

    # ---- Share price and currency
    put("Ticker", "Ticker", fin.symbol or fin.ticker, kind="input", reg=("Market", "Ticker"),
        note="Used for the live price. Without the exchange prefix Excel can pick a foreign listing: use e.g. "
             f"XNAS:{fin.ticker} (Nasdaq) or XNYS:{fin.ticker} (NYSE)")
    put("Live close", "Live close (STOCKHISTORY)", st.stockhistory(B("Ticker"), 2), st.FMT["usd_price"],
        note="Latest close of the US listing from Excel (Microsoft 365, internet)")
    put("Price date", "Price date", st.stockhistory(B("Ticker"), 1), "yyyy-mm-dd")
    put("Price override", "Price override ($)", None, st.FMT["usd_price"], "optional", reg=("Market", "Price override"),
        note="Optional: type a price here to use instead of the live close")
    if foreign:
        put("FX rate", f"{cur} per US dollar (live)", st.stockhistory(f'"USD:{cur}"', 2), st.FMT["fx"],
            note=f"The filings are in {cur}; the US listing trades in dollars")
        put("FX override", f"{cur} per US dollar override", None, st.FMT["fx"], "optional",
            reg=("Market", "FX override"), note="Optional: type a rate to use instead of the live one")
        put("Shares per listing", "Ordinary shares per US-listed share", fin.adr_ratio, "General", "input",
            reg=("Market", "Shares per listing"),
            note="ADR ratio: 1 unless one US-listed share represents several ordinary shares (e.g. 5 for TSM)")
    else:
        put("FX rate", "Currency conversion", 1, st.FMT["fx"], note="1: filings and share price are both in US dollars")
        put("FX override", "Currency conversion override", None, st.FMT["fx"], "optional", reg=("Market", "FX override"))
        put("Shares per listing", "Ordinary shares per listed share", 1, "General", "input",
            reg=("Market", "Shares per listing"))
    fx = f'IF({B("FX override")}<>"",{B("FX override")},{B("FX rate")})'
    put("Share price", "Share price ($)" if not foreign else f"Share price ({cur} per ordinary share)",
        f'=IFERROR(IF({B("Price override")}<>"",{B("Price override")},{B("Live close")})*{fx}/{B("Shares per listing")},"")',
        st.FMT["price"], bold=True,
        check=f'=IF(N({B("Share price")})=0,"No price: check the ticker prefix or type an override","")',
        note="Used by the valuation and market cap" + (
            f": dollar price x {cur} per dollar / ordinary shares per listed share" if foreign else ""))
    shares, as_of, how = fin.shares_out
    c = put("Shares filed", "Shares outstanding (filings)", shares, st.FMT["shares"],
        note=(f"Source: {how}" + (f", as of {as_of}" if as_of else "")) if shares else
        "NOT FOUND in the SEC data (common with several share classes): type the count in the override below")
    c.font = st.font(st.BLUE)                                             # data from the filings
    put("Shares override", "Shares outstanding override", None, st.FMT["shares"], "optional",
        reg=("Market", "Shares outstanding"),
        note="Optional, in full (e.g. 1900000000): replaces the filed count, which stays above")
    put("Shares outstanding", "Shares outstanding used",
        f'=IF({B("Shares override")}<>"",{B("Shares override")},IF({B("Shares filed")}="","",{B("Shares filed")}))',
        st.FMT["shares"], bold=True,
        check=f'=IF(N({B("Shares outstanding")})=0,"Needed for a value per share: type it in the override","")')
    put("Market cap", "Market cap",
        f'=IF(N({B("Share price")})=0,"",{B("Share price")}*{B("Shares outstanding")})', st.FMT["usd"])
    put("52w low", "52-week low ($)", st.stock_range(B("Ticker"), "MIN"), st.FMT["usd_price"])
    put("52w high", "52-week high ($)", st.stock_range(B("Ticker"), "MAX"), st.FMT["usd_price"])

    # ---- Dilution (treasury stock method for options; stock units count in full). With a Bridge tab the inputs
    # are typed there and listed here as links.
    if m.bridge:
        from . import xl_bridge as br
        for key, label, bkey, fmt in (("Options", "Stock options outstanding", "options", "shares"),
                                      ("Option strike", f"Average exercise price ({'$' if not foreign else cur})",
                                       "strike", "price"),
                                      ("RSUs", "Unvested stock units (RSUs / PSUs)", "rsus", "shares"),
                                      ("Diluted shares", "Diluted shares used in the valuation", "diluted", "shares")):
            src = br.ref(bkey)
            c = put(key, label, f'=IF({src}="","",{src})', st.FMT[fmt], bold=key == "Diluted shares",
                    note="Set on the Bridge tab (click the name)" + (": shares outstanding + stock units + options "
                                                                    "in the money" if key == "Diluted shares" else ""))
            c.font = st.font(st.GREEN, key == "Diluted shares")
            st.link_to(ws.cell(row=R[key], column=1), br.SHEET, f"{'D' if bkey != 'diluted' else 'B'}{br.K[bkey]}")
    d = fin.dilution or {}
    if m.bridge:
        d = None
    filed = lambda v, f: f"Filed in the latest annual report: {f.format(v)}. " if v else ""
    if d is not None:
        _dilution(put, B, d, filed, foreign, cur)

    # ---- settings collected while the valuation tabs were written (defaults, notes and dropdown options)
    has_rate = m.kind == "fin" or "dcf" in m.parts
    rate = xl_dcf.ref("wacc") if m.kind != "fin" else f"({B('rf')}+{B('beta')}*{B('erp')})"     # WACC / cost of equity
    rate_name = "WACC" if m.kind != "fin" else "the cost of equity"
    checks = {
        "rf": f'=IF(OR({B("rf")}<0.01,{B("rf")}>0.08),"Outside 1%-8%","")',
        "erp": f'=IF(OR({B("erp")}<0.03,{B("erp")}>0.07),"Outside the usual 3%-7%","")',
        "beta": f'=IF(OR({B("beta")}<0.5,{B("beta")}>2),"Unusual: outside 0.5-2.0","")',
        "g": (f'=IF({B("g")}>={rate},"At or above {rate_name}: the perpetuity value breaks",'
              f'IF({B("g")}>{B("rf")},"Above the risk-free rate: hard to sustain forever",""))' if has_rate else
              f'=IF({B("g")}>0.04,"High for growth that lasts forever","")'),
        "multiple": (f'=IF(N({B("multiple")})>30,"High: compare with the Comps tab","")' if m.kind != "fin" else
                     f'=IF(N({B("multiple")})>4,"High for a price-to-book multiple","")'),
        "roic_t": f'=IF({B("roic_t")}<{rate},"Below WACC: growth destroys value","")',
    }
    placed = ("rf", "erp", "beta", "kd", "g", "multiple", "tv_mode", "roic_t", "mid", "sbc_mode", "capex_mode",
              "lease_mode", "val_override", "other_assets", "other_claims")
    for key in placed:
        if key not in m.settings:               # the financials template has fewer settings
            continue
        s = m.settings[key]
        if s.get("home"):
            # typed on the tab it acts on: shown here as a link to the value in use, with a click-through
            sheet, addr = s["home"]
            src = s["used"]
            c = put(key, s["label"], f'=IF({src}="","",{src})', None if s.get("options") else st.FMT[s["fmt"]],
                    note=f"Set on the {sheet} tab (click the name): type in its yellow override cell. " + s["note"],
                    check=checks.get(key))
            c.font = st.font(st.GREEN)
            st.link_to(ws.cell(row=R[key], column=1), sheet, addr)
        else:
            optional = s["value"] is None
            c = put(key, s["label"], s["value"], None if s.get("options") else st.FMT[s["fmt"]],
                    "optional" if optional else "input", note=s["note"], reg=("DCF", s["label"]),
                    options=s.get("options"), check=checks.get(key))
        if key == "mid":
            c.number_format = "0"
    ws.cell(row=R["beta"], column=1).font = st.font(bold=True)
    if m.kind != "fin":                     # beta is measured on the WACC tab (or not needed without a DCF)
        for key in ("Beta benchmark", "Beta years", "Beta frequency", "Computed beta", "Adjusted beta"):
            ws.row_dimensions[R[key]].hidden = True
    if not has_rate:
        # built without the DCF: only the long-run growth rate is used, so the valuation settings are hidden
        ws.cell(row=R["g"], column=1).value = "Long-run growth (revenue growth fades toward it)"
        for key in ("rf", "erp", "Beta benchmark", "Beta years", "Beta frequency", "Computed beta", "Adjusted beta",
                    "beta", "kd", "multiple", "tv_mode", "roic_t", "mid", "sbc_mode", "capex_mode", "lease_mode",
                    "val_override", "other_assets", "other_claims"):
            ws.row_dimensions[R[key]].hidden = True
        for title in ("Discount rate", "Method choices", "Equity bridge adjustments"):
            ws.row_dimensions[INPUT_HEADERS[title]].hidden = True
        ws.cell(row=INPUT_HEADERS["Terminal value"], column=1).value = "Forecast"
        _index(ws, m)
        _finish(ws)
        return

    if m.kind != "fin":
        _index(ws, m)
        _finish(ws)
        return

    # ---- beta: measured from prices, then adjusted (the financials template; the standard model does this on
    # the WACC tab)
    def pick(key, text, value, options, note):
        put(key, text, value, kind="input", options=options, reg=("Market", text), note=note)

    put("Beta benchmark", "Beta benchmark", "SPY", kind="input", reg=("Market", "Beta benchmark"),
        note="Index ETF the beta is measured against")
    pick("Beta years", "Beta window (years)", 5, (2, 5), "2 or 5 years of history")
    pick("Beta frequency", "Beta return frequency", "Monthly", ("Weekly", "Monthly"),
         "5 years of monthly returns is the usual convention; 2 years of weekly reacts faster to recent trading")
    # TAKE / DROP keep this a plain formula: Excel adds an implicit-intersection "@" to INDEX(SEQUENCE()) in
    # formulas loaded from a file, which breaks it.
    px = lambda ref: (f"_xlfn.STOCKHISTORY({ref},TODAY()-ROUND(365.25*{B('Beta years')},0),TODAY(),"
                      f'IF({B("Beta frequency")}="Weekly",1,2),0,1)')
    ret = lambda v: f"_xlfn.DROP(_xlpm.{v},1)/_xlfn.DROP(_xlpm.{v},-1)-1"
    beta = ("=IFERROR(_xlfn.LET("
            f"_xlpm.a,{px(B('Ticker'))},_xlpm.b,{px(B('Beta benchmark'))},"
            "_xlpm.n,MIN(ROWS(_xlpm.a),ROWS(_xlpm.b)),"
            "_xlpm.x,_xlfn.TAKE(_xlpm.a,-_xlpm.n),_xlpm.y,_xlfn.TAKE(_xlpm.b,-_xlpm.n),"
            f'SLOPE({ret("x")},{ret("y")})),"")')
    put("Computed beta", "Raw beta (regression)", beta, st.FMT["num"],
        note="Slope of the stock's returns on the benchmark's over the window above")
    put("Adjusted beta", "Adjusted beta (Blume)",
        f'=IF(ISNUMBER({B("Computed beta")}),2/3*{B("Computed beta")}+1/3,"")', st.FMT["num"],
        note="2/3 x raw + 1/3 x 1.0: measured betas drift toward the market's 1.0 over time")

    _index(ws, m)
    _finish(ws)


OVERRIDE_HEADS = ("Used", "Default", "Your override", "Check")


def override_header(ws, r):
    """Column heads for a block of settings written with setting(): B used, C default, D override, E check."""
    for j, h in enumerate(OVERRIDE_HEADS, start=2):
        st.header_cell(ws.cell(row=r, column=j), h, "ltm" if h == "Your override" else "fcst")


def setting(ws, m, r, key, label, default, fmt, note="", options=None, reg=None):
    """One setting in the override pattern, on the tab it acts on:
        B  used      = the override when there is one, else the default (formulas read this cell)
        C  default   = what the model suggests (a number or a formula); never typed over
        D  override  = yellow, blank: type here to replace the default, clear it to go back
        E  check     = the range warning worked out on the Inputs tab
        F  note
    The Inputs tab lists it (from m.settings) with a link back here."""
    sheet = ws.title
    q = f"'{sheet}'" if " " in sheet else sheet
    number = st.FMT[fmt] if not options else "General"
    ws.cell(row=r, column=1, value=label)
    used = ws.cell(row=r, column=2, value=f'=IF(D{r}="",IF(C{r}="","",C{r}),D{r})')
    used.number_format, used.font = number, st.font(bold=True)
    d = ws.cell(row=r, column=3, value=default)
    d.number_format = number
    d.font = st.formula_font(default) if isinstance(default, str) and default.startswith("=") else st.font()
    o = ws.cell(row=r, column=4)
    o.number_format = number
    m.reg.add(ws, o, *(reg or ("DCF", label)))
    if options:
        dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(o)
        for c in (used, d, o):
            c.alignment = st.CENTER
    if key in R:
        chk = f"{SHEET_NAMES['INP']}!$C${R[key]}"
        ws.cell(row=r, column=5, value=f'=IF({chk}<>"",{chk},"")').font = WARN
    if note:
        st.hover_note(ws.cell(row=r, column=6), note).font = st.NOTE_FONT
    m.settings[key] = {"label": label, "value": default, "fmt": fmt, "note": note, "options": options,
                       "home": (sheet, f"D{r}"), "used": f"{q}!$B${r}"}
    return used


def _index(ws, m):
    """Inputs that live on other tabs, with click-throughs."""
    from .layout import RB_SHEET, RB_SWITCH
    from . import xl_multiples
    r = max(R.values()) + 2
    c = ws.cell(row=r, column=1, value="Inputs on other tabs")
    c.font, c.fill = st.HDR_FONT, st.FILL["fcst"]
    for j in (2, 3, 4):
        ws.cell(row=r, column=j).fill = st.FILL["fcst"]
    g = m.grid
    if m.kind == "fin":
        from .xl_fin import SCEN_TITLE, SHEET
        items = [(SHEET, f"B{SCEN_TITLE + 3}", "Bear / Base / Bull drivers",
                  "Return on equity, payout and share count for each scenario")]
    else:
        items = [(SHEET_NAMES["SCN"], f"B{g.scen_ovr_first}", "Bear / Base / Bull cases",
                  "Year 1, mid-point and final-year growth and margins for each scenario, and your overrides"),
                 ("Drivers", f"{g.fcst[0].L}{g.year_ovr_rows['Revenue growth']}", "Per-year driver overrides",
                  "Fix one driver in one forecast year")]
    if "dcf" in m.parts and m.kind != "fin":
        from .xl_wacc import SHEET as WS, W
        items.append((WS, f"D{W['rf']}", "WACC settings", "Risk-free rate, equity risk premium, beta and its window, "
                                                          "cost of debt"))
        items.append(("DCF", f"D{xl_dcf.R['g']}", "DCF settings", "Terminal growth, exit multiple, method choices, equity bridge"))
    if "multiples" in m.parts and m.kind != "fin":
        items.append(("Multiples", f"{xl_multiples.COL['override']}{xl_multiples.R['pe']}", "Target multiples",
                      "Forward P/E, PEG, EV/EBITDA, EV/Sales and P/FCF targets, ranges and On / Off"))
    if m.kind != "fin" and ({"dcf", "multiples"} & m.parts or m.comps):
        items.append(("Summary", "G13", "Price target weights", "How much each valuation method counts"))
    if "revenue_build" in m.parts and m.kind != "fin":
        items.append((RB_SHEET, RB_SWITCH.replace("$", ""), "Segment revenue build", "Segments, their revenue and growth"))
    if m.comps:
        items.append(("Comps", "A1", "Peer price overrides", "The yellow column on the Comps tab"))
    items.append(("Notes", "A5", "Your notes", "Thesis, risks and sources: kept across rebuilds"))
    for i, (sheet, addr, text, note) in enumerate(items, start=1):
        st.link_to(ws.cell(row=r + i, column=1), sheet, addr, text)
        ws.cell(row=r + i, column=4, value=note).font = st.NOTE_FONT


def _finish(ws):
    ws.conditional_formatting.add(f"C5:C{max(R.values())}", FormulaRule(formula=['LEN(C5)>0'], fill=st.WARN_FILL))
    for L, w in zip("ABCD", (44, 26, 46, 110)):
        ws.column_dimensions[L].width = w
    ws.freeze_panes = "B4"


def _dilution(put, B, d, filed, foreign, cur):
    """Dilution typed on Inputs: used when there is no Bridge tab (the financials template, statements only)."""
    put("Options", "Stock options outstanding", d.get("options"), "#,##0", "optional",
        note=filed(d.get("options"), "{:,.0f}") + "Optional, in full. Left blank = no dilution from options")
    put("Option strike", f"Average exercise price ({'$' if not foreign else cur})", d.get("strike"), st.FMT["price"],
        "optional", note=filed(d.get("strike"), "{:,.2f}") + "Options only dilute to the extent the share price is "
                                                             "above this")
    put("RSUs", "Unvested stock units (RSUs / PSUs)", d.get("rsus"), "#,##0", "optional",
        note=filed(d.get("rsus"), "{:,.0f}") + "Optional, in full: counted one for one. Usually the larger item "
                                               "for tech companies")
    put("Diluted shares", "Diluted shares used in the valuation",
        f'=IF(N({B("Shares outstanding")})=0,"",{B("Shares outstanding")}+N({B("RSUs")})'
        f'+IF(AND(N({B("Options")})>0,N({B("Share price")})>0),'
        f'MAX(0,N({B("Options")})*(1-N({B("Option strike")})/{B("Share price")})),0))', st.FMT["shares"], bold=True,
        note="Shares outstanding + stock units + options x (1 - exercise price / share price)")


def warnings_count(sheet=None):
    """Formula text counting the non-blank cells in the Inputs tab's Check column."""
    name = SHEET_NAMES["INP"]
    return f'COUNTIF({name}!$C$5:$C${max(R.values())},"?*")'


NOTE_SECTIONS = (("Thesis", 8), ("Key assumptions and why", 10), ("Risks and what would change my mind", 8),
                 ("Sources and to-do", 10))


def note_lines(title):
    """Rows of the typed lines under a Notes heading."""
    r = 4
    for t, n in NOTE_SECTIONS:
        if t == title:
            return range(r + 1, r + 1 + n)
        r += n + 2
    raise KeyError(title)


def write_notes(ws, m):
    """A place for the thesis and sources that survives rebuilds (column A only)."""
    ws["A1"], ws["A1"].font = f"{m.fin.name} ({m.fin.ticker}) - Notes", st.TITLE_FONT
    ws["A2"] = ("Your own notes. Type in column A under each heading: these lines are carried over when the "
                "workbook is rebuilt. Anything typed elsewhere on this tab, and any tab you add yourself, is not.")
    ws["A2"].font = st.NOTE_FONT
    r = 4
    for title, lines in NOTE_SECTIONS:
        st.header_cell(ws.cell(row=r, column=1), title)
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="left")
        for k in range(1, lines + 1):
            c = ws.cell(row=r + k, column=1)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            m.reg.add(ws, c, "Notes", title, k)
        r += lines + 2
    ws.column_dimensions["A"].width = 140
