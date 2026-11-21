"""Financials template: banks, insurers and asset managers.

Free cash flow, EBITDA and net debt don't describe these businesses: debt and deposits are raw material, not
financing. They are valued on equity instead. The forecast is a return on equity and a payout; value is
    residual income:  book equity + present value of (net income - cost of equity x equity)
    exit P/B:         present value of dividends and buybacks + final-year equity x a price-to-book multiple
Tabs: Summary | Inputs | Equity Model | Notes | Mapping | Raw. Inputs, scenarios, the roll-up and the write-up
work the same way as in the standard model."""
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from . import xl_style as st
from .layout import BASIS_OPTIONS, SCENARIOS, input_ref
from .mapping import fmt_tag
from .xl_inputs import warnings_count, write_inputs, write_notes

SHEET = "Equity Model"
DATA = [  # (row label, SEC line(s) summed, number format)
    ("Revenue", ("Revenue",), "usd"), ("Net interest income", ("Net Interest Income",), "usd"),
    ("Provision for credit losses", ("Provision for Credit Losses",), "usd"), ("Pre-tax income", ("Pre-Tax Income",), "usd"),
    ("Net income", ("Net Income",), "usd"), ("Diluted EPS", ("Diluted EPS",), "eps"),
    ("Diluted shares", ("Diluted Shares",), "shares"), ("Dividends paid", ("Dividends Paid",), "usd"),
    ("Share repurchases", ("Share Repurchases",), "usd"),
    ("Shareholders' equity", ("Shareholders' Equity (Parent)",), "usd"),
    ("Goodwill and intangibles", ("Goodwill", "Intangibles, Net"), "usd"), ("Tangible equity", (), "usd"),
    ("Total assets", ("Total Assets",), "usd"), ("Deposits", ("Deposits",), "usd"),
]
ROW = {label: 5 + i for i, (label, _, _) in enumerate(DATA)}
RATIO_HDR = 5 + len(DATA) + 1
RATIOS = ["Return on equity", "Return on tangible equity", "Dividends % of net income", "Buybacks % of net income",
          "Diluted share count change", "Equity / assets", "Book value per share", "Tangible book value per share"]
ROW.update({label: RATIO_HDR + 1 + i for i, label in enumerate(RATIOS)})
DRIVERS = RATIOS[0:1] + RATIOS[2:5]                       # the four forecast drivers
SCEN_TITLE = ROW[RATIOS[-1]] + 2
SCEN = {d: SCEN_TITLE + 3 + i for i, d in enumerate(DRIVERS)}
V = SCEN_TITLE + 3 + len(DRIVERS) + 1                    # "Valuation" block
VR = {k: V + 1 + i for i, k in enumerate(
    ["ke", "g", "pb", "val_date", "elapsed", "_", "year", "charge", "ri", "factor", "pv_ri", "dist", "pv_dist", "__",
     "hdr", "start", "sum_pv", "tv", "pv_tv", "at_fy", "rolled", "per_share", "price", "upside", "pb_implied"])}
GRID1 = VR["pb_implied"] + 3
GRID2 = GRID1 + 9
GRID_CELLS = tuple(f"$C${t + 2}:$G${t + 6}" for t in (GRID1, GRID2))
REF = lambda cell: f"'{SHEET}'!{cell}"


def build_financial(fin, n_fcst=10):
    from .xl_model import Model, write_header, write_mapping, write_raw
    m = Model(fin, n_fcst, ytd=False)
    m.kind = "fin"
    wb = Workbook()
    names = ["Summary", "Inputs", SHEET, "Notes", "Mapping", "Raw"]
    wb.active.title = names[0]
    for n in names[1:]:
        wb.create_sheet(n)
    _model(wb[SHEET], m, write_header)
    write_inputs(wb["Inputs"], m)
    _summary(wb["Summary"], m)
    write_notes(wb["Notes"], m)
    _mapping(wb["Mapping"], fin)
    write_raw(wb["Raw"], fin.sources)
    for name, color in (("Summary", "1F3864"), ("Inputs", "BF8F00"), (SHEET, "7030A0"), ("Notes", "808080"),
                        ("Mapping", "808080"), ("Raw", "808080")):
        wb[name].sheet_properties.tabColor = color
    for ws in wb.worksheets:                                # the house style: title bars, no gridlines
        if ws.title not in ("Mapping", "Raw"):
            st.title_bar(ws)
    wb["Summary"].sheet_properties.tabColor = st.THEME["accent"]
    m.reg.write(wb)
    wb.active = 0
    return wb, m


def _value(m, lines, col):
    vals = [m.data.get(l, {}).get(col.j) for l in lines]
    return sum(v for v in vals if v is not None) if any(v is not None for v in vals) else None


def _model(ws, m, write_header):
    g, fin = m.grid, m.fin
    write_header(ws, g, f"{fin.name} ({fin.ticker}) - {SHEET}",
                 "Financial companies are valued on equity: forecast = return on equity and payout; value = book "
                 "equity plus the present value of returns above the cost of equity. Blue = SEC filings, black = "
                 "formula; yellow = inputs you can change. Settings are on the Inputs tab.")
    R = ROW
    r_of = lambda label, col: f"{col.L}{R[label]}"
    last, n = g.last_fy, len(g.fcst)
    past = [c for c in g.cols if c.kind != "fcst"]

    # ---- reported lines, and their forecast formulas
    for label, lines, fmt in DATA:
        r = R[label]
        ws.cell(row=r, column=1, value=label).font = st.font(bold=label in ("Net income", "Shareholders' equity"))
        for col in g.cols:
            c = ws.cell(row=r, column=col.j)
            c.number_format = st.FMT[fmt]
            prev = g.prev(col)
            if col.kind == "fcst":
                c.fill = st.FCST_BG
                ni, eq = r_of("Net income", col), r_of("Shareholders' equity", col)
                f = {"Net income": f"={r_of('Return on equity', col)}*N({r_of(chr(83) + 'hareholders' + chr(39) + ' equity', prev)})",
                     "Diluted EPS": f'=IFERROR({ni}/{r_of("Diluted shares", col)},"")',
                     "Diluted shares": f"=N({r_of('Diluted shares', prev)})*(1+{r_of('Diluted share count change', col)})",
                     "Dividends paid": f"=MAX(0,{ni})*{r_of('Dividends % of net income', col)}",
                     "Share repurchases": f"=MAX(0,{ni})*{r_of('Buybacks % of net income', col)}",
                     "Shareholders' equity": (f"=N({r_of(label, prev)})+{ni}-{r_of('Dividends paid', col)}"
                                              f"-{r_of('Share repurchases', col)}"),
                     "Goodwill and intangibles": f"=N({r_of(label, prev)})",
                     "Tangible equity": f"={eq}-N({r_of('Goodwill and intangibles', col)})"}.get(label)
                if f:
                    c.value, c.font = f, st.font(bold=label in ("Net income", "Shareholders' equity"))
            elif label == "Tangible equity":
                eq = r_of("Shareholders' equity", col)
                c.value = f'=IF({eq}="","",{eq}-N({r_of("Goodwill and intangibles", col)}))'
            else:
                v = _value(m, lines, col)
                if v is not None:
                    c.value, c.font = v, st.font(st.BLUE, label in ("Net income", "Shareholders' equity"))
        if not any(_value(m, lines, c) is not None for c in past) and lines:
            ws.cell(row=r, column=1).value = f"{label}  (not reported)"
            ws.cell(row=r, column=1).font = st.font(st.GREY, italic=True)

    # ---- ratios (history) and drivers (forecast paths)
    st.header_cell(ws.cell(row=RATIO_HDR, column=1), "Ratios and forecast drivers", "ltm")
    ws.cell(row=RATIO_HDR, column=1).alignment = Alignment(horizontal="left")
    mid = (n + 1) // 2
    for label in RATIOS:
        r = R[label]
        ws.cell(row=r, column=1, value=label).font = st.font(bold=label in DRIVERS)
        fmt = "pct" if "per share" not in label else "price"
        for col in g.cols:
            c = ws.cell(row=r, column=col.j)
            c.number_format = st.FMT[fmt]
            prev = g.prev(col) if col.kind != "ltm" else last
            ni, eq, sh = (r_of(x, col) for x in ("Net income", "Shareholders' equity", "Diluted shares"))
            if col.kind == "fcst" and label in DRIVERS:
                k, sr = g.fcst.index(col) + 1, SCEN[label]
                y1, ym, yn = (f"${L}${sr}" for L in "KLM")
                c.value = (f"={y1}" if k == 1 else f"={y1}+({ym}-{y1})*{k - 1}/{mid - 1}" if k <= mid
                           else f"={ym}+({yn}-{ym})*{k - mid}/{n - mid}")
                c.fill = st.FCST_BG
                m.reg.add(ws, c, "Live", label, col.header, mark=False)
                continue
            if col.kind == "fcst":
                c.fill = st.FCST_BG
            f = None
            if label == "Return on equity" and prev is not None:
                f = f'=IF(OR({ni}="",N({r_of(chr(83) + "hareholders" + chr(39) + " equity", prev)})=0),"",{ni}/{r_of(chr(83) + "hareholders" + chr(39) + " equity", prev)})'
            elif label == "Return on tangible equity" and prev is not None:
                f = f'=IFERROR(IF({ni}="","",{ni}/{r_of("Tangible equity", prev)}),"")'
            elif label == "Dividends % of net income":
                f = f'=IF(N({ni})<=0,"",N({r_of("Dividends paid", col)})/{ni})'
            elif label == "Buybacks % of net income":
                f = f'=IF(N({ni})<=0,"",N({r_of("Share repurchases", col)})/{ni})'
            elif label == "Diluted share count change" and col.kind == "hist" and prev is not None:
                f = f'=IF(OR(N({sh})=0,N({r_of("Diluted shares", prev)})=0),"",{sh}/{r_of("Diluted shares", prev)}-1)'
            elif label == "Equity / assets":
                f = f'=IF(OR({eq}="",N({r_of("Total assets", col)})=0),"",{eq}/{r_of("Total assets", col)})'
            elif label == "Book value per share":
                f = f'=IF(OR({eq}="",N({sh})=0),"",{eq}/{sh})'
            elif label == "Tangible book value per share":
                f = f'=IFERROR(IF(N({sh})=0,"",{r_of("Tangible equity", col)}/{sh}),"")'
            if f and not (col.kind == "fcst" and label == "Equity / assets"):
                c.value, c.font = f, st.font()
    first_d, last_d = R[DRIVERS[0]], R[DRIVERS[-1]]
    top = f"{g.fcst[0].L}{first_d}"
    ws.conditional_formatting.add(f"{top}:{g.fcst[-1].L}{last_d}", FormulaRule(
        formula=[f'AND(NOT(ISFORMULA({top})),{top}<>"")'], font=st.font(st.BLUE, True), fill=st.INPUT_FILL))

    # ---- scenario inputs
    t = SCEN_TITLE
    ke = f"$B${VR['ke']}"
    ws.cell(row=t, column=1, value="Scenario inputs").font = st.font(bold=True)
    ws.cell(row=t, column=2, value="Each driver runs in straight lines from year 1 to the mid-point to the final year. "
                                   "Base year 1 comes from history; return on equity fades halfway back to the cost "
                                   "of equity by the final year; Bear and Bull default to Base return on equity "
                                   "-/+ 10%. Overwrite any yellow cell.").font = st.NOTE_FONT
    st.header_cell(ws.cell(row=t + 1, column=1), "")
    st.header_cell(ws.cell(row=t + 2, column=1), "Driver")
    heads = (g.fcst[0].header, g.fcst[mid - 1].header, g.fcst[-1].header)
    for i, name in enumerate(SCENARIOS + ("Selected",)):
        for k, sub in enumerate(heads):
            for rr, text in ((t + 1, name), (t + 2, sub)):
                st.header_cell(ws.cell(row=rr, column=2 + 3 * i + k), text, "ltm" if name == "Selected" else "label")
    hist3 = g.hist[-3:]
    basis = f'MATCH({input_ref("Basis")},{{{",".join(chr(34) + o + chr(34) for o in BASIS_OPTIONS)}}},0)'
    pick = f'MATCH({input_ref("Scenario")},{{{",".join(chr(34) + s + chr(34) for s in SCENARIOS)}}},0)'
    for d in DRIVERS:
        r, sr = R[d], SCEN[d]
        ws.cell(row=sr, column=1, value=d)
        fy = f'IF({last.L}{r}="",0,{last.L}{r})'
        ltm = f'IF({g.ltm.L}{r}="",{fy},{g.ltm.L}{r})' if g.ltm is not None and d == "Return on equity" else fy
        base1 = f"=IFERROR(CHOOSE({basis},AVERAGE({hist3[0].L}{r}:{hist3[-1].L}{r}),{fy},{ltm}),0)"
        base = (base1, f"=E{sr}", f"=(F{sr}+{ke})/2" if d == "Return on equity" else f"=F{sr}")
        shift = lambda cell, sign: f"={cell}*(1{'+' if sign > 0 else '-'}0.1)" if d == "Return on equity" else f"={cell}"
        values = {"Bear": tuple(shift(f"{L}{sr}", -1) for L in "EFG"), "Base": base,
                  "Bull": tuple(shift(f"{L}{sr}", 1) for L in "EFG")}
        for i, name in enumerate(SCENARIOS):
            for k, sub in enumerate(("Year 1", "Mid", "Final year")):
                c = ws.cell(row=sr, column=2 + 3 * i + k, value=values[name][k])
                c.number_format, c.font = st.FMT["pct"], st.font(st.BLUE)
                m.reg.add(ws, c, "Scenario", d, name, sub)
        for k, cols in enumerate((("B", "E", "H"), ("C", "F", "I"), ("D", "G", "J"))):
            c = ws.cell(row=sr, column=11 + k, value=f"=CHOOSE({pick}," + ",".join(f"{L}{sr}" for L in cols) + ")")
            c.number_format, c.font = st.FMT["pct"], st.font(bold=True)

    # ---- settings (shown here, set on Inputs)
    rf, rf_day = fin.rf
    erp, erp_src = fin.erp
    S = m.settings
    S["rf"] = {"label": "Risk-free rate", "value": rf or 0.0425, "fmt": "pct",
               "note": f"10-year Treasury yield as of {rf_day}" if rf else "Default: the yield could not be downloaded"}
    S["erp"] = {"label": "Equity risk premium", "value": erp or 0.055, "fmt": "pct",
                "note": erp_src if erp else "Default: set \"equity_risk_premium\" in watchlist.json"}
    S["beta"] = {"label": "Beta", "value": f"=IFERROR(ROUND({input_ref('Adjusted beta')},2),1)", "fmt": "num",
                 "note": "Defaults to the adjusted beta below (1.0 without live prices). Type a number to override"}
    S["g"] = {"label": "Terminal growth", "value": 0.025, "fmt": "pct",
              "note": "Growth of residual income after the final forecast year"}
    eq0, sh_in, px_in = f"${last.L}${R['Shareholders' + chr(39) + ' equity']}", input_ref("Diluted shares"), input_ref("Share price")
    S["multiple"] = {"label": "Exit price-to-book multiple",
                     "value": f"=IFERROR(MAX(1,MIN(4,ROUND({px_in}*{sh_in}/{REF(eq0)},1))),1.5)", "fmt": "x",
                     "note": "Defaults to today's price-to-book, held between 1.0x and 4.0x (1.5x without a price)"}
    S["val_override"] = {"label": "Valuation date override", "value": None, "fmt": "date", "note": "Optional: blank = today"}

    def kv(key, label, f, fmt, note=""):
        r = VR[key]
        ws.cell(row=r, column=1, value=label)
        c = ws.cell(row=r, column=2, value=f)
        c.number_format, c.font = st.FMT[fmt], st.formula_font(f) if isinstance(f, str) else st.font()
        if note:
            ws.cell(row=r, column=3, value=note).font = st.NOTE_FONT

    st.header_cell(ws.cell(row=V, column=1), "Valuation")
    kv("ke", "Cost of equity", f"={input_ref('rf')}+{input_ref('beta')}*{input_ref('erp')}", "pct",
       "Risk-free rate + beta x equity risk premium (Inputs tab)")
    kv("g", "Terminal growth", f"={input_ref('g')}", "pct")
    kv("pb", "Exit price-to-book multiple", f"={input_ref('multiple')}", "x")
    vo = input_ref("val_override")
    kv("val_date", "Valuation date", f'=IF({vo}="",TODAY(),{vo})', "date")
    ws.cell(row=VR["elapsed"], column=4, value=g.last_fy.end.date()).number_format = st.FMT["date"]
    ws.cell(row=VR["elapsed"], column=3, value="years since the fiscal year end ->").font = st.NOTE_FONT
    kv("elapsed", "Years since the last fiscal year end", f"=MAX(0,MIN(2,($B${VR['val_date']}-$D${VR['elapsed']})/365))", "num")
    ws.cell(row=VR["elapsed"], column=3).alignment = Alignment(horizontal="right")
    first_L, last_L = g.fcst[0].L, g.fcst[-1].L
    per_year = [("year", "Year", lambda c, k: k + 1, "0"),
                ("charge", "Charge for equity (cost of equity x opening equity)",
                 lambda c, k: f"={ke}*N({r_of('Shareholders' + chr(39) + ' equity', g.prev(c))})", st.FMT["usd"]),
                ("ri", "Residual income (net income - charge)", lambda c, k: f"={r_of('Net income', c)}-{c.L}{VR['charge']}", st.FMT["usd"]),
                ("factor", "Discount factor", lambda c, k: f"=1/(1+{ke})^{c.L}{VR['year']}", "0.00"),
                ("pv_ri", "PV of residual income", lambda c, k: f"={c.L}{VR['ri']}*{c.L}{VR['factor']}", st.FMT["usd"]),
                ("dist", "Dividends and buybacks", lambda c, k: f"={r_of('Dividends paid', c)}+{r_of('Share repurchases', c)}", st.FMT["usd"]),
                ("pv_dist", "PV of dividends and buybacks", lambda c, k: f"={c.L}{VR['dist']}*{c.L}{VR['factor']}", st.FMT["usd"])]
    for key, label, f, fmt in per_year:
        ws.cell(row=VR[key], column=1, value=label).font = st.font(bold=key == "ri")
        for k, c in enumerate(g.fcst):
            cell = ws.cell(row=VR[key], column=c.j, value=f(c, k))
            cell.number_format, cell.fill = fmt, st.FCST_BG
    rng = lambda key: f"${first_L}${VR[key]}:${last_L}${VR[key]}"
    N, gT, pb = f"${last_L}${VR['year']}", f"$B${VR['g']}", f"$B${VR['pb']}"
    eqN, riN = f"${last_L}${R['Shareholders' + chr(39) + ' equity']}", f"${last_L}${VR['ri']}"
    st.header_cell(ws.cell(row=VR["hdr"], column=1), "Equity value")
    st.header_cell(ws.cell(row=VR["hdr"], column=2), "Residual income", "fcst")
    st.header_cell(ws.cell(row=VR["hdr"], column=3), "Exit P/B", "fcst")
    C = lambda key, col: f"{col}{VR[key]}"
    lines = [("start", f"Book equity, {last.header}", f"=N({eq0})", None, "usd"),
             ("sum_pv", "PV of the forecast years", f"=SUM({rng('pv_ri')})", f"=SUM({rng('pv_dist')})", "usd"),
             ("tv", "Terminal value (end of final year)", f"={riN}*(1+{gT})/({ke}-{gT})", f"={eqN}*{pb}", "usd"),
             ("pv_tv", "PV of terminal value", f"=B{VR['tv']}/(1+{ke})^{N}", f"=C{VR['tv']}/(1+{ke})^{N}", "usd"),
             ("at_fy", f"Equity value at {last.header} year end", f"=B{VR['start']}+B{VR['sum_pv']}+B{VR['pv_tv']}",
              f"=C{VR['sum_pv']}+C{VR['pv_tv']}", "usd"),
             ("rolled", "Rolled forward to the valuation date", None, None, "usd"),
             ("per_share", "Value per share", None, None, "price"), ("price", "Current share price", None, None, "price"),
             ("upside", "Upside / (downside)", None, None, "pct"),
             ("pb_implied", "Value / book equity (x)", None, None, "x")]
    for key, label, fb, fc, fmt in lines:
        ws.cell(row=VR[key], column=1, value=label).font = st.font(bold=key in ("at_fy", "per_share"))
        for col, f in (("B", fb), ("C", fc)):
            f = {"rolled": f"={C('at_fy', col)}*(1+{ke})^$B${VR['elapsed']}",
                 "per_share": f'=IFERROR({C("rolled", col)}/N({sh_in}),"")',
                 "price": f'=IF(N({px_in})=0,"",{px_in})',
                 "upside": f'=IFERROR({C("per_share", col)}/{C("price", col)}-1,"")',
                 "pb_implied": f'=IFERROR({C("rolled", col)}/N({eq0}),"")'}.get(key, f)
            if f:
                c = ws[C(key, col)]
                c.value, c.number_format, c.font = f, st.FMT[fmt], st.font(bold=key in ("at_fy", "per_share"))
    ws.cell(row=VR["rolled"], column=4, value="Grown at the cost of equity for the time since the fiscal year end "
                                              "(dividends paid since then are ignored)").font = st.NOTE_FONT

    # ---- sensitivities: cost of equity x terminal growth, and x exit P/B
    ni_r, eq_prev = rng("ri").replace(str(VR["ri"]), str(R["Net income"])), None
    prev_cols = [g.prev(c) for c in g.fcst]
    eq_open = f"${prev_cols[0].L}${R['Shareholders' + chr(39) + ' equity']}:${prev_cols[-1].L}${R['Shareholders' + chr(39) + ' equity']}"
    yrs = rng("year")
    contiguous = prev_cols[0].j + len(prev_cols) - 1 == prev_cols[-1].j        # false when LTM sits between
    roll = f"*(1+{{w}})^$B${VR['elapsed']}/N({sh_in})"
    if contiguous:
        ri_pv = f"SUMPRODUCT(({ni_r}-{{w}}*{eq_open})/(1+{{w}})^{yrs})"
    else:        # opening equity of year 1 is the last fiscal year, which is not next to the forecast columns
        rest = (f"${g.fcst[0].L}${R['Shareholders' + chr(39) + ' equity']}:"
                f"${g.fcst[-2].L}${R['Shareholders' + chr(39) + ' equity']}")
        ni1, ni_rest = f"${first_L}${R['Net income']}", f"${g.fcst[1].L}${R['Net income']}:${last_L}${R['Net income']}"
        yr_rest = f"${g.fcst[1].L}${VR['year']}:${last_L}${VR['year']}"
        ri_pv = (f"(({ni1}-{{w}}*N({eq0}))/(1+{{w}})+SUMPRODUCT(({ni_rest}-{{w}}*{rest})/(1+{{w}})^{yr_rest}))")
    ri_last = f"(${last_L}${R['Net income']}-{{w}}*${g.fcst[-2].L}${R['Shareholders' + chr(39) + ' equity']})"
    _grid(ws, GRID1, "Value per share: cost of equity (rows) x terminal growth (columns)", gT, 0.005, "pct", ke,
          lambda w, x: ("=IFERROR((N(" + eq0 + ")+" + ri_pv.format(w=w) + "+" + ri_last.format(w=w)
                        + f"*(1+{x})/({w}-{x})/(1+{w})^{N})" + roll.format(w=w) + ',"")'))
    _grid(ws, GRID2, "Value per share: cost of equity (rows) x exit price-to-book (columns)", pb, 0.25, "x", ke,
          lambda w, x: (f"=IFERROR((SUMPRODUCT({rng('dist')}/(1+{w})^{yrs})+{eqN}*{x}/(1+{w})^{N})"
                        + roll.format(w=w) + ',"")'))
    ws.column_dimensions["A"].width = 50
    ws.column_dimensions[get_column_letter(g.cols[-1].j + 1)].width = 40


def _grid(ws, top, title, base_ref, step, fmt, ke, formula):
    ws.cell(row=top, column=1, value=title).font = st.font(bold=True)
    hdr = top + 1
    st.header_cell(ws.cell(row=hdr, column=2), "Cost of equity \\ ")
    offsets = (-2, -1, 0, 1, 2)
    for k, o in enumerate(offsets):
        c = ws.cell(row=hdr, column=3 + k, value=f"={base_ref}+{o * step}")
        c.number_format, c.font, c.fill = st.FMT[fmt], st.HDR_FONT, st.FILL["fcst"]
    for i, o in enumerate(offsets):
        r = hdr + 1 + i
        w = ws.cell(row=r, column=2, value=f"={ke}+{o * 0.005}")
        w.number_format, w.font, w.fill = st.FMT["pct"], st.HDR_FONT, st.FILL["fcst"]
        for k in range(len(offsets)):
            c = ws.cell(row=r, column=3 + k, value=formula(f"$B{r}", f"{get_column_letter(3 + k)}${hdr}"))
            c.number_format = st.FMT["price"]
            if o == 0 and k == 2:
                c.fill, c.font = st.PatternFill("solid", fgColor=st.THEME["pale"]), st.font(bold=True)


def _summary(ws, m):
    """Same labels as the standard Summary where the roll-up and write-up read them."""
    g, fin = m.grid, m.fin
    foreign = fin.currency != "USD"
    I = input_ref
    ws["A1"], ws["A1"].font = f"{fin.name} ({fin.ticker}) - Summary", st.TITLE_FONT
    ws["A2"] = ("Financials template: valued on equity (residual income and exit price-to-book), not free cash "
                "flow. Everything here is linked to the other tabs.")
    ws["A2"].font = st.NOTE_FONT

    def kv(r, label, f, fmt=None, col=1):
        ws.cell(row=r, column=col, value=label)
        c = ws.cell(row=r, column=col + 1, value=f)
        c.alignment = Alignment(horizontal="right")
        if fmt:
            c.number_format = fmt

    ws.cell(row=4, column=1, value="Snapshot").font = st.font(bold=True)
    kv(5, "Share price (US listing)", f'=IFERROR(IF({I("Price override")}<>"",{I("Price override")},{I("Live close")}),"")',
       st.FMT["usd_price"])
    kv(6, "Price date", f'=IF({I("Price override")}<>"","override",{I("Price date")})', st.FMT["date"])
    kv(7, "Market cap (millions)", f"={I('Market cap')}", st.FMT["usd"])
    kv(8, "Scenario", f"={I('Scenario')}")
    kv(5, "Valuation date", f"={REF('$B$' + str(VR['val_date']))}", st.FMT["date"], col=4)
    kv(6, "Cost of equity", f"={REF('$B$' + str(VR['ke']))}", st.FMT["pct"], col=4)
    kv(7, "Terminal growth", f"={REF('$B$' + str(VR['g']))}", st.FMT["pct"], col=4)
    kv(8, "Model checks", "n/a (equity model)", col=4)
    kv(10, "Input warnings", f'=IF({warnings_count()}=0,"none",{warnings_count()}&" - see the Check column on Inputs")', col=4)
    if foreign:
        kv(9, f"Share price ({fin.currency} per ordinary share)", f"={I('Share price')}", st.FMT["price"])

    top = 11
    ws.cell(row=top, column=1, value="Value per share by method").font = st.font(bold=True)
    heads = ["Method", "Low", "Value", "High", "Upside vs. price"] + (["$ per US-listed share"] if foreign else [])
    for j, h in enumerate(heads, start=1):
        st.header_cell(ws.cell(row=top + 1, column=j), h, "label" if j == 1 else "fcst")
    price = I("Share price")
    fx = f'IF({I("FX override")}<>"",{I("FX override")},{I("FX rate")})'
    to_usd, conv = f'*{I("Shares per listing")}/{fx}', f'*{fx}/{I("Shares per listing")}'
    g1, g2 = (REF(r) for r in GRID_CELLS)
    lo, hi = (lambda r: f'IF(COUNT({r})=0,"",MIN({r}))'), (lambda r: f'IF(COUNT({r})=0,"",MAX({r}))')
    rows = [("Equity model: residual income", lo(g1), REF(f"$B${VR['per_share']}"), hi(g1)),
            ("Equity model: exit P/B", lo(g2), REF(f"$C${VR['per_share']}"), hi(g2)),
            ("52-week trading range", f'{I("52w low")}{conv}', price, f'{I("52w high")}{conv}')]
    for i, (label, a, v, b) in enumerate(rows):
        r = top + 2 + i
        ws.cell(row=r, column=1, value=label)
        for j, f in ((2, a), (3, v), (4, b)):
            ws.cell(row=r, column=j, value=f'=IFERROR({f},"")').number_format = st.FMT["price"]
        ws.cell(row=r, column=3).font = st.font(bold=True)
        if i < 2:
            ws.cell(row=r, column=5, value=f'=IFERROR(C{r}/{price}-1,"")').number_format = st.FMT["pct"]
        if foreign:
            ws.cell(row=r, column=6, value=f'=IFERROR(C{r}{to_usd},"")').number_format = st.FMT["usd_price"]
    pr = top + 2 + len(rows)
    ws.cell(row=pr, column=1, value="Current share price").font = st.font(bold=True)
    for j in (2, 3, 4):
        ws.cell(row=pr, column=j, value=f'=IFERROR({price}*{(0.995, 1, 1.005)[j - 2]},"")').number_format = st.FMT["price"]

    kf = pr + 3
    ws.cell(row=kf, column=1, value="Key figures" + ("" if not foreign else f" ({fin.currency} millions)")).font = st.font(bold=True)
    cols = ([g.hist[-2]] if len(g.hist) > 1 else []) + [g.last_fy] + ([g.ltm] if g.ltm is not None else []) \
        + g.fcst[:2] + [g.fcst[-1]]
    st.header_cell(ws.cell(row=kf + 1, column=1), "")
    for j, col in enumerate(cols, start=2):
        st.header_cell(ws.cell(row=kf + 1, column=j), col.header, col.kind)
    figures = [("Revenue", "Revenue", "usd"), ("Net income", "Net income", "usd"), ("Diluted EPS", "Diluted EPS", "eps"),
               ("Return on equity", "Return on equity", "pct"), ("Shareholders' equity", "Shareholders' equity", "usd"),
               ("Book value per share", "Book value per share", "price"),
               ("Dividends % of net income", "Dividends % of net income", "pct"),
               ("Buybacks % of net income", "Buybacks % of net income", "pct")]
    for i, (text, label, fmt) in enumerate(figures):
        r = kf + 2 + i
        ws.cell(row=r, column=1, value=text)
        for j, col in enumerate(cols, start=2):
            ref = REF(f"{col.L}{ROW[label]}")
            ws.cell(row=r, column=j, value=f'=IF({ref}="","",{ref})').number_format = st.FMT[fmt]
    ws.column_dimensions["A"].width = 46
    for L in "BCDEFG":
        ws.column_dimensions[L].width = 18
    ws.freeze_panes = "A4"


def _mapping(ws, fin):
    tags_for = {label: tags for mp in fin.maps for label, _, _, tags in mp}
    ws.append(["Line on the Equity Model tab", "SEC line item(s)", "Tag priority (first with data wins)"])
    for c in ws[1]:
        c.fill, c.font = st.FILL["label"], st.HDR_FONT
    for label, lines, _ in DATA:
        ws.append([label, " + ".join(lines) or "formula",
                   "   |   ".join(" > ".join(fmt_tag(t) for t in tags_for.get(l, [])) for l in lines)])
    for col, w in zip("ABC", (34, 40, 150)):
        ws.column_dimensions[col].width = w


def coverage_report(m):
    """What the SEC data has for the lines this template uses (the standard statements don't apply)."""
    g = m.grid
    past = [c for c in g.cols if c.kind != "fcst"]
    ltm = f" + LTM to {m.fin.ltm_end.date()}" if m.fin.ltm_end is not None else ""
    print(f"\nCoverage across {len(g.hist)} fiscal years{ltm} (lines the financials template uses):")
    need = ("Net income", "Shareholders' equity", "Diluted shares")
    for label, lines, _ in DATA:
        if not lines:
            continue
        have = sum(any(m.has_data(x, c) for x in lines) for c in past)
        if have == len(past):
            flag = ""
        elif label in need:
            flag = "  <-- needed for the valuation: check the Mapping tab or add an overrides file"
        elif have == 0:
            flag = "  (not reported; shown blank)"
        else:
            flag = "  (partial)"
        print(f"  {label:<32}{have}/{len(past)}{flag}")

