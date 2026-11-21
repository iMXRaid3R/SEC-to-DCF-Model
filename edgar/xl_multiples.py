"""Multiples tab: value per share from forward multiples of the model's own forecast.

Each method applies a target multiple to a next-twelve-months figure (a blend of forecast years 1 and 2 by how
much of year 1 has passed):
    Forward P/E               EPS x P/E
    PEG                       EPS x (PEG x EPS growth in %)  - the P/E that growth justifies
    Forward EV/EBITDA         (EBITDA x multiple - net claims) / diluted shares
    Forward EV/Sales          (revenue x multiple - net claims) / diluted shares
    Price / free cash flow    free cash flow per share x multiple
Targets default to the peer medians when there is a Comps tab; type your own in the override column (the
default stays beside it). Low / High apply the range (+/- 20% by default). Every method can be switched off. The current multiples at today's price sit next to each, so you can see what the market pays."""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment
from openpyxl.worksheet.datavalidation import DataValidation

from . import xl_dcf
from . import xl_style as st
from .layout import SHEET_NAMES, Ctx

SHEET = "Multiples"
DEFAULTS = {"pe": 20.0, "peg": 1.0, "ev_ebitda": 15.0, "p_fcf": 25.0}
SPREAD = 0.2
# key, label on this tab and on Summary, metric row, metric format, multiple format
METHODS = [("pe", "Forward P/E", "eps", "eps", "x"),
           ("peg", "PEG", "eps_growth", "pct", "num"),
           ("ev_ebitda", "Forward EV/EBITDA", "ebitda", "usd", "x"),
           ("ev_sales", "Forward EV/Sales", "revenue", "usd", "x"),
           ("p_fcf", "Price / free cash flow", "fcf_ps", "eps", "x")]
LABELS = {k: label for k, label, *_ in METHODS}

# rows
R = {"val_date": 5, "fy_end": 6, "elapsed": 7, "price": 8, "shares": 9, "claims": 10, "peg_years": 11,
     "eps_growth": 12}
FWD_HDR = 14
FWD = ["revenue", "ebitda", "net_income", "eps", "fcf", "fcf_ps"]
R.update({k: FWD_HDR + 1 + i for i, k in enumerate(FWD)})
MET_HDR = R[FWD[-1]] + 3
R.update({k: MET_HDR + 1 + i for i, (k, *_) in enumerate(METHODS)})
R["average"] = MET_HDR + 1 + len(METHODS)
COL = {"metric": "B", "now": "C", "now_ltm": "D", "peer": "E", "default": "F", "override": "G", "target": "H",
       "range": "I", "v_low": "J", "value": "K", "v_high": "L", "upside": "M", "use": "N", "how": "O"}


def ref(key, col="K"):
    """Absolute reference to a Multiples cell from another sheet, e.g. ref("pe") for the P/E value."""
    return f"{SHEET}!${col}${R[key]}"


def write_multiples(ws, m):
    g, fin = m.grid, m.fin
    has_dcf = "dcf" in m.parts
    ws["A1"], ws["A1"].font = f"{fin.name} ({fin.ticker}) - Valuation by multiples", st.TITLE_FONT
    ws["A2"] = ("Target multiples applied to next-twelve-months figures from the forecast (the selected scenario). "
                + st.INPUT_LEGEND + " Here: targets, the Low / High range, and On / Off for each method. "
                + ("Amounts in millions except per share." if fin.currency == "USD" else
                   f"Amounts in {fin.currency}; per share = per ordinary share."))
    ws["A2"].font = st.NOTE_FONT
    last = Ctx(g, "MUL", g.last_fy)
    y1, y2 = Ctx(g, "MUL", g.fcst[0]), Ctx(g, "MUL", g.fcst[1])
    B = lambda k: f"$B${R[k]}"

    def put(key, label, value, fmt, note="", reg=None):
        ws.cell(row=R[key], column=1, value=label)
        c = ws.cell(row=R[key], column=2, value=value)
        c.number_format, c.alignment = st.FMT[fmt], Alignment(horizontal="right")
        if reg:
            c.font = st.font(st.BLUE)
            m.reg.add(ws, c, "Multiples", reg)
        elif isinstance(value, str):
            c.font = st.formula_font(value)
        if note:
            ws.cell(row=R[key], column=3, value=note).font = st.NOTE_FONT
        return c

    ws.cell(row=4, column=1, value="Inputs").font = st.font(bold=True)
    if has_dcf:
        put("val_date", "Valuation date", f"={xl_dcf.ref('val_date')}", "date", "From the DCF tab (Inputs: valuation date override)")
    else:
        put("val_date", "Valuation date", "=TODAY()", "date")
    put("fy_end", "Last fiscal year end", g.last_fy.end.date(), "date")
    put("elapsed", f"Share of {g.fcst[0].header} already elapsed",
        f"=MAX(0,MIN(1,({B('val_date')}-{B('fy_end')})/365))", "pct",
        f"Next twelve months = {g.fcst[0].header} x (1 - this) + {g.fcst[1].header} x this")
    put("price", "Share price" + ("" if fin.currency == "USD" else f" ({fin.currency} per ordinary share)"),
        f"=N({last.market('Share price')})", "price", "Inputs tab")
    from . import xl_bridge as br
    put("shares", "Diluted shares", f"={br.ref('diluted')}", "shares", "Bridge tab")
    put("claims", "Net claims ahead of common equity", f"={br.ref('claims')}", "usd",
        "Bridge tab: net debt + other owners' interests + other claims - other assets")
    span = min(5, len(g.fcst) - 1)
    put("peg_years", "EPS growth window for PEG (years)", span, "num",
        f"EPS growth = yearly growth from {g.fcst[0].header} over this many years (1 to {len(g.fcst) - 1})",
        reg="PEG growth years").number_format = "0"
    eps_row = g.rows["Diluted EPS"][1]
    is_ = f"'{SHEET_NAMES['IS']}'!"
    f1, fl = g.fcst[0].L, g.fcst[-1].L
    put("eps_growth", "EPS growth per year (Base of PEG)",
        f'=IFERROR(IF(AND({is_}{f1}{eps_row}>0,INDEX({is_}{f1}{eps_row}:{fl}{eps_row},1,1+{B("peg_years")})>0),'
        f'(INDEX({is_}{f1}{eps_row}:{fl}{eps_row},1,1+{B("peg_years")})/{is_}{f1}{eps_row})^(1/{B("peg_years")})-1,""),"")',
        "pct", "Blank when EPS is not positive at both ends")

    # ---- forward figures
    tcol = g.ltm if g.ltm is not None and m.has_data("Revenue", g.ltm) else g.last_fy     # reported, trailing
    tc = Ctx(g, "MUL", tcol)
    for j, h in enumerate(("Forward and reported figures", "Next 12 months", g.fcst[0].header, g.fcst[1].header,
                           f"Reported: {tcol.header}"), start=1):
        st.header_cell(ws.cell(row=FWD_HDR, column=j), h, ("fcst", "ltm")[j == 5] if j > 1 else "label")
    fwd = [("revenue", "Revenue", "Revenue", "usd"), ("ebitda", "EBITDA", "EBITDA", "usd"),
           ("net_income", "Net income", "Net Income", "usd"), ("eps", "Diluted EPS", "Diluted EPS", "eps"),
           ("fcf", "Free cash flow", "Free cash flow", "usd")]
    for key, text, label, fmt in fwd:
        r = R[key]
        ws.cell(row=r, column=1, value=text)
        for j, c in ((3, y1), (4, y2)):
            cell = ws.cell(row=r, column=j, value=f"={c.n(label)}")
            cell.number_format, cell.font = st.FMT[fmt], st.font(st.GREEN)
        cell = ws.cell(row=r, column=2, value=f"=C{r}*(1-{B('elapsed')})+D{r}*{B('elapsed')}")
        cell.number_format, cell.font = st.FMT[fmt], st.font(bold=True)
        cell = ws.cell(row=r, column=5, value=f'=IF({tc.v(label)}="","",{tc.v(label)})')
        cell.number_format, cell.font = st.FMT[fmt], st.font(st.GREEN)
    r = R["fcf_ps"]
    ws.cell(row=r, column=1, value="Free cash flow per share")
    for L in ("B", "E"):
        cell = ws.cell(row=r, column=ord(L) - 64, value=f'=IFERROR({L}{R["fcf"]}/{B("shares")},"")')
        cell.number_format, cell.font = st.FMT["eps"], st.font(bold=L == "B")

    # ---- methods. Columns are named in COL; L(name) gives the letter, so the formulas read as words.
    L = COL.__getitem__
    heads = {"metric": "Metric (next 12 months)", "now": "Today's price on your forecast (NTM)",
             "now_ltm": "Today's price on reported results (LTM)", "peer": "Peer median (LTM, Comps tab)",
             "default": "Default target", "override": "Your target", "target": "Target used", "range": "Range +/-",
             "v_low": "Value: low", "value": "Value per share", "v_high": "Value: high", "upside": "vs. price",
             "use": "Use", "how": "How it works"}
    st.header_cell(ws.cell(row=MET_HDR, column=1), "Method")
    for name, h in heads.items():
        kind = "label" if name == "how" else "ltm" if name in ("override", "range", "use") else "fcst"
        c = ws.cell(row=MET_HDR, column=ord(L(name)) - 64)
        st.header_cell(c, h, kind)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[MET_HDR].height = 45
    comps = m.comps or {}
    peer_h = {"pe": "P / E", "ev_ebitda": "EV / EBITDA", "ev_sales": "EV / Revenue"}   # what Comps measures

    def peer(h, fb):
        """The peer median for this multiple (falling back to fb when it has no value), or fb without peers."""
        if h not in comps:
            return fb
        return "=IFERROR(ROUND(" + comps[h] + ",1)," + ('""' if fb is None else str(fb)) + ")"
    ev = f"({B('price')}*{B('shares')}+{B('claims')})"
    # the metric each method uses: next twelve months (column B of the figures above) or reported (column E)
    fig = {"pe": "eps", "ev_ebitda": "ebitda", "ev_sales": "revenue", "p_fcf": "fcf_ps"}
    metric = {k: f"=B{R[v]}" for k, v in fig.items()}
    metric["peg"] = f"={B('eps_growth')}"

    def today(key, col, r):
        """Price over a figure from column `col` of the figures block (B forecast, E reported)."""
        if key == "peg":
            return f"{B('price')}/{col}{R['eps']}/(B{r}*100)", f"AND(B{r}>0,{col}{R['eps']}>0)"
        x = f"{col}{R[fig[key]]}"
        return (f"{ev}/{x}" if key in ("ev_ebitda", "ev_sales") else f"{B('price')}/{x}"), f"N({x})>0"
    value = {"pe": "B{r}*{m}", "peg": f"B{R['eps']}*{{m}}*B{{r}}*100",
             "ev_ebitda": f"(B{{r}}*{{m}}-{B('claims')})/{B('shares')}",
             "ev_sales": f"(B{{r}}*{{m}}-{B('claims')})/{B('shares')}", "p_fcf": "B{r}*{m}"}
    positive = {"peg": f"AND(B{{r}}>0,B{R['eps']}>0)"}
    ebitda_target = peer("EV / EBITDA", DEFAULTS["ev_ebitda"])    # separate from the DCF's exit multiple
    targets = {"pe": peer("P / E", DEFAULTS["pe"]), "peg": DEFAULTS["peg"], "ev_ebitda": ebitda_target,
               "ev_sales": peer("EV / Revenue", None) if comps else None,
               "p_fcf": (f"=IFERROR(ROUND(1/({xl_dcf.ref('wacc')}-{xl_dcf.ref('g')}),1),{DEFAULTS['p_fcf']:g})"
                         if has_dcf else DEFAULTS["p_fcf"])}
    peer_note = (" Default: the peer median, which is a trailing multiple: compare it with today's price on "
                 "reported results, not on your forecast. As a target for forward figures it leans generous."
                 if comps else "")
    how = {"pe": "EPS x P/E." + (peer_note if "P / E" in comps else
                                 f" Default {DEFAULTS['pe']:g}x, a market-level starting point: type your own."),
           "peg": "EPS x (PEG x EPS growth in %): a PEG of 1 means a P/E equal to the growth rate. Blank when "
                  "growth or EPS isn't positive. Both of today's PEGs use the forecast EPS growth.",
           "ev_ebitda": "(Next-12-months EBITDA x multiple - net claims) / diluted shares." + (
               peer_note if "EV / EBITDA" in comps else f" Default {DEFAULTS['ev_ebitda']:g}x.") +
               (" Separate from the DCF's exit multiple, which applies to final-year EBITDA." if has_dcf else ""),
           "ev_sales": "(Revenue x multiple - net claims) / diluted shares." + (
               peer_note if comps else " No default without peers: type a target and switch it On."),
           "p_fcf": "Free cash flow per share x multiple." + (
               " Default: 1 / (WACC - terminal growth), a one-stage DCF on next year's cash flow." if has_dcf else
               f" Default {DEFAULTS['p_fcf']:g}x (a 4% free-cash-flow yield).")}
    dv = DataValidation(type="list", formula1='"On,Off"', allow_blank=False)
    ws.add_data_validation(dv)
    for key, label, _, mfmt, xfmt in METHODS:
        r = R[key]
        at = lambda name: ws.cell(row=r, column=ord(L(name)) - 64)
        ws.cell(row=r, column=1, value=label).font = st.font(bold=True)
        c = at("metric")
        c.value, c.number_format = metric[key], st.FMT[mfmt]
        c.font = st.font(st.GREEN if key != "peg" else st.BLACK)
        pos = positive.get(key, "B{r}>0").format(r=r)
        for name, col in (("now", "B"), ("now_ltm", "E")):
            f, ok = today(key, col, r)
            c = at(name)
            c.value = f'=IFERROR(IF(AND({B("price")}>0,{ok}),{f},""),"")'
            c.number_format = st.FMT[xfmt]
        c = at("peer")
        c.number_format = st.FMT[xfmt]
        if key in peer_h and peer_h[key] in comps:
            c.value, c.font = f'=IF({comps[peer_h[key]]}="","",{comps[peer_h[key]]})', st.font(st.GREEN)
        t = targets[key]
        c = at("default")
        c.value, c.number_format = t, st.FMT[xfmt]
        c.font = st.formula_font(t) if isinstance(t, str) else st.font()
        c = at("override")                                  # yours: blank = use the default
        c.number_format = st.FMT[xfmt]
        m.reg.add(ws, c, "Multiples", label, "Target")
        D, O, T, RG, U = (f"{L(n)}{r}" for n in ("default", "override", "target", "range", "use"))
        c = at("target")
        c.value = f'=IF({O}="",IF({D}="","",{D}),{O})'
        c.number_format, c.font = st.FMT[xfmt], st.font(bold=True)
        c = at("range")
        c.value, c.number_format = SPREAD, st.FMT["pct"]
        m.reg.add(ws, c, "Multiples", label, "Range")
        use = at("use")
        use.value, use.alignment = ("On" if t is not None else "Off"), st.CENTER
        dv.add(use)
        m.reg.add(ws, use, "Multiples", label, "Use")
        for col, mult in (("v_low", f"{T}*(1-N({RG}))"), ("value", T), ("v_high", f"{T}*(1+N({RG}))")):
            f = value[key].format(r=r, m=f"({mult})")
            c = at(col)
            c.value = f'=IFERROR(IF(AND({U}="On",N({T})>0,{pos}),{f},""),"")'
            c.number_format, c.font = st.FMT["price"], st.font(bold=col == "value")
        V = f"{L('value')}{r}"
        c = at("upside")
        c.value = f'=IFERROR(IF(OR({V}="",{B("price")}=0),"",{V}/{B("price")}-1),"")'
        c.number_format = st.FMT["pct"]
        st.hover_note(at("how"), how[key], keep=70).font = st.NOTE_FONT
    first, lastr = R[METHODS[0][0]], R[METHODS[-1][0]]
    r = R["average"]
    ws.cell(row=r, column=1, value="Average of the methods switched on").font = st.font(bold=True)
    for name in ("v_low", "value", "v_high"):
        col = L(name)
        c = ws.cell(row=r, column=ord(col) - 64, value=f'=IFERROR(AVERAGE({col}{first}:{col}{lastr}),"")')
        c.number_format, c.font = st.FMT["price"], st.font(bold=True)
    V = f"{L('value')}{r}"
    c = ws.cell(row=r, column=ord(L("upside")) - 64, value=f'=IFERROR(IF(OR({V}="",{B("price")}=0),"",{V}/{B("price")}-1),"")')
    c.number_format = st.FMT["pct"]
    ws.conditional_formatting.add(f"A{first}:{L('upside')}{lastr}",
                                  FormulaRule(formula=[f'${L("use")}{first}="Off"'], font=st.font(st.GREY, italic=True)))
    widths = {"A": 40, "B": 16, "C": 14, "D": 14, "E": 13, "F": 11, "G": 11, "H": 11, "I": 9, "J": 12, "K": 14,
              "L": 12, "M": 10, "N": 7, "O": 110}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "B4"
