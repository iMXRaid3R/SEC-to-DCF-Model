"""DCF tab: WACC, unlevered FCF from the forecast columns, two terminal values, sensitivity grids.

Valued as of today (or a typed valuation date): the part of forecast year 1 that has already passed is left
out, later cash flows are discounted from the valuation date, and net debt comes from the latest balance
sheet (the newest 10-Q when there is one)."""
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from . import xl_style as st
from .layout import INPUT_ROWS, Ctx, input_ref
from . import xl_inputs          # setting() / override_header(); module import: the two modules use each other

# Input defaults (blue cells; overwrite on the sheet)
DEFAULTS = {"rf": 0.0425, "erp": 0.055, "beta": 1.0, "kd": 0.05, "g": 0.025, "multiple": 15.0, "mid": 1}

# Row map. Other modules (Assumptions, tests) refer to DCF cells through R / ref() only.
_INPUTS = ["wacc", "price", "shares", "bs_date", "debt", "cash",
           "nci", "net_debt", "other_assets", "other_claims", "claims", "g", "multiple", "mid", "fy_end", "val_override", "val_date", "elapsed",
           "choices", "tv_mode", "roic_n", "roic_t", "sbc_mode", "capex_mode"]
_FCF = ["year", "ebit", "taxes", "nopat", "da", "capex", "wc", "sbc_add", "ufcf", "share", "counted", "period",
        "factor", "pv", "ebitda"]
# Method choices (dropdowns on the DCF tab); the first option of each is what the formulas test for.
TV_MODES = ("NOPAT less reinvestment", "Final-year FCF")
SBC_MODES = ("Cash cost", "Add back")
CAPEX_MODES = ("Model path", "Depreciation + growth")
from .layout import LEASE_MODES     # the lease choice is made on the Bridge tab
_VAL = ["sum_pv", "tv", "tv_period", "pv_tv", "ev", "less_debt", "less_nci", "plus_other", "less_other", "equity",
        "per_share", "cur_price",
        "upside", "tv_pct", "implied"]
RESULTS_HDR = 4                          # rows 4-7: the results first (value per share, vs. price)
INPUT_HDR = 9
R = {k: INPUT_HDR + 1 + i for i, k in enumerate(_INPUTS)}
FCF_HDR = R[_INPUTS[-1]] + 2
R.update({k: FCF_HDR + 1 + i for i, k in enumerate(_FCF)})
VAL_HDR = R[_FCF[-1]] + 2
R.update({k: VAL_HDR + 1 + i for i, k in enumerate(_VAL)})
GRID_GROWTH = R[_VAL[-1]] + 3            # title row of the two sensitivity grids, side by side:
GRID_MULTIPLE = GRID_GROWTH              # WACC x terminal growth in B:G, WACC x exit multiple in J:O
GRID_COLS = (2, 10)                      # the WACC column of each grid; its 5 value columns follow
SENS_CENTER = (f"E{GRID_GROWTH + 4}", f"M{GRID_MULTIPLE + 4}")
GRID_CELLS = (f"$C${GRID_GROWTH + 2}:$G${GRID_GROWTH + 6}", f"$K${GRID_MULTIPLE + 2}:$O${GRID_MULTIPLE + 6}")
# Reverse DCF: what the current share price implies
_REV = ["target_ev", "model_cagr", "imp_cagr", "imp_scale", "model_margin", "imp_margin", "imp_terminal"]
REV_HDR = GRID_GROWTH + 9
R.update({k: REV_HDR + 1 + i for i, k in enumerate(_REV)})
SOLVER_HDR = 3                           # on the hidden Calc sheet
SOLVER_STEPS = 14
CALC_SHEET = "Calc"


def ref(key, col="B"):
    """Absolute reference to a DCF cell from another sheet, e.g. ref("per_share", "C")."""
    return f"DCF!${col}${R[key]}"


def write_dcf(ws, m, calc):
    """`calc` is the hidden working sheet for the reverse-DCF solver steps."""
    g = m.grid
    n = len(g.fcst)
    ws["A1"], ws["A1"].font = f"{m.fin.name} ({m.fin.ticker}) - DCF Valuation", st.TITLE_FONT
    ws["A2"] = ("Settings: Default is what the model suggests; type in the yellow Your override column to replace "
                "it (clear it to go back); Used is what the model uses. The banner above shows the effect. Valued as of the valuation date: the elapsed part "
                "of forecast year 1 is excluded and net debt is from the latest balance sheet. Millions except "
                "per share.")
    ws["A2"].font = st.NOTE_FONT
    last = Ctx(g, "DCF", g.last_fy)
    yr1 = Ctx(g, "DCF", g.fcst[0])
    # Latest balance sheet: the 10-Q column when it has one, else the last fiscal year end.
    bs_col = g.ltm if g.ltm is not None and m.has_data("Total Assets", g.ltm) else g.last_fy
    bs = Ctx(g, "DCF", bs_col)
    B = lambda k: f"$B${R[k]}"

    def put(key, label, value, fmt, kind="formula", note=""):
        """kind "input": a setting in the default / override / used pattern (xl_inputs.setting)."""
        r = R[key]
        ws.cell(row=r, column=1, value=label)
        c = ws.cell(row=r, column=2, value=value)
        c.number_format = st.FMT[fmt]
        if kind == "input":
            return xl_inputs.setting(ws, m, r, key, label, value, fmt, note)
        if isinstance(value, str):
            c.font = st.formula_font(value)
        if note:
            st.hover_note(ws.cell(row=r, column=6), note).font = st.NOTE_FONT
        return c

    ws.cell(row=INPUT_HDR, column=1, value="Settings and the bridge to equity").font = st.font(bold=True)
    xl_inputs.override_header(ws, INPUT_HDR)
    from . import xl_wacc
    c = put("wacc", "WACC", f"={xl_wacc.ref('wacc')}", "pct",
            note="Worked out on the WACC tab (click the name): cost of equity, cost of debt and their weights")
    c.font = st.font(st.GREEN, True)
    st.link_to(ws.cell(row=R["wacc"], column=1), xl_wacc.SHEET, f"B{xl_wacc.W['wacc']}")
    from . import xl_bridge as br
    bridge = [("price", "Share price", "price", "price"), ("shares", "Diluted shares", "diluted", "shares"),
              ("bs_date", "Balance sheet date", "bs_date", "date"), ("debt", "Total debt", "debt", "usd"),
              ("cash", "Cash & investments", "cash", "usd"), ("nci", "Noncontrolling & redeemable interests", "nci", "usd"),
              ("net_debt", "Net debt", "net_debt", "usd"), ("other_assets", "Other non-operating assets", "other_assets", "usd"),
              ("other_claims", "Other claims on the business", "other_claims", "usd"),
              ("claims", "Net claims ahead of common equity", "claims", "usd")]
    for key, label, bkey, fmt in bridge:
        src = br.ref(bkey)
        c = put(key, label, f'=IF({src}="","",{src})' if key in ("other_assets", "other_claims") else f"={src}", fmt)
        c.font = st.font(st.GREEN, key in ("shares", "claims"))
    st.link_to(ws.cell(row=R["claims"], column=1), br.SHEET, f"B{br.K['claims']}")
    st.hover_note(ws.cell(row=R["claims"], column=6),
                  "From the Bridge tab (click the name): debt, cash, other owners' interests, other assets and claims, "
                  "and the diluted share count").font = st.NOTE_FONT
    put("g", "Terminal growth", DEFAULTS["g"], "pct", "input")
    if m.comps:
        put("multiple", "Exit EV/EBITDA multiple", f"=IFERROR(ROUND({m.comps['EV / EBITDA']},1),{DEFAULTS['multiple']:g})",
            "x", "input", note=f"Applied to {g.fcst[-1].header} EBITDA for the terminal value. Defaults to the peer "
                               f"median on the Comps tab ({DEFAULTS['multiple']:.0f}x when prices are unavailable). "
                               "Not the forward EV/EBITDA target on the Multiples tab")
    else:
        put("multiple", "Exit EV/EBITDA multiple", DEFAULTS["multiple"], "x", "input",
            note=f"Applied to {g.fcst[-1].header} EBITDA for the terminal value. Not the forward EV/EBITDA target "
                 "on the Multiples tab. Build with peers to default this to their median")
    put("mid", "Mid-year convention (1 = on)", DEFAULTS["mid"], "num", "input")
    for j in (2, 3, 4):
        ws.cell(row=R["mid"], column=j).number_format = "0"
    put("fy_end", "Last fiscal year end", g.last_fy.end.date(), "date")
    put("val_override", "Valuation date override", None, "date", "input", note="Optional: blank = today")
    put("val_date", "Valuation date", f'=IF({B("val_override")}="",TODAY(),{B("val_override")})', "date")
    put("elapsed", "Share of forecast year 1 already elapsed",
        f"=MAX(0,MIN(1,({B('val_date')}-{B('fy_end')})/365))", "pct",
        note="That part of year-1 free cash flow is excluded; later years are discounted from the valuation date")

    # ---- method choices
    ws.cell(row=R["choices"], column=1, value="Method choices").font = st.font(bold=True)
    final = Ctx(g, "DCF", g.fcst[-1])
    last_Lc = get_column_letter(2 + n)

    def choice(key, label, options, default, note):
        xl_inputs.setting(ws, m, R[key], key, label, default, "num", note, options)

    choice("tv_mode", "Terminal cash flow (perpetuity method)", TV_MODES, TV_MODES[0],
           f"{TV_MODES[0]}: NOPAT x (1 - growth / return on new capital), so a capex-heavy final year doesn't set "
           f"the terminal value. {TV_MODES[1]}: grows the last forecast year's free cash flow")
    ic = (f"({final.n('Total Equity')}+{final.n('Short-Term Borrowings')}+{final.n('Current Portion of LT Debt')}"
          f"+{final.n('Long-Term Debt')}-{final.n('Cash & Equivalents')}-{final.n('Short-Term Investments')}"
          f"-{final.n('Long-Term Investments')})")
    put("roic_n", f"Return on invested capital, {g.fcst[-1].header}",
        f'=IFERROR(IF({ic}<=0,"n/m",${last_Lc}${R["nopat"]}/{ic}),"n/m")', "pct",
        note="NOPAT / (equity + debt - cash and investments) in the final forecast year; n/m when capital is negative")
    put("roic_t", "Return on new capital in the terminal period",
        f"=IF(ISNUMBER({ref('roic_n')}),MAX({ref('wacc')},MIN(0.3,{ref('roic_n')})),0.3)", "pct", "input",
        note="Defaults to the final-year ROIC, held between WACC and 30%. Lower = more reinvestment needed to grow")
    choice("sbc_mode", "Stock-based compensation", SBC_MODES, SBC_MODES[0],
           f"{SBC_MODES[0]}: left in as an expense. {SBC_MODES[1]}: added to free cash flow (dilution then needs "
           "handling in the share count)")
    choice("capex_mode", "Capex in the free cash flow", CAPEX_MODES, CAPEX_MODES[0],
           f"{CAPEX_MODES[0]}: the forecast on the Cash Flow tab. {CAPEX_MODES[1]}: depreciation plus the revenue "
           "increase x last year's PP&E / revenue")

    # ---- Unlevered FCF, one column per forecast year (C..)
    dcol = {fc.j: 3 + k for k, fc in enumerate(g.fcst)}
    first_L, last_L = get_column_letter(3), get_column_letter(2 + n)
    st.header_cell(ws.cell(row=FCF_HDR, column=1), "Unlevered free cash flow")
    for fc in g.fcst:
        st.header_cell(ws.cell(row=FCF_HDR, column=dcol[fc.j]), fc.header, "fcst")
    row = lambda k, L: f"{L}{R[k]}"
    lines = [
        ("year", "Year", lambda c, L, k: k + 1, "int"),
        ("ebit", "EBIT", lambda c, L, k: f"={c.v('Operating Income')}", "usd"),
        ("taxes", "Taxes on EBIT", lambda c, L, k: f"=-{row('ebit', L)}*{c.v('Tax rate')}", "usd"),
        ("nopat", "NOPAT", lambda c, L, k: f"={row('ebit', L)}+{row('taxes', L)}", "usd"),
        ("da", "+ D&A", lambda c, L, k: f"={c.v('D&A')}", "usd"),
        ("capex", "- Capex",
         lambda c, L, k: (f'=IF({B("capex_mode")}="{CAPEX_MODES[0]}",-{c.v("Capex")},'
                          f'-(({c.v("D&A")}-{c.n("Amortization of Intangibles")})+MAX(0,{c.v("Revenue")}-{c.p("Revenue")})'
                          f'*IFERROR({last.last("PP&E, Net", n=True)}/{last.last("Revenue")},0)))'), "usd"),
        ("wc", "+/- Change in working capital", lambda c, L, k: f"={c.v('Working Capital & Other')}", "usd"),
        ("sbc_add", "+ Stock-based comp (when added back)",
         lambda c, L, k: f'=IF({B("sbc_mode")}="{SBC_MODES[1]}",{c.v("Stock-Based Comp")},0)', "usd"),
        ("ufcf", "Unlevered FCF (full year)", lambda c, L, k: f"=SUM({row('nopat', L)}:{row('sbc_add', L)})", "usd"),
        ("share", "Share of year after valuation date", lambda c, L, k: f"=1-{B('elapsed')}" if k == 0 else 1, "pct"),
        ("counted", "Unlevered FCF counted", lambda c, L, k: f"={row('ufcf', L)}*{row('share', L)}", "usd"),
        ("period", "Discount period (years)",
         lambda c, L, k: (f"=(1-{B('elapsed')})*(1-0.5*{B('mid')})" if k == 0
                          else f"={row('year', L)}-{B('elapsed')}-0.5*{B('mid')}"), "num"),
        ("factor", "Discount factor", lambda c, L, k: f"=1/(1+{B('wacc')})^{row('period', L)}", "num"),
        ("pv", "PV of unlevered FCF", lambda c, L, k: f"={row('counted', L)}*{row('factor', L)}", "usd"),
        ("ebitda", "EBITDA (memo)", lambda c, L, k: f"={c.v('EBITDA')}", "usd"),
    ]
    bold_rows = ("nopat", "ufcf", "counted")
    for key, label, f, fmt in lines:
        r = R[key]
        ws.cell(row=r, column=1, value=label).font = st.font(bold=key in bold_rows)
        for k, fc in enumerate(g.fcst):
            L = get_column_letter(dcol[fc.j])
            c = ws.cell(row=r, column=dcol[fc.j], value=f(Ctx(g, "DCF", fc), L, k))
            c.number_format = "0" if fmt == "int" else ("0.00" if fmt == "num" else st.FMT[fmt])
            if isinstance(c.value, str):
                c.font = st.formula_font(c.value, key in bold_rows)

    # ---- Valuation: B = perpetuity growth, C = exit multiple
    st.header_cell(ws.cell(row=VAL_HDR, column=1), "Valuation")
    st.header_cell(ws.cell(row=VAL_HDR, column=2), "Perpetuity growth", "fcst")
    st.header_cell(ws.cell(row=VAL_HDR, column=3), "Exit multiple", "fcst")
    N = f"${last_L}${R['year']}"
    fcf_n, ebitda_n, nopat_n = f"${last_L}${R['ufcf']}", f"${last_L}${R['ebitda']}", f"${last_L}${R['nopat']}"
    # Cash flow the perpetuity grows from, for a growth rate x: the final year's FCF, or NOPAT less the
    # reinvestment that growth needs (x / return on new capital).
    tcf = lambda x: f'IF({B("tv_mode")}="{TV_MODES[1]}",{fcf_n},{nopat_n}*(1-{x}/{B("roic_t")}))'
    pv_rng = f"${first_L}${R['pv']}:${last_L}${R['pv']}"
    V = lambda k, col: f"{col}{R[k]}"
    val = [
        ("sum_pv", "Sum of PV of FCF", lambda c: f"=SUM({pv_rng})", "usd"),
        ("tv", "Terminal value (end of final year)",
         lambda c: (f"={tcf(B('g'))}*(1+{B('g')})/({B('wacc')}-{B('g')})" if c == "B"
                    else f"={ebitda_n}*{B('multiple')}"), "usd"),
        ("tv_period", "Terminal value discount period (years)", lambda c: f"={N}-{B('elapsed')}", "num"),
        ("pv_tv", "PV of terminal value", lambda c: f"={V('tv', c)}/(1+{B('wacc')})^{V('tv_period', c)}", "usd"),
        ("ev", "Enterprise value", lambda c: f"={V('sum_pv', c)}+{V('pv_tv', c)}", "usd"),
        ("less_debt", "Less: net debt", lambda c: f"=-{B('net_debt')}", "usd"),
        ("less_nci", "Less: noncontrolling & redeemable interests", lambda c: f"=-{B('nci')}", "usd"),
        ("plus_other", "Plus: other non-operating assets", lambda c: f"=N({B('other_assets')})", "usd"),
        ("less_other", "Less: other claims", lambda c: f"=-N({B('other_claims')})", "usd"),
        ("equity", "Equity value", lambda c: f"=SUM({V('ev', c)}:{V('less_other', c)})", "usd"),
        ("per_share", "Value per share", lambda c: f'=IFERROR({V("equity", c)}/{B("shares")},"")', "price"),
        ("cur_price", "Current share price", lambda c: f'=IF({B("price")}=0,"",{B("price")})', "price"),
        ("upside", "Upside / (downside)", lambda c: f'=IFERROR({V("per_share", c)}/{V("cur_price", c)}-1,"")', "pct"),
        ("tv_pct", "Terminal value % of EV", lambda c: f'=IFERROR({V("pv_tv", c)}/{V("ev", c)},"")', "pct"),
        ("implied", "Implied exit multiple / implied growth",
         lambda c: (f'=IFERROR({V("tv", c)}/{ebitda_n},"")' if c == "B"
                    else f'=IFERROR({_growth_for_tv(V("tv", c), B, fcf_n, nopat_n)},"")'), None),
    ]
    bold_val = ("ev", "equity", "per_share")
    for key, label, f, fmt in val:
        r = R[key]
        ws.cell(row=r, column=1, value=label).font = st.font(bold=key in bold_val)
        for col in ("B", "C"):
            c = ws[f"{col}{r}"]
            c.value = f(col)
            c.number_format = ("0.00" if fmt == "num" else st.FMT[fmt]) if fmt else \
                (st.FMT["x"] if col == "B" else st.FMT["pct"])
            c.font = st.formula_font(c.value, key in bold_val)

    # ---- Sensitivities (each cell is a self-contained formula, so they work without Excel Data Tables)
    pv = (f"SUMPRODUCT(${first_L}${R['counted']}:${last_L}${R['counted']}"
          f"/(1+{{w}})^${first_L}${R['period']}:${last_L}${R['period']})")
    tvp = f"$B${R['tv_period']}"
    tail = f"-{B('claims')})/{B('shares')}"
    _grid(ws, GRID_GROWTH, "Perpetuity: WACC (rows) x terminal growth (columns)", B("g"), 0.005, "pct",
          lambda w, x: f'=IFERROR(({pv.format(w=w)}+{tcf(x)}*(1+{x})/({w}-{x})/(1+{w})^{tvp}{tail},"")', B("wacc"),
          GRID_COLS[0])
    _grid(ws, GRID_MULTIPLE, "Exit multiple: WACC (rows) x exit EV/EBITDA (columns)", B("multiple"), 1.0, "x",
          lambda w, x: f'=IFERROR(({pv.format(w=w)}+{ebitda_n}*{x}/(1+{w})^{tvp}{tail},"")', B("wacc"),
          GRID_COLS[1])
    st.section(ws, GRID_GROWTH - 1, 1, 15, "Value per share: sensitivity")

    _reverse(ws, m, B, V, first_L, last_L, N, nopat_n, calc)

    _results(ws)
    ws.column_dimensions["A"].width = 40
    for j in range(2, max(16, 3 + n)):
        ws.column_dimensions[get_column_letter(j)].width = 23 if j == 2 else 15 if j < 3 + n else 12
    ws.column_dimensions["I"].width = max(ws.column_dimensions["I"].width or 0, 15)
    ws.freeze_panes = "B4"


def _results(ws):
    """Rows 4-7: value per share and its upside under both terminal-value methods, before the workings."""
    st.section(ws, RESULTS_HDR, 1, 3, "Results")
    for j, h in ((2, "Perpetuity growth"), (3, "Exit multiple")):
        st.header_cell(ws.cell(row=RESULTS_HDR + 1, column=j), h, "fcst")
    for i, (label, key, fmt) in enumerate((("Value per share", "per_share", "price"), ("vs. share price", "upside", "pct"))):
        r = RESULTS_HDR + 2 + i
        ws.cell(row=r, column=1, value=label).font = st.font(bold=i == 0)
        for col in ("B", "C"):
            c = ws[f"{col}{r}"]
            c.value = f'=IF({col}{R[key]}="","",{col}{R[key]})'
            c.number_format, c.font = st.FMT[fmt], st.font(bold=i == 0)
            c.fill = st.PatternFill("solid", fgColor=st.THEME["pale"])
    st.row_rules(ws, [RESULTS_HDR + 2, RESULTS_HDR + 3], 3)


def _grid(ws, top, title, base_ref, step, fmt, formula, wacc_ref, wcol=2):
    """A 5 x 5 grid: WACC down column `wcol`, the other input across the next five columns."""
    WL = get_column_letter(wcol)
    ws.cell(row=top, column=max(1, wcol - 1), value=title).font = st.font(bold=True)
    hdr = top + 1
    st.header_cell(ws.cell(row=hdr, column=wcol), "WACC \\ ")
    offsets = (-2, -1, 0, 1, 2)
    for k, o in enumerate(offsets):
        c = ws.cell(row=hdr, column=wcol + 1 + k, value=f"={base_ref}+{o * step}")
        c.number_format, c.font, c.fill = st.FMT[fmt], st.HDR_FONT, st.FILL["fcst"]
    for i, o in enumerate(offsets):
        r = hdr + 1 + i
        w = ws.cell(row=r, column=wcol, value=f"={wacc_ref}+{o * 0.005}")
        w.number_format, w.font, w.fill = st.FMT["pct"], st.HDR_FONT, st.FILL["fcst"]
        for k in range(len(offsets)):
            x_ref = f"{get_column_letter(wcol + 1 + k)}${hdr}"
            c = ws.cell(row=r, column=wcol + 1 + k, value=formula(f"${WL}{r}", x_ref))
            c.number_format = st.FMT["price"]
            if o == 0 and k == 2:
                c.fill, c.font = PatternFill("solid", fgColor=st.THEME["pale"]), st.font(bold=True)   # today's inputs


def _growth_for_tv(X, B, fcf_n, nopat_n):
    """Expression for the perpetual growth rate at which the terminal value equals X.
    Final-year FCF basis: F(1+g)/(W-g) = X  ->  g = (X*W - F)/(X + F).
    NOPAT-less-reinvestment basis: N(1+g)(1-g/R)/(W-g) = X is a quadratic in g; the smaller root is the answer."""
    W, Rr, Nn = B("wacc"), B("roic_t"), nopat_n
    b = f"({Nn}*({Rr}-1)+{X}*{Rr})"
    cc = f"({X}*{Rr}*{W}-{Nn}*{Rr})"
    return (f'IF({B("tv_mode")}="{TV_MODES[1]}",({X}*{W}-{fcf_n})/({X}+{fcf_n}),'
            f"({b}-SQRT({b}^2-4*{Nn}*{cc}))/(2*{Nn}))")


def _reverse(ws, m, B, V, first_L, last_L, N, nopat_n, calc):
    """Reverse DCF: hold the model's margins and discount rate, and solve for what the share price implies."""
    g = m.grid
    rev0 = Ctx(g, "DCF", g.last_fy).ref("Revenue", g.last_fy, abs_col=True, abs_row=True)
    revn = Ctx(g, "DCF", g.fcst[-1]).ref("Revenue", g.fcst[-1], abs_col=True, abs_row=True)
    fcf_n, yrs, pvs = f"${last_L}${R['ufcf']}", f"${first_L}${R['year']}:${last_L}${R['year']}", \
        f"${first_L}${R['pv']}:${last_L}${R['pv']}"
    st.header_cell(ws.cell(row=REV_HDR, column=1), "What the share price implies (reverse DCF)")
    st.header_cell(ws.cell(row=REV_HDR, column=2), "Perpetuity growth", "fcst")
    st.header_cell(ws.cell(row=REV_HDR, column=3), "Exit multiple", "fcst")
    last_step = SOLVER_HDR + SOLVER_STEPS
    solved = {"B": ("B", "C"), "C": ("D", "E")}          # valuation column -> (growth, value) solver columns

    def implied_cagr(c):
        gcol, vcol = (f"{CALC_SHEET}!{x}" for x in solved[c])
        return (f'=IF({B("price")}=0,"",IFERROR(IF(ABS({vcol}{last_step}-{B("price")})<0.005*{B("price")},'
                f'{gcol}{last_step},"no solution"),"no solution"))')

    need = f"({B('price')}*{B('shares')}+{B('claims')})"

    def implied_g(c):
        X = f'(({need}-{V("sum_pv", c)})*(1+{B("wacc")})^{V("tv_period", c)})'       # terminal value the price needs
        return f'=IF({B("price")}=0,"",IFERROR({_growth_for_tv(X, B, fcf_n, nopat_n)},""))'

    lines = [
        ("target_ev", "Enterprise value at the current price", lambda c: f'=IF({B("price")}=0,"",{need})', "usd",
         "Diluted shares x price + net claims ahead of common equity"),
        ("model_cagr", f"Model: revenue growth per year, {g.last_fy.header} to {g.fcst[-1].header}",
         lambda c: f'=IFERROR(({revn}/{rev0})^(1/{N})-1,"")', "pct", "What the selected scenario assumes"),
        ("imp_cagr", "Implied: revenue growth per year the price needs", implied_cagr, "pct",
         "Solved with the model's margins, capital intensity and discount rate held as they are"),
        ("imp_scale", "Implied: free cash flow needed vs. the model (x)",
         lambda c: f'=IF({B("price")}=0,"",IFERROR({need}/{V("ev", c)},""))', "x",
         "1.0x = the price matches the model; 2.0x = it needs twice the model's cash flows"),
        ("model_margin", f"Model: {g.fcst[-1].header} unlevered FCF margin",
         lambda c: f'=IFERROR({fcf_n}/{revn},"")', "pct", ""),
        ("imp_margin", f"Implied: {g.fcst[-1].header} unlevered FCF margin at the model's revenue",
         lambda c: f'=IFERROR({V("imp_scale", c)}*{V("model_margin", c)},"")', "pct", ""),
        ("imp_terminal", "Implied: terminal growth / exit multiple",
         lambda c: (implied_g(c) if c == "B" else
                    f'=IF({B("price")}=0,"",IFERROR(({need}-{V("sum_pv", c)})*(1+{B("wacc")})^{V("tv_period", c)}'
                    f'/${last_L}${R["ebitda"]},""))'), None,
         "With the forecast cash flows as modelled: the terminal growth (left) or exit multiple (right) that "
         "gets to the current price"),
    ]
    for key, label, f, fmt, note in lines:
        r = R[key]
        ws.cell(row=r, column=1, value=label).font = st.font(bold=key.startswith("imp_"))
        for col in ("B", "C"):
            c = ws[f"{col}{r}"]
            c.value = f(col)
            c.number_format = st.FMT[fmt] if fmt else (st.FMT["pct"] if col == "B" else st.FMT["x"])
            c.font = st.formula_font(c.value, key.startswith("imp_"))
        if note:
            ws.cell(row=r, column=4, value=note).font = st.NOTE_FONT

    # Solver (on the hidden Calc sheet): secant iterations on a constant growth rate. Revenue in year t is the
    # model's revenue scaled by ((1+g)/(1+model growth))^t, so at the model's own growth rate the value equals
    # the DCF above.
    D = lambda cell: f"DCF!{cell}"
    calc["A1"], calc["A1"].font = "Working cells for the reverse DCF on the DCF tab. Nothing to edit here.", st.NOTE_FONT
    for j, h in enumerate(("Growth: perpetuity", "Value / share", "Growth: exit multiple", "Value / share"), start=2):
        c = calc.cell(row=SOLVER_HDR, column=j, value=h)
        c.font, c.alignment = st.NOTE_FONT, st.CENTER
    price, cagr = D(B("price")), D(B("model_cagr"))
    for val_col, (gcol, vcol) in solved.items():
        for k in range(SOLVER_STEPS):
            r = SOLVER_HDR + 1 + k
            if k == 0:
                gf = f"=N({cagr})"
            elif k == 1:
                gf = f"={gcol}{r - 1}+0.02"
            else:
                gf = (f"=IFERROR(MAX(-0.9,MIN(5,{gcol}{r - 1}-({vcol}{r - 1}-{price})*({gcol}{r - 1}-{gcol}{r - 2})"
                      f"/({vcol}{r - 1}-{vcol}{r - 2}))),{gcol}{r - 1})")
            ratio = f"((1+{gcol}{r})/(1+N({cagr})))"
            vf = (f'=IFERROR((SUMPRODUCT({D(pvs)}*{ratio}^{D(yrs)})+{D(f"${val_col}$" + str(R["pv_tv"]))}*{ratio}^{D(N)}'
                  f'-{D(B("claims"))})/{D(B("shares"))},0)')
            calc.cell(row=r, column=1, value=f"Step {k}").font = st.NOTE_FONT
            for addr, f, fmt in ((f"{gcol}{r}", gf, st.FMT["pct"]), (f"{vcol}{r}", vf, st.FMT["price"])):
                calc[addr].value, calc[addr].number_format, calc[addr].font = f, fmt, st.NOTE_FONT
    for L in "ABCDE":
        calc.column_dimensions[L].width = 22
