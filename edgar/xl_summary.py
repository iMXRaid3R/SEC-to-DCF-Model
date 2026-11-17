"""Summary tab: price, valuation by method with a football-field chart, what the price implies, key figures."""
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment

from . import xl_dcf
from . import xl_style as st
from .layout import SCENARIO_CELL, SHEET_NAMES, Ctx


def write_summary(ws, m):
    g, fin = m.grid, m.fin
    foreign = fin.currency != "USD"
    c0 = Ctx(g, "SUM", g.last_fy)
    mk = c0.market
    A = f"'{SHEET_NAMES['A']}'!" if " " in SHEET_NAMES["A"] else f"{SHEET_NAMES['A']}!"
    ws["A1"], ws["A1"].font = f"{fin.name} ({fin.ticker}) - Summary", st.TITLE_FONT
    ws["A2"] = ("Everything here is linked to the other tabs. " +
                ("$ per share." if not foreign else
                 f"Per-share values are in {fin.currency} per ordinary share; the last column converts to dollars "
                 "per US-listed share."))
    ws["A2"].font = st.NOTE_FONT

    def kv(r, label, f, fmt=None, col=1):
        ws.cell(row=r, column=col, value=label)
        c = ws.cell(row=r, column=col + 1, value=f)
        c.alignment = Alignment(horizontal="right")
        if fmt:
            c.number_format = fmt
        return c

    # ---- snapshot
    ws.cell(row=4, column=1, value="Snapshot").font = st.font(bold=True)
    kv(5, "Share price (US listing)", f'=IFERROR(IF({mk("Price override")}<>"",{mk("Price override")},'
                                      f'{mk("Live close")}),"")', st.FMT["usd_price"])
    kv(6, "Price date", f'=IF({mk("Price override")}<>"","override",{mk("Price date")})', st.FMT["date"])
    kv(7, "Market cap (millions)", f"={mk('Market cap')}", st.FMT["usd"])
    kv(8, "Scenario", f"={mk('Scenario')}")
    kv(5, "Valuation date", f"={xl_dcf.ref('val_date')}", st.FMT["date"], col=4)
    kv(6, "WACC", f"={xl_dcf.ref('wacc')}", st.FMT["pct"], col=4)
    kv(7, "Terminal growth", f"={xl_dcf.ref('g')}", st.FMT["pct"], col=4)
    kv(8, "Model checks", f"='{SHEET_NAMES['CHK']}'!$D$1", col=4)
    from .xl_inputs import warnings_count
    c = kv(10, "Input warnings", f'=IF({warnings_count()}=0,"none",{warnings_count()}&" - see the Check column on Inputs")',
           col=4)
    if foreign:
        kv(9, f"Share price ({fin.currency} per ordinary share)", f"={mk('Share price')}", st.FMT["price"])
        kv(9, f"{fin.currency} per US dollar", f'=IF({mk("FX override")}<>"",{mk("FX override")},{mk("FX rate")})',
           st.FMT["fx"], col=4)

    # ---- valuation by method (also the chart's data)
    top = 11
    ws.cell(row=top, column=1, value="Value per share by method").font = st.font(bold=True)
    heads = ["Method", "Low", "Value", "High", "Upside vs. price"] + (["$ per US-listed share"] if foreign else [])
    for j, h in enumerate(heads, start=1):
        st.header_cell(ws.cell(row=top + 1, column=j), h, "label" if j == 1 else "fcst")
    price = mk("Share price")
    to_usd = (f'*{mk("Shares per listing")}/IF({mk("FX override")}<>"",{mk("FX override")},{mk("FX rate")})')
    conv = f'*IF({mk("FX override")}<>"",{mk("FX override")},{mk("FX rate")})/{mk("Shares per listing")}'
    g1, g2 = (f"DCF!{rng}" for rng in xl_dcf.GRID_CELLS)
    lo_of = lambda rng: f'IF(COUNT({rng})=0,"",MIN({rng}))'
    hi_of = lambda rng: f'IF(COUNT({rng})=0,"",MAX({rng}))'
    rows = [("DCF: perpetuity growth", lo_of(g1), xl_dcf.ref("per_share", "B"), hi_of(g1),
             "Range: the WACC x terminal growth grid on the DCF tab"),
            ("DCF: exit multiple", lo_of(g2), xl_dcf.ref("per_share", "C"), hi_of(g2),
             "Range: the WACC x exit multiple grid on the DCF tab")]
    if m.comps and "_implied" in m.comps:
        imp = m.comps["_implied"]
        rows.append(("Peer multiples", lo_of(imp), f"MEDIAN({imp})", hi_of(imp),
                     "EV/EBITDA, EV/Revenue and P/E at the peer medians (Comps tab)"))
    rows.append(("52-week trading range", f'{mk("52w low")}{conv}', price, f'{mk("52w high")}{conv}', ""))
    first = top + 2
    for i, (label, lo, val, hi, note) in enumerate(rows):
        r = first + i
        ws.cell(row=r, column=1, value=label)
        for j, f in ((2, lo), (3, val), (4, hi)):
            c = ws.cell(row=r, column=j, value=f'=IFERROR({f},"")')
            c.number_format = st.FMT["price"]
        ws.cell(row=r, column=3).font = st.font(bold=True)
        if label != "52-week trading range":
            c = ws.cell(row=r, column=5, value=f'=IFERROR(C{r}/{price}-1,"")')
            c.number_format = st.FMT["pct"]
        if foreign:
            c = ws.cell(row=r, column=6, value=f'=IFERROR(C{r}{to_usd},"")')
            c.number_format = st.FMT["usd_price"]
        if note:
            ws.cell(row=r, column=len(heads) + 1, value=note).font = st.NOTE_FONT
    last = first + len(rows) - 1
    pr = last + 1
    ws.cell(row=pr, column=1, value="Current share price").font = st.font(bold=True)
    for j in (2, 3, 4):
        c = ws.cell(row=pr, column=j, value=f'=IFERROR({price}*{(0.995, 1, 1.005)[j - 2]},"")')
        c.number_format = st.FMT["price"]

    # chart helper columns (numbers only): start of each bar and its length
    hc = 12
    ws.cell(row=top + 1, column=hc, value="chart: start").font = st.NOTE_FONT
    ws.cell(row=top + 1, column=hc + 1, value="chart: length").font = st.NOTE_FONT
    for r in range(first, pr + 1):
        ws.cell(row=r, column=hc, value=f"=N(B{r})").font = st.NOTE_FONT
        ws.cell(row=r, column=hc + 1, value=f"=MAX(0,N(D{r})-N(B{r}))").font = st.NOTE_FONT
    # ---- reverse DCF
    rv = pr + 3
    ws.cell(row=rv, column=1, value="What the share price implies").font = st.font(bold=True)
    for j, h in enumerate(("", "DCF: perpetuity growth", "DCF: exit multiple"), start=1):
        st.header_cell(ws.cell(row=rv + 1, column=j), h, "label" if j == 1 else "fcst")
    lines = [("Revenue growth per year in the model", "model_cagr", "pct", "pct"),
             ("Revenue growth per year the price needs", "imp_cagr", "pct", "pct"),
             ("Free cash flow needed vs. the model", "imp_scale", "x", "x"),
             ("Terminal growth / exit multiple the price needs", "imp_terminal", "pct", "x")]
    for i, (label, key, fb, fc) in enumerate(lines):
        r = rv + 2 + i
        ws.cell(row=r, column=1, value=label)
        for j, (col, fmt) in enumerate((("B", fb), ("C", fc)), start=2):
            c = ws.cell(row=r, column=j, value=f"={xl_dcf.ref(key, col)}")
            c.number_format, c.alignment = st.FMT[fmt], Alignment(horizontal="right")

    # ---- key figures
    kf = rv + 2 + len(lines) + 2
    ws.cell(row=kf, column=1, value="Key figures" + ("" if not foreign else f" ({fin.currency} millions)")
            ).font = st.font(bold=True)
    cols = [g.hist[-2]] if len(g.hist) > 1 else []
    cols += [g.last_fy] + ([g.ltm] if g.ltm is not None else []) + g.fcst[:2] + ([g.fcst[-1]] if len(g.fcst) > 2 else [])
    st.header_cell(ws.cell(row=kf + 1, column=1), "")
    for j, col in enumerate(cols, start=2):
        st.header_cell(ws.cell(row=kf + 1, column=j), col.header, col.kind)
    figures = [("Revenue", "Revenue", "usd"), ("Revenue growth", "Revenue growth", "pct"),
               ("Gross margin", "Gross margin", "pct"), ("Operating margin", "Operating margin", "pct"),
               ("Net income", "Net Income", "usd"), ("Diluted EPS", "Diluted EPS", "eps"),
               ("Free cash flow", "Free cash flow", "usd"), ("FCF margin", "FCF margin", "pct")]
    for i, (text, label, fmt) in enumerate(figures):
        r = kf + 2 + i
        ws.cell(row=r, column=1, value=text)
        for j, col in enumerate(cols, start=2):
            ref = Ctx(g, "SUM", col).v(label)
            c = ws.cell(row=r, column=j, value=f'=IF({ref}="","",{ref})')
            c.number_format = st.FMT[fmt]

    chart = BarChart()
    chart.type, chart.grouping, chart.overlap = "bar", "stacked", 100
    chart.add_data(Reference(ws, min_col=hc, max_col=hc + 1, min_row=top + 1, max_row=pr), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=first, max_row=pr))
    chart.series[0].graphicalProperties.noFill = True
    chart.series[0].graphicalProperties.line.noFill = True
    chart.series[1].graphicalProperties.solidFill = "1E5C4F"
    chart.legend = None
    chart.y_axis.number_format = "#,##0"
    chart.y_axis.delete = False
    chart.x_axis.delete = False
    chart.x_axis.scaling.orientation = "maxMin"          # first method at the top
    chart.y_axis.crosses = "max"                         # ... and the value axis along the bottom
    chart.y_axis.majorGridlines = None
    chart.gapWidth = 60
    chart.height, chart.width = 7, 24
    ws.cell(row=kf + len(figures) + 3, column=1, value="Value per share by method: low to high" + (
        "" if not foreign else f" ({fin.currency} per ordinary share)")).font = st.font(bold=True)
    ws.add_chart(chart, f"A{kf + len(figures) + 4}")



    ws.column_dimensions["A"].width = 46
    for L in "BCDEFG":
        ws.column_dimensions[L].width = 18
    ws.freeze_panes = "A4"
