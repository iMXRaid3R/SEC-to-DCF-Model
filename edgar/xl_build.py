"""Revenue Build tab: an optional segment build. Type each segment's history (at least the last fiscal year)
and its growth by year; switch it on and the income statement's revenue forecast comes from here instead of
the single growth driver. Amounts on this tab are in $ millions exactly as typed."""
from openpyxl.styles import Alignment
from openpyxl.worksheet.datavalidation import DataValidation

from . import xl_style as st
from .layout import (RB_OVR_FIRST, RB_GROW_FIRST, RB_GROW_TOTAL, RB_MODEL, RB_N, RB_SEG_FIRST, RB_SHEET, RB_SWITCH, RB_TOTAL,
                     RB_UNALLOC, Ctx)

MM = '#,##0.0;(#,##0.0);"-"'


def write_build(ws, m, write_header):
    g, fin = m.grid, m.fin
    write_header(ws, g, f"{fin.name} ({fin.ticker}) - {RB_SHEET}",
                 "Optional. Type segment names and their revenue in $ millions (yellow), set growth by year, then "
                 "switch the build On. Growth follows the Drivers revenue growth unless you type a rate in Your "
                 "growth overrides below. Yellow cells are inputs.",
                 first="($ millions, as typed)")
    sw = ws[RB_SWITCH.replace("$", "")]
    ws["E1"], ws["E1"].font = "Use build:", st.font(bold=True)
    sw.value, sw.font, sw.fill, sw.alignment = "Off", st.font(st.BLUE, True), st.INPUT_FILL, st.CENTER
    dv = DataValidation(type="list", formula1='"Off,On"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(sw)
    m.reg.add(ws, sw, "Build", "Switch")
    last = g.last_fy.L
    seg_last = f"${last}${RB_SEG_FIRST}:${last}${RB_TOTAL - 1}"
    ws["G1"] = (f'=IF({RB_SWITCH}<>"On","Off: revenue uses the growth driver",IF(SUM({seg_last})>0,'
                f'"On: this tab drives forecast revenue","On, but no {g.last_fy.header} segment revenue typed: '
                f'still using the growth driver"))')
    ws["G1"].font = st.NOTE_FONT

    for s in range(RB_N):
        r, gr = RB_SEG_FIRST + s, RB_GROW_FIRST + s
        name = ws.cell(row=r, column=1, value=f"Segment {s + 1}")
        name.font, name.fill = st.font(st.BLUE), st.INPUT_FILL
        m.reg.add(ws, name, "Build", f"Name {s + 1}")
        ws.cell(row=gr, column=1, value=f'=A{r}&" growth"').alignment = Alignment(indent=1)
        for col in g.cols:
            c, gc = ws.cell(row=r, column=col.j), ws.cell(row=gr, column=col.j)
            c.number_format, gc.number_format = MM, st.FMT["pct"]
            prev = g.prev(col)
            if col.kind == "hist":
                c.font, c.fill = st.font(st.BLUE), st.INPUT_FILL
                m.reg.add(ws, c, "Build", f"Segment {s + 1}", col.header)
                if prev is not None:
                    gc.value = f'=IF(OR(N({prev.L}{r})=0,N({col.L}{r})=0),"",{col.L}{r}/{prev.L}{r}-1)'
            elif col.kind == "fcst":
                c.fill = gc.fill = st.FCST_BG
                c.value = f"=N({prev.L}{r})*(1+{col.L}{gr})"
                ov = ws.cell(row=RB_OVR_FIRST + s, column=col.j)      # yours; blank = the Drivers growth
                ov.number_format = st.FMT["pct"]
                m.reg.add(ws, ov, "Build", f"Growth {s + 1}", col.header)
                gc.value = f'=IF({ov.coordinate}="",{Ctx(g, "RB", col).v("Revenue growth")},{ov.coordinate})'
                gc.font = st.font()

    def line(r, label, f, fmt=MM, bold=False, kinds=("hist", "ltm", "fcst")):
        ws.cell(row=r, column=1, value=label).font = st.font(bold=bold)
        for col in g.cols:
            if col.kind in kinds:
                c = ws.cell(row=r, column=col.j, value=f(col))
                c.number_format = fmt
                if isinstance(c.value, str):
                    c.font = st.formula_font(c.value, bold)
                if col.kind == "fcst":
                    c.fill = st.FCST_BG

    seg = lambda col: f"{col.L}{RB_SEG_FIRST}:{col.L}{RB_TOTAL - 1}"
    line(RB_TOTAL, "Total of segments", lambda col: f"=SUM({seg(col)})", bold=True, kinds=("hist", "fcst"))
    line(RB_MODEL, "Revenue in the model (reported / forecast)",
         lambda col: f'=IF({Ctx(g, "RB", col).v("Revenue")}="","",{Ctx(g, "RB", col).v("Revenue")}/1000000)')
    line(RB_UNALLOC, "Not allocated to a segment (reported - segments)",
         lambda col: f'=IF(OR(SUM({seg(col)})=0,{col.L}{RB_MODEL}=""),"",{col.L}{RB_MODEL}-{col.L}{RB_TOTAL})',
         kinds=("hist",))
    ws.cell(row=RB_GROW_FIRST - 1, column=1, value="Growth").font = st.font(bold=True)

    def total_growth(col):
        prev = g.prev(col)
        if prev is None:
            return None
        return f'=IF(OR(N({prev.L}{RB_TOTAL})=0,N({col.L}{RB_TOTAL})=0),"",{col.L}{RB_TOTAL}/{prev.L}{RB_TOTAL}-1)'
    line(RB_GROW_TOTAL, "Total of segments growth", total_growth, st.FMT["pct"], bold=True, kinds=("hist", "fcst"))
    t = RB_OVR_FIRST - 2
    ws.cell(row=t, column=1, value="Your growth overrides").font = st.font(bold=True)
    ws.cell(row=t, column=2, value="Blank = the revenue growth on the Drivers tab for that year; type a rate to use "
                                   "your own for a segment.").font = st.NOTE_FONT
    st.header_cell(ws.cell(row=t + 1, column=1), "Segment")
    for col in g.fcst:
        st.header_cell(ws.cell(row=t + 1, column=col.j), col.header, "ltm")
    for s_ in range(RB_N):
        ws.cell(row=RB_OVR_FIRST + s_, column=1, value=f'=A{RB_SEG_FIRST + s_}&" growth"')
    ws.column_dimensions["A"].width = 46
