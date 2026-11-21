"""Summary tab: a one-screen dashboard.

    title bar
    headline cards        share price | price target | upside | 12-month target | scenario | model checks
    three panels          snapshot | valuation inputs | free-cash-flow yield
    value per share       each method with its range and your weight, the weighted price target, and the
                          football-field chart beside a revenue-and-margin chart
    market vs. your model what the price implies (reverse DCF) | today's multiples on forecast, reported, peers
    key figures

The roll-up and the write-up find what they need by the labels in column A (and the panels' labels), so keep
those labels when changing the layout. Chart data sits in hidden helper columns on the right."""
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import xl_dcf, xl_multiples, xl_wacc
from . import xl_style as st
from .layout import INPUT_ROWS, SHEET_NAMES, Ctx

# Default weight of each method in the price target. Methods without a value (part not built, switched off,
# negative earnings) are left out and the remaining weights scaled back to 100%.
# Price / free cash flow defaults to 0%: at its default multiple, 1 / (WACC - growth), it is a one-stage DCF and
# would count the DCF twice. FCF yield is shown in its own panel instead.
WEIGHTS = {"DCF: perpetuity growth": 0.30, "DCF: exit multiple": 0.25, "Forward P/E": 0.20, "PEG": 0.05,
           "Forward EV/EBITDA": 0.20, "Forward EV/Sales": 0.0, "Price / free cash flow": 0.0, "Peer multiples": 0.0}
TARGET, TARGET_12M, ROLL = "Price target (weighted)", "12-month price target", "Return to roll it forward a year"

# Grid: A labels | B..F numbers | G weight | H gutter | I..L right-hand panels
WEIGHT_COL = 7                                  # G
YIELD_COL = 9                                   # I: the free-cash-flow yield panel (label; value in J)
WIDTHS = {"A": 36, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13, "G": 10, "H": 2, "I": 40, "J": 13, "K": 13, "L": 13}
LAST_COL = "L"
CARD_ROW = 4                                    # rows 4-6: label, value, sub-line
PANEL_ROW = 8                                   # section headers of the three panels; their rows follow
TABLE_TOP = 16
FF_COL, RV_COL = 30, 26                         # hidden helper columns: football field (AD:AE), revenue chart (Z:AB)

NAVY, INK, MUTED, CARD_BG, RULE, GOLD_BG = "1F3864", "1F2937", "6B7280", "F3F6FA", "D0D7E2", "FFF4D6"
THIN, MEDIUM = Side(style="thin", color=RULE), Side(style="medium", color=NAVY)
UP, DOWN = Font(color="1E7B34", bold=True), Font(color="B42318", bold=True)


# ---- small layout helpers ---------------------------------------------------------------------------------
def _col(letter):
    return ord(letter) - 64


def _section(ws, r, c1, c2, text, note=None):
    """A section title with a navy rule under it across columns c1..c2."""
    ws.cell(row=r, column=_col(c1), value=text).font = Font(bold=True, size=12, color=NAVY)
    for j in range(_col(c1), _col(c2) + 1):
        ws.cell(row=r, column=j).border = Border(bottom=MEDIUM)
    if note:
        ws.cell(row=r + 1, column=_col(c1), value=note).font = Font(italic=True, size=9, color=MUTED)


def _kv(ws, r, label, f, fmt=None, col="A", value_col=None, bold=False):
    """Label / value on one row of a panel, with a hairline under it."""
    vc = value_col or get_column_letter(_col(col) + 1)
    ws.cell(row=r, column=_col(col), value=label).font = Font(color=INK, bold=bold)
    c = ws.cell(row=r, column=_col(vc), value=f)
    c.alignment = Alignment(horizontal="right")
    c.font = Font(color=INK, bold=bold)
    if fmt:
        c.number_format = fmt
    for j in range(_col(col), _col(vc) + 1):
        ws.cell(row=r, column=j).border = Border(bottom=Side(style="hair", color=RULE))
    return c


def _card(ws, c1, c2, label, value, fmt=None, sub=None):
    """A headline number: small caps label, big value, a grey line under it, on a pale panel."""
    a, b = _col(c1), _col(c2)
    fill = PatternFill("solid", fgColor=CARD_BG)
    for r in range(CARD_ROW, CARD_ROW + 3):
        for j in range(a, b + 1):
            c = ws.cell(row=r, column=j)
            c.fill = fill
            c.border = Border(left=THIN if j == a else None, right=THIN if j == b else None,
                              top=THIN if r == CARD_ROW else None, bottom=THIN if r == CARD_ROW + 2 else None)
        if b > a:
            ws.merge_cells(start_row=r, start_column=a, end_row=r, end_column=b)
    top = ws.cell(row=CARD_ROW, column=a, value=label.upper())
    top.font, top.alignment = Font(size=8, bold=True, color=MUTED), Alignment(horizontal="center", vertical="bottom")
    v = ws.cell(row=CARD_ROW + 1, column=a, value=value)
    v.font, v.alignment = Font(size=18, bold=True, color=INK), Alignment(horizontal="center", vertical="center")
    if fmt:
        v.number_format = fmt
    if sub is not None:
        s = ws.cell(row=CARD_ROW + 2, column=a, value=sub)
        s.font, s.alignment = Font(size=8, italic=True, color=MUTED), Alignment(horizontal="center", vertical="top")
    return v


def _note(cell, text):
    """How a number is worked out, as a hover note (keeps the table itself uncluttered)."""
    cell.comment = Comment(text, "Model")
    cell.comment.width, cell.comment.height = 320, 90


# ---- the tab ------------------------------------------------------------------------------------------------
def _use_theme():
    """Pick up the workbook's accent colour (xl_style.set_theme) for the module's colour constants."""
    global NAVY, CARD_BG, MEDIUM
    NAVY, CARD_BG = st.THEME["accent"], st.THEME["pale"]
    MEDIUM = Side(style="medium", color=NAVY)


def write_summary(ws, m):
    _use_theme()
    g, fin = m.grid, m.fin
    foreign = fin.currency != "USD"
    mk = Ctx(g, "SUM", g.last_fy).market
    dcf, mult = "dcf" in m.parts, "multiples" in m.parts
    has_methods = dcf or mult or bool(m.comps)
    ws.sheet_view.showGridLines = False
    for col, w in WIDTHS.items():
        ws.column_dimensions[col].width = w

    # ---- title bar and the line under it
    ws.merge_cells(f"A1:{LAST_COL}1")
    title = ws["A1"]
    title.value = f"{fin.name} ({fin.ticker}) - Summary"           # the roll-up reads the name from here
    title.font = Font(size=16, bold=True, color="FFFFFF")
    title.fill = PatternFill("solid", fgColor=NAVY)
    title.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 30
    ws.merge_cells(f"A2:{LAST_COL}2")
    val_date = xl_dcf.ref("val_date") if dcf else "TODAY()"
    unit = "$ per share" if not foreign else (f"per-share values in {fin.currency} per ordinary share; the $ column "
                                              "converts to dollars per US-listed share")
    ws["A2"] = (f'="Valuation summary  |  "&{mk("Scenario")}&" scenario  |  valued "&TEXT({val_date},"mmm d, yyyy")'
                f'&"  |  {unit}  |  yellow cells are yours; click a method to go to its main input"')
    ws["A2"].font, ws["A2"].alignment = Font(italic=True, size=9, color=MUTED), Alignment(indent=1)
    ws.row_dimensions[3].height = 6
    ws.row_dimensions[7].height = 8

    price_usd = f'IFERROR(IF({mk("Price override")}<>"",{mk("Price override")},{mk("Live close")}),"")'
    price = mk("Share price")                    # reporting currency per ordinary share: what the values compare with

    # ---- three panels
    r0 = PANEL_ROW
    _section(ws, r0, "A", "B", "Snapshot")
    _kv(ws, r0 + 1, "Share price (US listing)", "=" + price_usd, st.FMT["usd_price"], bold=True)
    _kv(ws, r0 + 2, "Price date", f'=IF({mk("Price override")}<>"","override",{mk("Price date")})', st.FMT["date"])
    _kv(ws, r0 + 3, "Market cap (millions)", f"={mk('Market cap')}", st.FMT["usd"])
    _kv(ws, r0 + 4, "Scenario", f"={mk('Scenario')}")
    if foreign:
        _kv(ws, r0 + 5, f"Share price ({fin.currency} per ordinary share)", f"={price}", st.FMT["price"])
    _section(ws, r0, "D", "F", "Valuation inputs")
    from .xl_inputs import warnings_count
    side = [("Valuation date", f"={val_date}", st.FMT["date"]),
            ("WACC", f"={xl_dcf.ref('wacc')}" if dcf else "n/a (no DCF)", st.FMT["pct"]),
            ("Terminal growth", f"={mk('g')}", st.FMT["pct"]),
            ("Model checks", f"='{SHEET_NAMES['CHK']}'!$D$1", None),
            ("Input warnings", f'=IF({warnings_count()}=0,"none",{warnings_count()}&" on Inputs")', None)]
    if foreign:
        side.append((f"{fin.currency} per US dollar", f'=IF({mk("FX override")}<>"",{mk("FX override")},{mk("FX rate")})',
                     st.FMT["fx"]))
    for i, (label, f, fmt) in enumerate(side):
        _kv(ws, r0 + 1 + i, label, f, fmt, col="D", value_col="F")
    _fcf_yield(ws, m, dcf, r0)

    if not has_methods:
        _cards(ws, m, None, price_usd)
        r = TABLE_TOP
        ws.cell(row=r, column=1, value="Built with the three statements only: rebuild with --parts all for the DCF, "
                                       "multiples and comps.").font = Font(italic=True, color=MUTED)
        kf = _key_figures(ws, m, r + 2, foreign)
        _revenue_chart(ws, m, kf + 1)
        _finish(ws)
        return

    # ---- value per share by method
    top = TABLE_TOP
    _section(ws, top, "A", "G", "Value per share by method",
             "Hover over a method for how it is worked out. The weights (yellow) set the price target; methods "
             "without a value are left out and the rest scaled to 100%.")
    hdr = top + 2
    heads = ["Method", "Low", "Value", "High", "vs. price", "$ per US share" if foreign else "", "Weight"]
    for j, h in enumerate(heads, start=1):
        if h or j == 6:
            st.header_cell(ws.cell(row=hdr, column=j), h, "label" if j == 1 else ("ltm" if j == 7 else "fcst"))
    if foreign:
        ws.cell(row=hdr, column=6).value = "$ per US-listed share"     # the roll-up looks for this heading
    ws.cell(row=hdr, column=1).alignment = Alignment(horizontal="left", indent=1)
    to_usd = (f'*{mk("Shares per listing")}/IF({mk("FX override")}<>"",{mk("FX override")},{mk("FX rate")})')
    conv = f'*IF({mk("FX override")}<>"",{mk("FX override")},{mk("FX rate")})/{mk("Shares per listing")}'
    g1, g2 = (f"DCF!{rng}" for rng in xl_dcf.GRID_CELLS)
    lo_of = lambda rng: f'IF(COUNT({rng})=0,"",MIN({rng}))'
    hi_of = lambda rng: f'IF(COUNT({rng})=0,"",MAX({rng}))'
    goto = {"DCF: perpetuity growth": ("DCF", f"D{xl_dcf.R['g']}"),
            "DCF: exit multiple": ("DCF", f"D{xl_dcf.R['multiple']}"), "Peer multiples": ("Comps", "A1")}
    goto.update({label: ("Multiples", f"{xl_multiples.COL['override']}{xl_multiples.R[key]}")
                 for key, label, *_ in xl_multiples.METHODS})
    rows = []
    if dcf:
        rows += [("DCF: perpetuity growth", lo_of(g1), xl_dcf.ref("per_share", "B"), hi_of(g1),
                  "Free cash flow over the forecast plus a terminal value growing at the terminal growth rate, "
                  "discounted at the WACC. Range: the WACC x terminal growth grid on the DCF tab."),
                 ("DCF: exit multiple", lo_of(g2), xl_dcf.ref("per_share", "C"), hi_of(g2),
                  "Free cash flow over the forecast plus a terminal value at the exit EV/EBITDA multiple on final-year "
                  "EBITDA. Range: the WACC x exit multiple grid on the DCF tab.")]
    if mult:
        C = xl_multiples.COL
        for key, label, *_ in xl_multiples.METHODS:
            rows.append((label, xl_multiples.ref(key, C["v_low"]), xl_multiples.ref(key, C["value"]),
                         xl_multiples.ref(key, C["v_high"]),
                         "A target multiple on next-twelve-months figures from your forecast (Multiples tab). Low / "
                         "High use the range there. Blank when switched off or the figure isn't positive."))
    if m.comps and "_implied" in m.comps:
        imp = m.comps["_implied"]
        rows.append(("Peer multiples", lo_of(imp), f"MEDIAN({imp})", hi_of(imp),
                     "EV/EBITDA, EV/Revenue and P/E at the peer medians, on reported figures (Comps tab)."))
    rows.append(("52-week trading range", f'{mk("52w low")}{conv}', price, f'{mk("52w high")}{conv}',
                 "Lowest and highest close of the last year."))
    first = hdr + 1
    for i, (label, lo, val, hi, how) in enumerate(rows):
        r = first + i
        a = ws.cell(row=r, column=1, value=label)
        a.alignment = Alignment(indent=1)
        if label in goto:
            st.link_to(a, *goto[label])
        _note(a, how)
        for j, f in ((2, lo), (3, val), (4, hi)):
            c = ws.cell(row=r, column=j, value=f'=IFERROR({f},"")')
            c.number_format = st.FMT["price"]
            c.font = Font(color=INK, bold=j == 3)
        if label != "52-week trading range":
            c = ws.cell(row=r, column=5, value=f'=IFERROR(C{r}/{price}-1,"")')
            c.number_format = st.FMT["pct"]
        if foreign:
            c = ws.cell(row=r, column=6, value=f'=IFERROR(C{r}{to_usd},"")')
            c.number_format = st.FMT["usd_price"]
        if label in WEIGHTS:
            c = ws.cell(row=r, column=WEIGHT_COL, value=WEIGHTS[label])
            c.number_format, c.alignment = st.FMT["pct"], Alignment(horizontal="right")
            m.reg.add(ws, c, "Summary", "Weight", label)
        for j in range(1, WEIGHT_COL + 1):
            ws.cell(row=r, column=j).border = Border(bottom=Side(style="hair", color=RULE))
    last = first + len(rows) - 1
    pr = last + 1
    ws.cell(row=pr, column=1, value="Current share price").font = Font(bold=True, color=INK)
    ws.cell(row=pr, column=1).alignment = Alignment(indent=1)
    for j in (2, 3, 4):
        c = ws.cell(row=pr, column=j, value=f'=IFERROR({price}*{(0.995, 1, 1.005)[j - 2]},"")')
        c.number_format = st.FMT["price"]
        c.font = Font(color=MUTED, bold=j == 3)

    # ---- price target: the methods weighted as you choose
    pt = pr + 1
    m.target_ref = f"Summary!$C${pt}"
    W = lambda col: f"{col}{first}:{col}{last - 1}"          # the methods (not the 52-week range below them)
    wcol = get_column_letter(WEIGHT_COL)
    gold = PatternFill("solid", fgColor=GOLD_BG)
    ws.cell(row=pt, column=1, value=TARGET)
    for j, col in ((2, "B"), (3, "C"), (4, "D")):
        den = f"SUMPRODUCT({W(wcol)},--ISNUMBER({W(col)}))"
        c = ws.cell(row=pt, column=j, value=f'=IFERROR(IF({den}<=0,"",SUMPRODUCT({W(wcol)},{W(col)})/{den}),"")')
        c.number_format = st.FMT["price"]
    c = ws.cell(row=pt, column=5, value=f'=IFERROR(C{pt}/{price}-1,"")')
    c.number_format = st.FMT["pct"]
    c = ws.cell(row=pt, column=WEIGHT_COL, value=f"=SUMPRODUCT({W(wcol)},--ISNUMBER({W('C')}))")
    c.number_format, c.alignment = st.FMT["pct"], Alignment(horizontal="right")
    _note(c, "Weight in use: the weights of methods that have a value. The target scales them back to 100%.")
    # 12 months out, the price should have earned the return shareholders require, less what is paid out
    t12, roll = pt + 1, pt + 2
    ws.cell(row=t12, column=1, value=TARGET_12M)
    for j, col in ((2, "B"), (3, "C"), (4, "D")):
        c = ws.cell(row=t12, column=j, value=f'=IFERROR(IF(OR({col}{pt}="",C{roll}=""),"",{col}{pt}*(1+C{roll})),"")')
        c.number_format = st.FMT["price"]
    c = ws.cell(row=t12, column=5, value=f'=IFERROR(C{t12}/{price}-1,"")')
    c.number_format = st.FMT["pct"]
    if foreign:
        for r in (pt, t12):
            c = ws.cell(row=r, column=6, value=f'=IFERROR(C{r}{to_usd},"")')
            c.number_format = st.FMT["usd_price"]
    for r in (pt, t12):
        for j in range(1, WEIGHT_COL + 1):
            c = ws.cell(row=r, column=j)
            c.fill = gold
            c.font = Font(bold=True, color=INK, size=12 if r == pt else 11)    # numbers in ink, never the brand colour
            c.border = Border(top=Side(style="thin", color=NAVY) if r == pt else None)
        ws.cell(row=r, column=1).alignment = Alignment(indent=1)
    ws.cell(row=pt, column=5).number_format = ws.cell(row=t12, column=5).number_format = st.FMT["pct"]
    a = ws.cell(row=roll, column=1, value=ROLL)
    a.font, a.alignment = Font(color=MUTED, size=9), Alignment(indent=1)
    c0 = Ctx(g, "SUM", g.fcst[0])
    dps_yield = f'IFERROR({c0.n("Dividends Paid")}/{mk("Diluted shares")}/{price},0)'
    default = f"=IFERROR(ROUND({xl_wacc.ref('ke')}-{dps_yield},4),\"\")" if dcf else None
    d = ws.cell(row=roll, column=2, value=default)                 # default (B) / used (C) / your override (D)
    d.number_format, d.font = st.FMT["pct"], Font(color=MUTED, size=9)
    c = ws.cell(row=roll, column=3, value=f'=IF(D{roll}="",IF(B{roll}="","",B{roll}),D{roll})')
    c.number_format, c.font = st.FMT["pct"], Font(bold=True, size=9)
    o = ws.cell(row=roll, column=4)
    o.number_format = st.FMT["pct"]
    m.reg.add(ws, o, "Summary", "Roll-forward return")
    _note(a, ("Grey (default): cost of equity from the DCF tab less next year's dividend yield, since a share is "
              "expected to rise by the return its owners require, less what it pays out. Type your own in the "
              "yellow cell; the bold one is used.") if dcf else
          "Type the yearly return you expect (e.g. 9%) in the yellow cell for a 12-month target: there is no DCF "
          "tab to take the cost of equity from.")

    # upside in green / red, weights as bars
    for rng in (f"E{first}:E{t12}",):
        ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"], font=UP))
        ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], font=DOWN))
    ws.conditional_formatting.add(f"{wcol}{first}:{wcol}{last - 1}",
                                  DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                              color=st.THEME["light"]))

    _cards(ws, m, (pt, t12, roll, first, last), price_usd)
    _thesis(ws, top)

    # ---- charts: football field (left) and revenue and margins (right)
    ch = roll + 3
    _section(ws, ch, "A", "G", "Range of values")
    _section(ws, ch, "I", LAST_COL, "Revenue and operating margin" + (" ($ billions)" if m.fin.currency == "USD" else f" ({m.fin.currency} billions)"))
    for r in range(first, pt + 1):
        ws.cell(row=r, column=FF_COL, value=f"=N(B{r})")
        ws.cell(row=r, column=FF_COL + 1, value=f"=MAX(0,N(D{r})-N(B{r}))")
    ws.cell(row=hdr, column=FF_COL, value="start")
    ws.cell(row=hdr, column=FF_COL + 1, value="length")
    _football(ws, hdr, first, pt, f"A{ch + 1}")
    _revenue_chart(ws, m, ch + 1, anchor=f"I{ch + 1}")

    # ---- market vs. your model
    mv = ch + 18
    _section(ws, mv, "A", "G", "What the share price implies")
    _section(ws, mv, "I", LAST_COL, "Today's multiples")
    if dcf:
        for j, h in enumerate(("", "Perpetuity", "Exit multiple"), start=1):
            st.header_cell(ws.cell(row=mv + 1, column=j), h, "label" if j == 1 else "fcst")
        lines = [("Revenue growth per year in the model", "model_cagr", "pct", "pct"),
                 ("Revenue growth per year the price needs", "imp_cagr", "pct", "pct"),
                 ("Free cash flow needed vs. the model", "imp_scale", "x", "x"),
                 ("Terminal growth / exit multiple the price needs", "imp_terminal", "pct", "x")]
        for i, (label, key, fb, fc) in enumerate(lines):
            r = mv + 2 + i
            _kv(ws, r, label, f"={xl_dcf.ref(key, 'B')}", st.FMT[fb])
            c = ws.cell(row=r, column=3, value=f"={xl_dcf.ref(key, 'C')}")
            c.number_format, c.alignment = st.FMT[fc], Alignment(horizontal="right")
            c.border = Border(bottom=Side(style="hair", color=RULE))
        ws.cell(row=mv + 6, column=1, value="Compare the growth the price needs with the growth in your model: the "
                                            "gap is what the market doubts, or what it expects beyond your forecast."
                ).font = Font(italic=True, size=9, color=MUTED)
    else:
        ws.cell(row=mv + 1, column=1, value="Build with the DCF for the reverse DCF.").font = Font(italic=True,
                                                                                                    color=MUTED)
    if mult:
        C, R = xl_multiples.COL, xl_multiples.R
        for j, h in enumerate(("", "Your forecast", "Reported", "Peer median"), start=_col("I")):
            st.header_cell(ws.cell(row=mv + 1, column=j), h, "label" if h == "" else "fcst")
        for i, (key, label, *_rest) in enumerate(xl_multiples.METHODS):
            fmt = st.FMT["num" if key == "peg" else "x"]
            r = mv + 2 + i
            ws.cell(row=r, column=_col("I"), value=label).font = Font(color=INK)
            for j, name in enumerate(("now", "now_ltm", "peer"), start=_col("J")):
                ref = xl_multiples.ref(key, C[name])
                c = ws.cell(row=r, column=j, value=f'=IF({ref}="","",{ref})')
                c.number_format, c.alignment = fmt, Alignment(horizontal="right")
            for j in range(_col("I"), _col(LAST_COL) + 1):
                ws.cell(row=r, column=j).border = Border(bottom=Side(style="hair", color=RULE))
        ws.cell(row=mv + 2 + len(xl_multiples.METHODS), column=_col("I"),
                value="Today's price on your forecast (next 12 months) and on reported results (last 12 months); "
                      "peer medians are trailing, so compare them with Reported.").font = Font(italic=True, size=9,
                                                                                                color=MUTED)

    # ---- key figures
    _key_figures(ws, m, mv + 9, foreign)
    _finish(ws)


def _finish(ws):
    for j in range(RV_COL, FF_COL + 2):                  # chart data
        ws.column_dimensions[get_column_letter(j)].hidden = True
    ws.freeze_panes = "A3"


def _thesis(ws, top):
    """The first lines of the thesis and risks from the Notes tab, beside the valuation table."""
    from .xl_inputs import note_lines
    _section(ws, top, "I", LAST_COL, "Your thesis")
    st.link_to(ws.cell(row=top + 1, column=_col("I")), "Notes", "A5", "Write it on the Notes tab")
    ws.cell(row=top + 1, column=_col("I")).font = Font(size=9, color="0563C1", underline="single")
    r = top + 2
    for title, n, prompt in (("Thesis", 6, "Why you own it, in a sentence or two"),
                             ("Risks and what would change my mind", 4, "What would make you sell")):
        lines = list(note_lines(title))[:n]
        if title != "Thesis":
            ws.cell(row=r, column=_col("I"), value="Risks").font = Font(bold=True, size=9, color=MUTED)
            r += 1
        for k, nr in enumerate(lines):
            empty = f'"{prompt}"' if k == 0 else '""'
            c = ws.cell(row=r, column=_col("I"), value=f'=IF(Notes!$A${nr}="",{empty},"- "&Notes!$A${nr})')
            c.font = Font(size=9, color=INK if k else MUTED, italic=False)
            ws.merge_cells(start_row=r, start_column=_col("I"), end_row=r, end_column=_col(LAST_COL))
            c.alignment = Alignment(vertical="top", shrink_to_fit=True)
            r += 1


def _cards(ws, m, rows, price_usd):
    """The headline row. rows = (price target, 12-month target, roll-forward, first method, last row) or None."""
    fin = m.fin
    mk = Ctx(m.grid, "SUM", m.grid.last_fy).market
    foreign = fin.currency != "USD"
    _card(ws, "A", "A", "Share price", "=" + price_usd, st.FMT["usd_price"],
          f'=IF({mk("Price override")}<>"","price override",IFERROR("close of "&TEXT({mk("Price date")},"mmm d"),""))')
    if rows:
        pt, t12, roll, first, last = rows
        val = (lambda r: f'=IF(F{r}="","",F{r})') if foreign else (lambda r: f'=IF(C{r}="","",C{r})')
        n = (f'=SUMPRODUCT((G{first}:G{last - 1}>0)*ISNUMBER(C{first}:C{last - 1}))&" methods, your weights"')
        _card(ws, "B", "C", "Price target", val(pt), st.FMT["usd_price"], n)
        _card(ws, "D", "E", "Upside to target", f'=IF(E{pt}="","",E{pt})', '+0.0%;-0.0%;0.0%', "vs. today's price")
        ws.conditional_formatting.add(f"D{CARD_ROW + 1}", CellIsRule(operator="greaterThan", formula=["0"],
                                                                      font=Font(size=18, bold=True, color="1E7B34")))
        ws.conditional_formatting.add(f"D{CARD_ROW + 1}", CellIsRule(operator="lessThan", formula=["0"],
                                                                      font=Font(size=18, bold=True, color="B42318")))
        _card(ws, "F", "G", "12-month target", val(t12), st.FMT["usd_price"],
              f'=IF(C{roll}="","type a return below the table","rolled forward at "&TEXT(C{roll},"0.0%"))')
    else:
        _card(ws, "B", "C", "Market cap (millions)", f"={mk('Market cap')}", st.FMT["usd"], "")
    _card(ws, "I", "J", "Scenario", f"={mk('Scenario')}", None, "change it on the Inputs tab")
    st.link_to(ws.cell(row=CARD_ROW + 2, column=_col("I")), SHEET_NAMES["INP"], f"B{INPUT_ROWS['Scenario']}")
    ws.cell(row=CARD_ROW + 2, column=_col("I")).font = Font(size=8, color="0563C1", underline="single")
    _card(ws, "K", "L", "Model checks", f"='{SHEET_NAMES['CHK']}'!$D$1", None, "balance sheet and cash ties")
    ws.cell(row=CARD_ROW + 1, column=_col("K")).font = Font(size=11, bold=True, color=INK)
    ws.row_dimensions[CARD_ROW].height = 15
    ws.row_dimensions[CARD_ROW + 1].height = 30
    ws.row_dimensions[CARD_ROW + 2].height = 14


def _fcf_yield(ws, m, dcf, r0):
    """Free-cash-flow yield (FCF / market cap) on reported results, on your forecast and after stock-based
    compensation, with the rates to compare it against."""
    g = m.grid
    tcol = g.ltm if g.ltm is not None and m.has_data("Revenue", g.ltm) else g.last_fy
    t, y1, y2 = (Ctx(g, "SUM", c) for c in (tcol, g.fcst[0], g.fcst[1]))
    mcap = Ctx(g, "SUM", g.last_fy).market("Market cap")
    end = g.last_fy.end
    val = xl_dcf.ref("val_date") if dcf else "TODAY()"
    e = f"MAX(0,MIN(1,({val}-DATE({end.year},{end.month},{end.day}))/365))"
    ntm = f"({y1.n('Free cash flow')}*(1-{e})+{y2.n('Free cash flow')}*{e})"
    yld = lambda num: f'=IFERROR(IF(N({mcap})<=0,"",{num}/{mcap}),"")'
    rows = [("FCF yield, reported " + ("(LTM)" if tcol.kind == "ltm" else f"({tcol.header})"),
             yld(t.n("Free cash flow")), True),
            ("FCF yield, your forecast (NTM)", yld(ntm), False),
            ("FCF yield after stock-based comp (reported)", yld(f"({t.n('Free cash flow')}-{t.n('Stock-Based Comp')})"),
             False)]
    if dcf:
        rows += [("10-year Treasury yield", f"={xl_wacc.ref('rf')}", False),
                 ("Cost of equity", f"={xl_wacc.ref('ke')}", False)]
    col = get_column_letter(YIELD_COL)
    _section(ws, r0, col, get_column_letter(YIELD_COL + 1), "Free-cash-flow yield")
    for i, (label, f, bold) in enumerate(rows):
        c = _kv(ws, r0 + 1 + i, label, f, st.FMT["pct"], col=col, bold=bold)
        if i >= 3:
            c.font = Font(color=MUTED)
            ws.cell(row=r0 + 1 + i, column=YIELD_COL).font = Font(color=MUTED)


def _key_figures(ws, m, kf, foreign):
    """Key figures block starting at row kf. Returns the row below it."""
    g, fin = m.grid, m.fin
    _section(ws, kf, "A", "G", "Key figures" + ("" if not foreign else f" ({fin.currency} millions)"))
    cols = [g.hist[-2]] if len(g.hist) > 1 else []
    cols += [g.last_fy] + ([g.ltm] if g.ltm is not None else []) + g.fcst[:2] + ([g.fcst[-1]] if len(g.fcst) > 2 else [])
    cols = cols[:6]
    st.header_cell(ws.cell(row=kf + 1, column=1), "")
    for j, col in enumerate(cols, start=2):
        st.header_cell(ws.cell(row=kf + 1, column=j), col.header, col.kind)
    figures = [("Revenue", "Revenue", "usd"), ("Revenue growth", "Revenue growth", "pct"),
               ("Gross margin", "Gross margin", "pct"), ("Operating margin", "Operating margin", "pct"),
               ("Net income", "Net Income", "usd"), ("Diluted EPS", "Diluted EPS", "eps"),
               ("Free cash flow", "Free cash flow", "usd"), ("FCF margin", "FCF margin", "pct")]
    for i, (text, label, fmt) in enumerate(figures):
        r = kf + 2 + i
        ws.cell(row=r, column=1, value=text).alignment = Alignment(indent=1)
        for j, col in enumerate(cols, start=2):
            ref = Ctx(g, "SUM", col).v(label)
            c = ws.cell(row=r, column=j, value=f'=IF({ref}="","",{ref})')
            c.number_format = st.FMT[fmt]
            if col.kind == "fcst":
                c.fill = st.FCST_BG
        for j in range(1, len(cols) + 2):
            ws.cell(row=r, column=j).border = Border(bottom=Side(style="hair", color=RULE))
    return kf + len(figures) + 3


def _football(ws, hdr, first, last_row, anchor):
    """Low-to-high bar per method; the price target bar (the last row) in gold."""
    from openpyxl.chart.marker import DataPoint
    chart = BarChart()
    chart.type, chart.grouping, chart.overlap = "bar", "stacked", 100
    chart.add_data(Reference(ws, min_col=FF_COL, max_col=FF_COL + 1, min_row=hdr, max_row=last_row),
                   titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last_row))
    chart.series[0].graphicalProperties.noFill = True
    chart.series[0].graphicalProperties.line.noFill = True
    chart.series[1].graphicalProperties.solidFill = st.THEME["mid"]
    chart.series[1].graphicalProperties.line.noFill = True
    target = DataPoint(idx=last_row - first)
    target.graphicalProperties.solidFill = "BF8F00"
    chart.series[1].dPt.append(target)
    price = DataPoint(idx=last_row - first - 1)                # the current share price: a thin grey mark
    price.graphicalProperties.solidFill = "9CA3AF"
    chart.series[1].dPt.append(price)
    chart.legend = None
    chart.y_axis.number_format = "#,##0"
    chart.y_axis.txPr = _axis_text(800)
    chart.x_axis.txPr = _axis_text(800)
    chart.y_axis.delete = False
    chart.x_axis.delete = False
    chart.x_axis.scaling.orientation = "maxMin"          # first method at the top
    chart.y_axis.crosses = "max"                         # ... and the value axis along the bottom
    chart.y_axis.majorGridlines = None
    chart.gapWidth = 50
    chart.visible_cells_only = False                     # its data sits in hidden columns
    chart.height, chart.width = 7.6, 19
    ws.add_chart(chart, anchor)


def _axis_text(size, horizontal=False):
    """Axis label font size (hundredths of a point); horizontal=True stops Excel tilting the labels."""
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RichTextProperties
    body = RichTextProperties(rot=0, vert="horz") if horizontal else RichTextProperties()
    # r=[]: openpyxl otherwise adds an empty text run, which blanks the axis labels
    return RichText(bodyPr=body, p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=size)),
                                              endParaRPr=CharacterProperties(), r=[])])


def _revenue_chart(ws, m, row, anchor=None):
    """Revenue (bars, history then forecast) with operating margin (line, right axis), from hidden helper cells."""
    g = m.grid
    cols = g.hist + g.fcst
    r0 = 4
    for j, h in enumerate(("Year", "Revenue", "Operating margin"), start=RV_COL):
        ws.cell(row=r0, column=j, value=h)
    for i, col in enumerate(cols, start=1):
        c = Ctx(g, "SUM", col)
        ws.cell(row=r0 + i, column=RV_COL, value=col.header.replace("FY20", "FY"))     # FY2026E -> FY26E
        ws.cell(row=r0 + i, column=RV_COL + 1, value=f'=IFERROR({c.v("Revenue")}/1000000,"")')
        ws.cell(row=r0 + i, column=RV_COL + 2, value=f'=IFERROR({c.v("Operating margin")},"")')
    n = len(cols)
    bars = BarChart()
    bars.add_data(Reference(ws, min_col=RV_COL + 1, min_row=r0, max_row=r0 + n), titles_from_data=True)
    bars.set_categories(Reference(ws, min_col=RV_COL, min_row=r0 + 1, max_row=r0 + n))
    bars.series[0].graphicalProperties.solidFill = st.THEME["light"]
    bars.series[0].graphicalProperties.line.noFill = True
    from openpyxl.chart.marker import DataPoint
    for k in range(len(g.hist)):                          # history darker than the forecast
        dp = DataPoint(idx=k)
        dp.graphicalProperties.solidFill = st.THEME["accent"]
        bars.series[0].dPt.append(dp)
    bars.y_axis.number_format = '#,##0,"B"' if m.fin.currency == "USD" else '#,##0,"bn"'   # $ millions -> billions
    bars.y_axis.majorGridlines = None
    bars.x_axis.txPr = _axis_text(800, horizontal=True)
    bars.y_axis.txPr = _axis_text(800)
    bars.y_axis.delete = False
    bars.x_axis.delete = False
    bars.gapWidth = 40
    line = LineChart()
    line.add_data(Reference(ws, min_col=RV_COL + 2, min_row=r0, max_row=r0 + n), titles_from_data=True)
    line.series[0].graphicalProperties.line.solidFill = "BF8F00"
    line.series[0].graphicalProperties.line.width = 28575
    line.series[0].smooth = False                         # straight segments: no made-up curves between years
    line.y_axis.axId = 200
    line.y_axis.number_format = "0%"
    line.y_axis.crosses = "max"
    line.y_axis.majorGridlines = None
    line.y_axis.delete = False
    line.y_axis.txPr = _axis_text(800)
    bars += line
    bars.legend.position = "t"                            # above the plot, clear of the year labels
    bars.visible_cells_only = False
    bars.height, bars.width = 7.6, 15.5
    ws.add_chart(bars, anchor or f"A{row}")


BANNER_TABS = ("Inputs", "Drivers", "WACC", "DCF", "Multiples", "Revenue Build", "Comps")


def write_banners(wb, m):
    """One line of results at the top of every tab with inputs, so a change shows its effect where it is made:
    price target and upside, the two DCF values, the multiples average, and the scenario. Call after the
    Summary is written (the price target cell is placed there)."""
    from .layout import input_ref
    cur = m.fin.currency
    code = "$" if cur == "USD" else cur + " "
    money = lambda x: f'"{code}"&TEXT({x},"#,##0.00")'
    price = input_ref("Share price")
    pieces = [f'"Scenario: "&{input_ref("Scenario")}']
    target = getattr(m, "target_ref", None)
    if target:
        pieces.append(f'IFERROR("Price target "&{money(target)}&" ("&TEXT({target}/{price}-1,"+0.0%;-0.0%")'
                      f'&" vs. "&{money(price)}&")","Price target: n/a")')
    else:
        pieces.append(f'IFERROR("Price "&{money(price)},"")')
    if "dcf" in m.parts:
        pieces.append(f'IFERROR("DCF "&{money(xl_dcf.ref("per_share", "B"))}&" / "&{money(xl_dcf.ref("per_share", "C"))},"")')
    if "multiples" in m.parts:
        pieces.append(f'IFERROR("Multiples average "&{money(xl_multiples.ref("average"))},"")')
    formula = "=" + '&"      |      "&'.join(pieces)
    for name in BANNER_TABS:
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        title = str(ws["A1"].value or "")
        used = max([c.column for c in ws[1] if c.value is not None] + [1])
        need, j, width = len(title) * 1.25 + 2, 1, 0.0
        while True:                                     # first column clear of the title and of row 1's cells
            dim = ws.column_dimensions[get_column_letter(j)]
            width += dim.width if dim.width else 8.43
            j += 1
            if width >= need and j > used + 1:
                break
        c = ws.cell(row=1, column=j, value=formula)
        c.font = st.font("1E5C4F", True)
