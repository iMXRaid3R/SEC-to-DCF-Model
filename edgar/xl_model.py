"""LOAD: build the model workbook (Assumptions, IS, BS, CF, DCF, Checks & Metrics, Mapping, Raw)."""
from datetime import date

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from . import xl_dcf
from . import xl_style as st
from .comps import _metrics, write_comps
from .inputs import Registry
from .layout import (BASIS_CELL, BASIS_OPTIONS, BS_ROWS, CF_ROWS, CHECK_ROWS, DRIVER_ROWS, IS_ROWS,
                     METRIC_ROWS, input_ref,
                     FIRST_ROW, RB_SHEET, ROW_BY_LABEL, SCENARIO_CELL, SCENARIOS, SHEET_NAMES, Ctx, Grid)
from .mapping import MAP_BY_LABEL, fmt_tag
from .xl_build import write_build
from .xl_dcf import write_dcf
from .xl_inputs import write_inputs, write_notes
from .xl_summary import write_summary

STATEMENTS = (("IS", IS_ROWS), ("BS", BS_ROWS), ("CF", CF_ROWS))


class Model:
    """Grid + SEC values per column, and the rule for when a fallback formula can be used."""

    kind = "standard"       # "fin" for the financials template (xl_fin.py)

    def __init__(self, fin, n_fcst, ytd=True):
        self.fin = fin
        self.grid = Grid(fin.cols, fin.ltm_end, n_fcst, fin.ytd_months if fin.ytd and ytd else 0)
        self.settings = {}      # DCF settings (label, default, note, options) shown on the Inputs tab
        self.data = {}      # label -> {column index: value}
        self.reg = Registry()   # every input cell, so a rebuild can keep what the user typed
        self.comps = None       # {multiple: peer-median cell reference} when a Comps tab is built
        for col in self.grid.cols:
            if col.kind == "hist":
                for label, series in fin.values.items():
                    if col.end in series:
                        self.data.setdefault(label, {})[col.j] = float(series[col.end])
            elif col.kind == "ltm":
                for label, v in fin.ltm.items():
                    self.data.setdefault(label, {})[col.j] = float(v)
            elif col.kind == "ytd":
                for label, v in fin.ytd.items():
                    self.data.setdefault(label, {})[col.j] = float(v)

    def has_data(self, label, col):
        return col.j in self.data.get(label, {})

    def available(self, label, col, seen=frozenset()):
        """True if the cell will hold a value: SEC data, a plug, or a fallback whose inputs are available."""
        if self.has_data(label, col):
            return True
        row = ROW_BY_LABEL[label]
        if row.plug:
            return True
        if row.fallback and label not in seen:
            return all(self.available(n, col, seen | {label}) for n in row.needs)
        return False

    def uses_fallback(self, row, col):
        return (row.data and row.fallback is not None and not self.has_data(row.label, col)
                and all(self.available(n, col, frozenset({row.label})) for n in row.needs))

    def reported_anywhere(self, row):
        past = [c for c in self.grid.cols if c.kind != "fcst"]
        return any(self.has_data(row.label, c) or self.uses_fallback(row, c) for c in past)


# --------------------------------------------------------------------------
def write_header(ws, grid, title, subtitle, first="($ millions)"):
    first = first.replace("$", st.CURRENCY) if st.CURRENCY != "USD" else first
    ws["A1"], ws["A1"].font = title, st.TITLE_FONT
    ws["A2"], ws["A2"].font = subtitle, st.NOTE_FONT
    st.header_cell(ws.cell(row=4, column=1), first)
    ws.cell(row=4, column=1).alignment = Alignment(horizontal="left")
    for col in grid.cols:
        if col.end is not None:
            c = ws.cell(row=3, column=col.j, value=col.end.date())
            c.number_format, c.font, c.alignment = "yyyy-mm-dd", st.NOTE_FONT, st.CENTER
        st.header_cell(ws.cell(row=4, column=col.j), col.header, col.kind)
    ws.column_dimensions["A"].width = 38
    for col in grid.cols:
        ws.column_dimensions[col.L].width = 13
    ws.freeze_panes = "B5"


def _put_formula(cell, f, bold=False):
    cell.value = f
    cell.font = st.formula_font(f, bold)


def write_rows(ws, key, rows, m):
    g = m.grid
    for row in rows:
        r = g.rows[row.label][1]
        label = ws.cell(row=r, column=1, value=row.label)
        label.font = st.font(bold=row.bold)
        label.alignment = Alignment(indent=row.indent)
        if row.data and not m.reported_anywhere(row):
            label.value = f"{row.label}  (not reported)"
            label.font = st.font(st.GREY, italic=True)
        for col in g.cols:
            cell = ws.cell(row=r, column=col.j)
            cell.number_format = st.FMT[row.fmt]
            c = Ctx(g, key, col)
            if col.kind == "fcst":
                cell.fill = st.FCST_BG
                f = row.fcst(c) if row.fcst else None
                if f:
                    _put_formula(cell, f, row.bold)
            elif row.data and m.has_data(row.label, col):
                cell.value = m.data[row.label][col.j]
                cell.font = st.font(st.BLUE, row.bold)
            elif row.plug:
                f = row.plug(c)
                if f:
                    _put_formula(cell, f, row.bold)
            elif m.uses_fallback(row, col):
                _put_formula(cell, row.fallback(c), row.bold)


def write_statement(ws, key, rows, m):
    fin = m.fin
    write_header(ws, m.grid, f"{fin.name} ({fin.ticker}) - {SHEET_NAMES[key]}",
                 "Blue = SEC filings, black = formula, green = link. Shaded columns are forecasts driven by the "
                 "Assumptions tab." + ("" if fin.currency == "USD" else f" Amounts are in {fin.currency}."))
    write_rows(ws, key, rows, m)


def _choices(options):
    return "{" + ",".join(f'"{o}"' for o in options) + "}"


def _dropdown(ws, addr, label_addr, label, value, options):
    ws[label_addr], ws[label_addr].font = label, st.font(bold=True)
    ws[label_addr].alignment = Alignment(horizontal="right")
    c = ws[addr]
    c.value, c.font, c.fill, c.alignment = value, st.font(st.BLUE, True), st.INPUT_FILL, st.CENTER
    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(c)
    return c


def write_assumptions(ws, m):
    g, fin = m.grid, m.fin
    write_header(ws, g, f"{fin.name} ({fin.ticker}) - Drivers",
                 "Forecast columns follow the selected scenario's year-1, mid-point and final-year values (table "
                 "below). Type over any forecast cell to override that year (it turns yellow).", first="Driver")
    # The scenario and the base-case starting point are chosen on the Inputs tab; shown here for reference
    # (the formulas on this tab read these two cells).
    for addr, label_addr, label, key in ((SCENARIO_CELL, "E1", "Scenario:", "Scenario"),
                                         (BASIS_CELL, "H1", "Base starts from:", "Basis")):
        ws[label_addr], ws[label_addr].font = label, st.font(bold=True)
        ws[label_addr].alignment = Alignment(horizontal="right")
        c = ws[addr.replace("$", "")]
        c.value, c.font, c.alignment = f"={input_ref(key)}", st.font(st.GREEN, True), st.CENTER
    ws["K1"], ws["K1"].font = "(set on the Inputs tab)", st.NOTE_FONT
    for title, hr in g.driver_headings.items():
        st.header_cell(ws.cell(row=hr, column=1), title, "ltm")
        ws.cell(row=hr, column=1).alignment = Alignment(horizontal="left")
    pick = f"MATCH({SCENARIO_CELL},{_choices(SCENARIOS)},0)"
    basis_pick = f"MATCH({BASIS_CELL},{_choices(BASIS_OPTIONS)},0)"

    # ---- live driver rows: history = actual ratios, forecast = path through the selected scenario's three points
    hist3, last, n = g.hist[-3:], g.last_fy, len(g.fcst)
    mid = (n + 1) // 2                                   # forecast year the "Mid" input applies to (5 of 10)
    SEL = ("K", "L", "M")                                # Selected: year 1, mid, final
    for row in DRIVER_ROWS:
        r, sr = g.rows[row.label][1], g.scen_rows[row.label]
        ws.cell(row=r, column=1, value=row.label)
        for col in g.cols:
            cell = ws.cell(row=r, column=col.j)
            cell.number_format = st.FMT[row.fmt]
            if col.kind != "fcst":
                f = row.plug(Ctx(g, "A", col))
                if f:
                    _put_formula(cell, f)
                continue
            cell.fill = st.FCST_BG
            k = g.fcst.index(col) + 1                    # forecast year 1..n
            y1, ym, yn = (f"${L}${sr}" for L in SEL)
            if k == 1:
                cell.value = f"={y1}"
            elif k <= mid:
                cell.value = f"={y1}+({ym}-{y1})*{k - 1}/{mid - 1}"
            else:
                cell.value = f"={ym}+({yn}-{ym})*{k - mid}/{n - mid}"
            cell.font = st.font()
            m.reg.add(ws, cell, "Live", row.label, col.header)
        if row.note:
            ws.cell(row=r, column=g.cols[-1].j + 1, value=row.note).font = st.NOTE_FONT
    first_r, last_r = FIRST_ROW, g.driver_last
    top = f"{g.fcst[0].L}{first_r}"
    ws.conditional_formatting.add(                      # a typed number in a forecast cell = a per-year override
        f"{top}:{g.fcst[-1].L}{last_r}",
        FormulaRule(formula=[f'AND(NOT(ISFORMULA({top})),{top}<>"")'], font=st.font(st.BLUE, True), fill=st.INPUT_FILL))

    # ---- scenario inputs: year 1, a mid-point and the final year for Bear / Base / Bull, plus the selected set
    t = g.scen_title
    points = ("Year 1", "Mid", "Final year")
    heads = (g.fcst[0].header, g.fcst[mid - 1].header, g.fcst[-1].header)
    ws.cell(row=t, column=1, value="Scenario inputs").font = st.font(bold=True)
    ws.cell(row=t, column=2, value="Each driver runs in straight lines from the first forecast year to the mid-point to "
                                   "the final year, which should be a steady state. Base year 1 comes from history "
                                   "(\"Base starts from\"); growth fades toward the DCF terminal growth; Bear and "
                                   "Bull default to Base -/+ a spread on growth, gross margin and SG&A. Overwrite "
                                   "any blue cell.").font = st.NOTE_FONT
    st.header_cell(ws.cell(row=t + 1, column=1), "")
    st.header_cell(ws.cell(row=t + 2, column=1), "Driver")
    for i, name in enumerate(SCENARIOS + ("Selected",)):
        for k, sub in enumerate(heads):
            j = 2 + 3 * i + k
            st.header_cell(ws.cell(row=t + 1, column=j), name, "ltm" if name == "Selected" else "label")
            st.header_cell(ws.cell(row=t + 2, column=j), sub, "ltm" if name == "Selected" else "label")
    g_term = xl_dcf.ref("g")
    for row in DRIVER_ROWS:
        r, sr = g.rows[row.label][1], g.scen_rows[row.label]
        ws.cell(row=sr, column=1, value=row.label)
        if row.default == "zero":
            base1 = 0
        else:
            err = row.err_default
            fy = f'IF({last.L}{r}="",{err},{last.L}{r})'
            ltm = f'IF({g.ltm.L}{r}="",{fy},{g.ltm.L}{r})' if g.ltm is not None else fy
            pickd = f"IFERROR(CHOOSE({basis_pick},AVERAGE({hist3[0].L}{r}:{hist3[-1].L}{r}),{fy},{ltm}),{err})"
            base1 = f"=MIN({row.base_cap},{pickd})" if row.base_cap is not None else f"={pickd}"
        if row.label == "Revenue growth":
            # Growth halves by the mid-point and is halfway to the terminal rate by the final year, never below it.
            base = (base1, f"=MAX({g_term},E{sr}/2)", f"=MAX({g_term},(F{sr}+{g_term})/2)")
        else:
            base = (base1, f"=E{sr}", f"=F{sr}")

        def shifted(cell, sign):
            """Bear (sign -1) / Bull (sign +1) default as a formula off the Base cell."""
            if not row.spread:
                return f"={cell}"
            kind, x = row.spread
            op = "+" if sign * x > 0 else "-"
            f = f"{cell}*(1{op}{abs(x)})" if kind == "rel" else f"{cell}{op}MAX(0.01,ABS({cell})*{abs(x)})"
            return f"=MIN({row.cap},{f})" if row.cap is not None else f"={f}"

        values = {"Bear": tuple(shifted(f"{L}{sr}", -1) for L in "EFG"), "Base": base,
                  "Bull": tuple(shifted(f"{L}{sr}", 1) for L in "EFG")}
        for i, name in enumerate(SCENARIOS):
            for k, sub in enumerate(points):
                c = ws.cell(row=sr, column=2 + 3 * i + k, value=values[name][k])
                c.number_format, c.font = st.FMT[row.fmt], st.font(st.BLUE)
                m.reg.add(ws, c, "Scenario", row.label, name, sub)
        for k, cols in enumerate((("B", "E", "H"), ("C", "F", "I"), ("D", "G", "J"))):
            c = ws.cell(row=sr, column=11 + k, value=f"=CHOOSE({pick}," + ",".join(f"{L}{sr}" for L in cols) + ")")
            c.number_format, c.font = st.FMT[row.fmt], st.font(bold=True)

    # ---- headline results for the selected scenario (flip the Scenario dropdown to compare).
    # Not an Excel Data Table: with the live STOCKHISTORY price in the chain, a data table makes Excel
    # recalculate for a minute or more after every edit.
    v = g.val_title
    ws.cell(row=v, column=1, value="Results for the selected scenario").font = st.font(bold=True)
    ws.cell(row=v, column=2, value="Change the Scenario dropdown (F1) to see Bear / Base / Bull.").font = st.NOTE_FONT
    heads = [("DCF: perpetuity growth", xl_dcf.ref("per_share", "B"), "price"),
             ("DCF: exit multiple", xl_dcf.ref("per_share", "C"), "price"),
             ("Upside (perpetuity)", xl_dcf.ref("upside", "B"), "pct"),
             ("Upside (exit multiple)", xl_dcf.ref("upside", "C"), "pct"),
             (f"{g.fcst[-1].header} revenue", f"'{SHEET_NAMES['IS']}'!${g.fcst[-1].L}${g.rows['Revenue'][1]}", "usd"),
             (f"{g.fcst[-1].header} EPS", f"'{SHEET_NAMES['IS']}'!${g.fcst[-1].L}${g.rows['Diluted EPS'][1]}", "eps")]
    st.header_cell(ws.cell(row=v + 1, column=1), "Scenario")
    c = ws.cell(row=v + 2, column=1, value=f"={SCENARIO_CELL}")
    c.font = st.font(bold=True)
    for i, (h, target, fmt) in enumerate(heads):
        j = 2 + i
        st.header_cell(ws.cell(row=v + 1, column=j), h, "fcst")
        ws.cell(row=v + 1, column=j).alignment = Alignment(horizontal="center", wrap_text=True)
        c = ws.cell(row=v + 2, column=j, value=f"={target}")
        c.number_format, c.font = st.FMT[fmt], st.font(bold=True)
    ws.row_dimensions[v + 1].height = 30

    ws.column_dimensions[get_column_letter(g.cols[-1].j + 1)].width = 60


def write_checks(ws, m):
    g = m.grid
    write_header(ws, g, "Model Integrity Checks & Metrics",
                 "Checks are live formulas and should all be 0 (n/a = an input isn't reported).")
    first, last = g.rows[CHECK_ROWS[0].label][1], g.rows[CHECK_ROWS[-1].label][1]
    rng = f"B{first}:{g.cols[-1].L}{last}"
    ws["C1"] = "Status:"
    ws["C1"].font = st.font(bold=True)
    rev = Ctx(g, "CHK", g.last_fy).ref("Revenue", g.last_fy, abs_col=True, abs_row=True)
    ws["F1"], ws["F1"].font = "Tolerance:", st.NOTE_FONT
    ws["G1"] = f"=MAX(1000,0.002*N({rev}))"
    ws["G1"].number_format, ws["G1"].font = st.FMT["usd"], st.NOTE_FONT
    ws["H1"], ws["H1"].font = "0.2% of last fiscal year revenue: smaller differences are rounding or minor items", st.NOTE_FONT
    ws["D1"] = (f'=IF(COUNTIF({rng},">"&$G$1)+COUNTIF({rng},"<-"&$G$1)=0,"ALL CHECKS OK",'
                f'"CHECK ERRORS - see rows {first}-{last}")')
    ws["D1"].font = st.font(bold=True)
    write_rows(ws, "CHK", CHECK_ROWS, m)
    ws.cell(row=g.metric_header, column=1, value="Metrics").font = st.font(bold=True)
    write_rows(ws, "CHK", METRIC_ROWS, m)


def write_mapping(ws, fin):
    tags_for = {label: tags for mp in fin.maps for label, _, _, tags in mp} or {k: v[3] for k, v in MAP_BY_LABEL.items()}
    ws.append(["Statement", "Line Item", "Source", "Tag priority (first with data wins) / formula"])
    for c in ws[1]:
        c.fill, c.font = st.FILL["label"], st.HDR_FONT
    for key, rows in STATEMENTS:
        for row in rows:
            if row.data:
                tags = " > ".join(fmt_tag(t) for t in tags_for[row.label])
                ws.append([SHEET_NAMES[key], row.label, "SEC tags", tags + (f"   |   {row.note}" if row.note else "")])
            else:
                ws.append([SHEET_NAMES[key], row.label, "Formula", row.note])
    for col, w in zip("ABCD", (18, 32, 10, 140)):
        ws.column_dimensions[col].width = w


def write_raw(ws, src):
    if src.empty:
        ws.append(["No data"])
        return
    ws.append(list(src.columns))
    for c in ws[1]:
        c.fill, c.font = st.FILL["label"], st.HDR_FONT
    for r in src.itertuples(index=False):
        ws.append(list(r))
    for i, w in enumerate((30, 12, 70, 18, 12, 8, 12, 24), start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


TAB_COLORS = {"Summary": "1F3864", "Inputs": "BF8F00", "Drivers": "BF8F00", RB_SHEET: "BF8F00",
              "Income Statement": "1E5C4F", "Balance Sheet": "1E5C4F", "Cash Flow": "1E5C4F",
              "DCF": "7030A0", "Comps": "7030A0", "Checks & Metrics": "7030A0",
              "Notes": "808080", "Mapping": "808080", "Raw": "808080"}
TAB_ORDER = ["Summary", "Inputs", "Drivers", RB_SHEET, "Income Statement", "Balance Sheet", "Cash Flow", "DCF",
             "Comps", "Checks & Metrics", "Notes", "Mapping", "Raw", xl_dcf.CALC_SHEET]


def build_workbook(fin, n_fcst=10, peers=(), template=None):
    """template: "standard", "financial", or None to pick the financials template for banks and insurers.
    peers: _metrics() rows for comparable companies; adds a Comps tab and links the DCF exit multiple to it.
    Tabs: Summary | Inputs, Drivers, Revenue Build (yours to change) | the three statements | DCF, Comps, Checks
    | Notes, Mapping, Raw (reference). Working cells and the input registry sit on hidden sheets."""
    st.set_currency(fin.currency)
    if template == "financial" or (template is None and fin.looks_financial()):
        from .xl_fin import build_financial
        return build_financial(fin, n_fcst)
    m = Model(fin, n_fcst)
    wb = Workbook()
    wb.active.title = SHEET_NAMES["A"]
    sheets = {"A": wb.active}
    for name in TAB_ORDER:
        if name == "Comps" and not peers:
            continue
        if name not in wb.sheetnames:
            wb.create_sheet(name)
    wb._sheets = [wb[n] for n in TAB_ORDER if n in wb.sheetnames]
    for key in ("IS", "BS", "CF", "DCF", "CHK", "SUM", "INP"):
        sheets[key] = wb[SHEET_NAMES[key]]
    if peers:
        m.comps = write_comps(list(peers), wb["Comps"], subject=_metrics(fin),
                              price_ref=Ctx(m.grid, "CMP", m.grid.last_fy).market("Share price"), reg=m.reg)
    write_assumptions(sheets["A"], m)
    write_build(wb[RB_SHEET], m, write_header)
    for key, rows in STATEMENTS:
        write_statement(sheets[key], key, rows, m)
    write_dcf(sheets["DCF"], m, wb[xl_dcf.CALC_SHEET])
    write_inputs(sheets["INP"], m)                 # after the DCF: it places the settings the DCF declared
    write_checks(sheets["CHK"], m)
    write_notes(wb["Notes"], m)
    write_mapping(wb["Mapping"], fin)
    write_raw(wb["Raw"], fin.sources)
    write_summary(sheets["SUM"], m)
    wb[xl_dcf.CALC_SHEET].sheet_state = "hidden"
    for name, color in TAB_COLORS.items():
        if name in wb.sheetnames:
            wb[name].sheet_properties.tabColor = color
    m.reg.write(wb)
    wb.active = 0                                           # open on the summary
    return wb, m


# Lines the valuation depends on, with words to look for among the company's own tags when one is empty.
GAP_HINTS = {
    "Revenue": (["revenue"],), "Operating Income": (["operating", "income"],),
    "Pre-Tax Income": (["before", "tax"],), "Income Tax Expense": (["income", "tax", "expense"],),
    "Net Income": (["net", "income"], ["profit", "loss"]),
    "D&A": (["depreciation"], ["amortization"], ["amortisation"]),
    "Stock-Based Comp": (["sharebased"], ["share", "based", "compensation"]),
    "Capex": (["acquire", "property"], ["acquire", "productive"], ["purchase", "property"], ["capital", "expenditure"]),
    "Cash from Operations": (["operating", "activities"],), "Cash & Equivalents": (["cash", "equivalents"],),
    "Total Assets": (["assets"],), "Total Equity": (["equity"],),
    "Diluted Shares": (["diluted", "shares"], ["weighted", "average", "shares"]),
}


def gap_report(m, facts):
    """Print the valuation-critical lines that came back empty, each with the company's closest-matching tags,
    ready to paste into overrides/<TICKER>.json. Returns the number of gaps."""
    from .transform import find_tags
    fin = m.fin
    end = fin.cols[-1].date().isoformat()
    gaps = [(label, words) for label, words in GAP_HINTS.items() if not m.reported_anywhere(ROW_BY_LABEL[label])]
    if not fin.shares_out[0]:
        print("  ! Shares outstanding not found in the SEC data: type the share count into the Assumptions tab "
              "(Market data), otherwise there is no value per share.")
    if not gaps:
        return 0 if fin.shares_out[0] else 1
    print(f"  ! {len(gaps)} line(s) the valuation uses came back empty. Closest tags this company reports for "
          f"{end} (add the right one to overrides/{fin.ticker}.json):")
    for label, word_sets in gaps:
        is_shares = label == "Diluted Shares"
        units = {"shares"} if is_shares else {fin.currency}
        seen, hits = set(), []
        for words in word_sets:
            for h in find_tags(facts, words, 6, end=end, units=units, kind=MAP_BY_LABEL[label][1]):
                if h[1] not in seen and "Accumulated" not in h[1]:
                    seen.add(h[1])
                    hits.append(h)
        print(f"      {label}:" + ("" if hits else "  no similar tag found (the company may not report it)"))
        for _, tag, text, val, unit in hits[:4]:
            shown = f"{val:,.0f}" if is_shares else f"{val / 1e6:,.0f}M"
            print(f"          {tag}  = {shown}   ({text[:55]})")
        if hits:
            print(f'          format:  "{label}": ["TagName"]   or   [["TagA", "TagB"]] to add two together')
    return len(gaps)


def coverage_report(m):
    g = m.grid
    past = [c for c in g.cols if c.kind in ("hist", "ltm")]
    ltm = f" + LTM to {m.fin.ltm_end.date()}" if m.fin.ltm_end is not None else " (no 10-Q newer than the last 10-K)"
    print(f"\nCoverage across {len(g.hist)} fiscal years{ltm}:")
    for key, rows in STATEMENTS:
        print(f"[{SHEET_NAMES[key]}]")
        for row in rows:
            if not row.data:
                continue
            have = sum(m.has_data(row.label, c) for c in past)
            derived = sum(m.uses_fallback(row, c) for c in past)
            if have == len(past):
                flag = ""
            elif have + derived == len(past):
                flag = "  (gaps filled by formula)"
            elif have + derived == 0:
                flag = "  <-- not reported (blank); add tags in mapping.py if the company does report it"
            else:
                flag = "  <-- partial"
            print(f"  {row.label:<32}{have}/{len(past)}{flag}")
    print(f"Pulled {date.today()}")
