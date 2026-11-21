"""Comps: LTM (or latest FY) operating metrics per company, plus price-driven multiples. The tab is tall: one
row per metric and one column per company, then the peer median and mean."""
import time

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from . import xl_style as st
from .sec import EdgarError, load_company
from . import sec
from .transform import load_financials


def _metrics(fin):
    """Values for one company. Flows use LTM when a 10-Q is newer than the last 10-K, else the last FY;
    a line missing from that period stays blank rather than mixing periods."""
    use_ltm = fin.ltm_end is not None
    last = fin.cols[-1]

    def get(label):
        if use_ltm:
            return fin.ltm.get(label)
        return fin.values.get(label, {}).get(last)

    def fy(label, end):
        return fin.values.get(label, {}).get(end)

    rev, cogs, gp = get("Revenue"), get("Cost of Revenue"), get("Gross Profit")
    if gp is None and rev is not None and cogs is not None:
        gp = rev - cogs
    notes = []
    ebit, da, ni = get("Operating Income"), get("D&A"), get("Net Income")
    if da is None and use_ltm and fy("D&A", last) is not None:
        da = fy("D&A", last)          # some companies report D&A only in the 10-K; it moves slowly
        notes.append("EBITDA uses last-FY D&A")
    cfo, capex = get("Cash from Operations"), get("Capex") or 0
    ebitda = ebit + da if ebit is not None and da is not None else None
    fcf = cfo - capex if cfo is not None else None

    growth = None
    if len(fin.cols) >= 2 and fy("Revenue", fin.cols[-2]):
        growth = fy("Revenue", last) / fy("Revenue", fin.cols[-2]) - 1
    tax, pretax = fy("Income Tax Expense", last), fy("Pre-Tax Income", last)
    tax_rate = min(max(tax / pretax, 0), 0.5) if tax is not None and pretax else 0.21

    debt = sum(get(l) or 0 for l in ("Short-Term Borrowings", "Current Portion of LT Debt", "Long-Term Debt"))
    cash = sum(get(l) or 0 for l in ("Cash & Equivalents", "Short-Term Investments", "Long-Term Investments"))
    nci = get("Noncontrolling Interest") or 0
    equity = get("Total Equity")
    if equity is None and get("Shareholders' Equity (Parent)") is not None:
        equity = get("Shareholders' Equity (Parent)") + nci
    invested = (equity or 0) + debt - cash
    roic = ebit * (1 - tax_rate) / invested if ebit is not None and equity is not None and invested > 0 else None

    ratio = lambda a, b: a / b if a is not None and b else None
    return {
        "Ticker": fin.symbol or fin.ticker, "Key": fin.ticker, "Company": fin.name,
        "Period": f"LTM {fin.ltm_end:%b-%y}" if use_ltm else f"FY ending {last:%b-%y}",
        "Currency": fin.currency, "ADR ratio": fin.adr_ratio,
        "Revenue": rev, "Revenue growth (last FY)": growth,
        "Gross margin": ratio(gp, rev), "EBIT margin": ratio(ebit, rev), "EBITDA margin": ratio(ebitda, rev),
        "Net margin": ratio(ni, rev), "FCF margin": ratio(fcf, rev), "ROIC": roic,
        "EBITDA": ebitda, "Net income": ni, "Net debt": debt - cash, "NCI": nci, "Shares": fin.shares_out[0],
        "Notes": "; ".join(notes),
    }


# (header, key, number format, width)
VALUE_COLS = [
    ("Ticker", "Ticker", None, 12), ("Company", "Company", None, 30), ("Period", "Period", None, 16),
    ("Ccy", "Currency", None, 6),
    ("Revenue", "Revenue", "usd", 12), ("Rev growth", "Revenue growth (last FY)", "pct", 10),
    ("Gross margin", "Gross margin", "pct", 10), ("EBIT margin", "EBIT margin", "pct", 10),
    ("EBITDA margin", "EBITDA margin", "pct", 10), ("Net margin", "Net margin", "pct", 10),
    ("FCF margin", "FCF margin", "pct", 10), ("ROIC", "ROIC", "pct", 9),
    ("EBITDA", "EBITDA", "usd", 12), ("Net income", "Net income", "usd", 12), ("Net debt", "Net debt", "usd", 12),
    ("NCI", "NCI", "usd", 10), ("Shares (M)", "Shares", "shares", 11), ("Ord / listed sh", "ADR ratio", "ratio", 9),
]
# Formula columns: (header, formula template using {col} letters by header, format)
CALC_COLS = [
    ("Live close", st.stockhistory("{Ticker}", 2), "usd_price"),
    ("Price date", st.stockhistory("{Ticker}", 1), "date"),
    ("Price override", None, "usd_price"),                    # optional manual input
    ("Share price", '=IF({Price override}<>"",{Price override},{Live close})', "usd_price"),
    # Financials are in each company's reporting currency, so the dollar price is converted before it meets them.
    ("FX per USD", '=IF({Ccy}="USD",1,' + st.stockhistory('"USD:"&{Ccy}', 2)[1:] + ')', "fx"),
    ("Market cap", '=IFERROR(IF(N({Share price})=0,"",{Share price}*{FX per USD}/{Ord / listed sh}*{Shares (M)}),"")', "usd"),
    ("Enterprise value", '=IF({Market cap}="","",{Market cap}+{Net debt}+{NCI})', "usd"),
    ("EV / Revenue", '=IFERROR(IF({Revenue}>0,{Enterprise value}/{Revenue},""),"")', "x"),
    # A multiple on negative EBITDA or earnings is meaningless and would drag the peer median down: left blank.
    ("EV / EBITDA", '=IFERROR(IF(AND({EBITDA}>0,{Enterprise value}>0),{Enterprise value}/{EBITDA},""),"")', "x"),
    ("P / E", '=IFERROR(IF({Net income}>0,{Market cap}/{Net income},""),"")', "x"),
]
NOTES_COL = "Notes"


def peer_rows(tickers, email, refresh=False, skip=(), adr=None):
    """One _metrics() row per ticker that loads; the rest are reported and skipped."""
    rows = []
    for t in tickers:
        if t.upper() in skip:
            continue
        try:
            ticker, name, facts = load_company(t, email, refresh)
            fin = load_financials(facts, ticker, name)
            fin.symbol = sec.price_symbol(ticker, email, refresh)
            fin.adr_ratio = (adr or {}).get(ticker, 1)
            rows.append(_metrics(fin))
        except EdgarError as e:
            print(f"  skipped {t}: {e}")
        except Exception as e:                  # network trouble on one peer shouldn't stop the model
            print(f"  skipped {t}: {type(e).__name__}: {e}")
        if not sec.OFFLINE:
            time.sleep(0.2)     # stay well under the SEC's 10 requests/second
    return rows


def build_comps(tickers, email, refresh=False):
    rows = peer_rows(tickers, email, refresh)
    if not rows:
        raise EdgarError("No comps could be built.")
    wb = Workbook()
    write_comps(rows, wb.active)
    return wb


# The tab is tall: one row per metric, grouped, and one column per company, then the peer median and mean.
SECTIONS = [
    ("Company", ["Ticker", "Company", "Period", "Ccy"]),
    ("Operations", ["Revenue", "Rev growth", "Gross margin", "EBIT margin", "EBITDA margin", "Net margin",
                    "FCF margin", "ROIC"]),
    ("Earnings, debt and shares", ["EBITDA", "Net income", "Net debt", "NCI", "Shares (M)", "Ord / listed sh"]),
    ("Price and value", ["Live close", "Price date", "Price override", "Share price", "FX per USD", "Market cap",
                         "Enterprise value"]),
    ("Multiples", ["EV / Revenue", "EV / EBITDA", "P / E"]),
]
FIRST_COL = 2                                   # column B: the first company (the subject in a model workbook)


def write_comps(rows, ws, subject=None, price_ref=None, reg=None):
    """rows: list of _metrics() dicts. In a model workbook, `subject` is the company's own column: it comes
    first, takes its share price from `price_ref` (the Inputs tab), is left out of the peer median/mean, and
    gets an implied-value block from the peer medians. Returns {multiple header: median cell reference}."""
    if subject:
        rows = [subject] + rows
    ws.title = "Comps"
    ws["A1"], ws["A1"].font = ("Comparable Companies" if not subject
                               else f"{subject['Company']} ({subject['Key']}) - Comparable Companies"), st.TITLE_FONT
    ws["A2"] = ("One column per company. Prices are the latest close from Excel's STOCKHISTORY (Microsoft 365); "
                "type in the yellow Price override row to replace one. Flows are LTM where a newer 10-Q exists, else "
                "the latest fiscal year; millions, in each company's reporting currency (Ccy), with prices converted "
                "by the FX row.")
    ws["A2"].font = st.NOTE_FONT
    fmt_of = {h: f for h, _, f, _ in VALUE_COLS}
    fmt_of.update({h: f for h, _, f in CALC_COLS})
    key_of = {h: k for h, k, _, _ in VALUE_COLS}
    tpl_of = {h: t for h, t, _ in CALC_COLS}

    # rows of the tab
    hdr = 4
    row_of, r, section_rows = {}, hdr + 1, []
    for title, heads in SECTIONS:
        section_rows.append((r, title))
        r += 1
        for h in heads:
            row_of[h] = r
            r += 1
        r += 1
    notes_row = r - 1
    cols = [get_column_letter(FIRST_COL + i) for i in range(len(rows))]
    med_col, mean_col = get_column_letter(FIRST_COL + len(rows)), get_column_letter(FIRST_COL + len(rows) + 1)
    peers = cols[1:] if subject else cols                # the subject is not its own comparable
    last_col = FIRST_COL + len(rows) + 1

    # header row: tickers, then median and mean
    st.header_cell(ws.cell(row=hdr, column=1), "")
    for i, (col, row) in enumerate(zip(cols, rows)):
        st.header_cell(ws[f"{col}{hdr}"], row["Key"], "ltm" if subject and i == 0 else "fcst")
    for col, label in ((med_col, "Peer median" if subject else "Median"), (mean_col, "Peer mean" if subject else "Mean")):
        st.header_cell(ws[f"{col}{hdr}"], label, "label")
    for sr, title in section_rows:
        st.section(ws, sr, 1, last_col, title, size=11)

    for h, rr in row_of.items():
        a = ws.cell(row=rr, column=1, value=h)
        a.alignment = st.Alignment(indent=1)
        fmt = fmt_of.get(h)
        for i, (col, row) in enumerate(zip(cols, rows)):
            c = ws[f"{col}{rr}"]
            is_subject = bool(subject) and i == 0
            if h in key_of:                                         # a figure from the filings
                c.value = row[key_of[h]]
                if is_subject and h == "Ord / listed sh":
                    c.value = 1                 # the subject's price is already per ordinary share (Inputs tab)
                if fmt:
                    c.number_format, c.font = ("General" if fmt == "ratio" else st.FMT[fmt]), st.font(st.BLUE)
                if h == "Company":
                    c.alignment = st.Alignment(wrap_text=True, vertical="top")
                    c.font = st.font(bold=True)
            else:                                                   # price-driven
                c.number_format = st.FMT[fmt]
                tpl = tpl_of[h]
                if is_subject and h == "Share price":
                    c.value, c.number_format = f"={price_ref}", st.FMT["price"]
                    c.font = st.formula_font(c.value)
                elif is_subject and h == "FX per USD":
                    c.value = 1                 # ... and already in the reporting currency
                elif tpl is None:                                   # the price override
                    if not is_subject:
                        c.fill, c.font = st.INPUT_FILL, st.font(st.BLUE)
                        if reg is not None:
                            reg.add(ws, c, "Comps", row["Key"], h)
                else:
                    c.value = tpl.format(**{k: f"{col}{v}" for k, v in row_of.items()})
                    c.font = st.formula_font(c.value)
            if h != "Company":
                c.alignment = st.Alignment(horizontal="right")
        if fmt in ("pct", "x") and peers:                           # peer median and mean of ratios
            rng = f"{peers[0]}{rr}:{peers[-1]}{rr}"
            for col, fn in ((med_col, "MEDIAN"), (mean_col, "AVERAGE")):
                c = ws[f"{col}{rr}"]
                c.value = f'=IFERROR({fn}({rng}),"")'
                c.number_format, c.font = st.FMT[fmt], st.font(bold=True)
                c.fill = st.PatternFill("solid", fgColor=st.THEME["pale"])
    ws.row_dimensions[row_of["Company"]].height = 45      # long names wrap to three lines
    if subject:
        for rr in row_of.values():
            ws[f"{cols[0]}{rr}"].fill = st.PatternFill("solid", fgColor="F3F4F6")    # the subject, set apart
    st.row_rules(ws, list(row_of.values()), last_col)
    # notes (e.g. "EBITDA uses last-FY D&A") under each company
    ws.cell(row=notes_row, column=1, value="Notes").alignment = st.Alignment(indent=1)
    for col, row in zip(cols, rows):
        c = ws[f"{col}{notes_row}"]
        c.value, c.font, c.alignment = row["Notes"] or None, st.NOTE_FONT, st.Alignment(wrap_text=True, vertical="top")
    if any(row["Notes"] for row in rows):
        ws.row_dimensions[notes_row].height = 30

    medians = {h: f"Comps!${med_col}${row_of[h]}" for h in ("EV / Revenue", "EV / EBITDA", "P / E", "Rev growth",
                                                             "Gross margin", "EBIT margin", "EBITDA margin",
                                                             "Net margin", "FCF margin", "ROIC")} if peers else {}
    after = notes_row + 2
    if subject and peers:
        top = after
        end = _implied_block(ws, top, cols[0], med_col, row_of)
        medians["_implied"] = f"Comps!$F${top + 2}:$F${top + 4}"    # the three implied values per share
        after = end + 2
    ws.cell(row=after, column=1, value="If a live close is blank or wrong, check the exchange prefix on the ticker "
                                       "(XNAS: Nasdaq, XNYS: NYSE) or type the price in the Price override row."
            ).font = st.NOTE_FONT
    ws.column_dimensions["A"].width = 26
    for i in range(len(rows) + 2):
        ws.column_dimensions[get_column_letter(FIRST_COL + i)].width = 15
    for L in "FG":                                                  # the implied block's wider columns
        ws.column_dimensions[L].width = max(ws.column_dimensions[L].width or 0, 15)
    ws.freeze_panes = f"B{hdr + 1}"
    return medians


def _implied_block(ws, top, subj, med, row_of):
    """What the subject would be worth at the peer median multiples. Returns the last row used.
    subj / med: the subject's and the peer median's column letters."""
    L = lambda col, h: f"${col}${row_of[h]}"
    st.section(ws, top, 1, 7, "Implied value at peer median multiples", size=11)
    heads = ["Multiple", "Peer median", "Company metric", "Implied EV", "Implied equity", "Per share", "vs. price"]
    for i, h in enumerate(heads):
        st.header_cell(ws.cell(row=top + 1, column=1 + i), h, "fcst" if i else "label")
    bridge = f"-{L(subj, 'Net debt')}-{L(subj, 'NCI')}"
    lines = [("EV / EBITDA", "EBITDA", True), ("EV / Revenue", "Revenue", True), ("P / E", "Net income", False)]
    for k, (mult, metric, is_ev) in enumerate(lines):
        r = top + 2 + k
        ws.cell(row=r, column=1, value=mult).alignment = st.Alignment(indent=1)
        base = f"{L(med, mult)}*{L(subj, metric)}"
        cells = [(f"={L(med, mult)}", "x"), (f"={L(subj, metric)}", "usd"),
                 (f'=IFERROR({base},"")' if is_ev else None, "usd"),
                 (f'=IFERROR({base}{bridge if is_ev else ""},"")', "usd"),
                 (f'=IFERROR(E{r}/{L(subj, "Shares (M)")},"")', "price"),
                 (f'=IFERROR(F{r}/{L(subj, "Share price")}-1,"")', "pct")]
        for i, (f, fmt) in enumerate(cells, start=2):
            c = ws.cell(row=r, column=i, value=f)
            c.number_format = st.FMT[fmt]
            if f:
                c.font = st.formula_font(f)
    st.row_rules(ws, [top + 2 + k for k in range(len(lines))], 7)
    return top + 1 + len(lines)
