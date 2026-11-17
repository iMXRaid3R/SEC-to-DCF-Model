"""Comps: one row of LTM (or latest FY) operating metrics per ticker, plus price-driven multiples."""
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


def write_comps(rows, ws, subject=None, price_ref=None, reg=None):
    """rows: list of _metrics() dicts. In a model workbook, `subject` is the company's own row: it is listed
    first, takes its share price from `price_ref` (the Assumptions tab), is left out of the peer median/mean,
    and gets an implied-value block from the peer medians. Returns {multiple header: median cell reference}."""
    if subject:
        rows = [subject] + rows
    ws.title = "Comps"
    ws["A1"], ws["A1"].font = ("Comparable Companies" if not subject
                               else f"{subject['Company']} ({subject['Key']}) - Comparable Companies"), st.TITLE_FONT
    ws["A2"] = ("Prices are the latest close from Excel's STOCKHISTORY (Microsoft 365); type in the yellow column to "
                "override. Flows are LTM where a newer 10-Q exists, else the latest fiscal year. Millions, in "
                "each company's reporting currency (Ccy); prices are converted with the FX column.")
    ws["A2"].font = st.NOTE_FONT

    headers = [h for h, *_ in VALUE_COLS] + [h for h, *_ in CALC_COLS] + [NOTES_COL]
    letter = {h: get_column_letter(i + 1) for i, h in enumerate(headers)}
    for i, h in enumerate(headers, start=1):
        st.header_cell(ws.cell(row=4, column=i), h, "fcst" if len(VALUE_COLS) < i < len(headers) else "label")
    first, last = 5, 4 + len(rows)
    for r, row in enumerate(rows, start=first):
        for i, (h, key, fmt, _) in enumerate(VALUE_COLS, start=1):
            c = ws.cell(row=r, column=i, value=row[key])
            if subject and r == first and h == "Ord / listed sh":
                c.value = 1                     # the subject's price is already per ordinary share (Assumptions tab)
            if fmt:
                c.number_format, c.font = ("General" if fmt == "ratio" else st.FMT[fmt]), st.font(st.BLUE)
        for h, tpl, fmt in CALC_COLS:
            c = ws.cell(row=r, column=headers.index(h) + 1)
            c.number_format = st.FMT[fmt]
            if subject and r == first and h == "Share price":
                c.value, c.number_format = f"={price_ref}", st.FMT["price"]
                c.font = st.formula_font(c.value)
            elif subject and r == first and h == "FX per USD":
                c.value = 1                     # ... and already in the reporting currency
            elif tpl is None:
                c.fill, c.font = st.INPUT_FILL, st.font(st.BLUE)
                if reg is not None and not (subject and r == first):
                    reg.add(ws, c, "Comps", row["Key"], h)
            else:
                c.value = tpl.format(**{k: f"{v}{r}" for k, v in letter.items()})
                c.font = st.formula_font(c.value)
        ws.cell(row=r, column=len(headers), value=row["Notes"]).font = st.NOTE_FONT

    p_first = first + 1 if subject else first           # peers only: the subject is not its own comparable
    if subject:
        for i in (1, 2):
            ws.cell(row=first, column=i).font = st.font(bold=True)
    medians = {}
    for k, (label, fn) in enumerate((("Peer median" if subject else "Median", "MEDIAN"),
                                     ("Peer mean" if subject else "Mean", "AVERAGE")), start=last + 2):
        ws.cell(row=k, column=1, value=label).font = st.font(bold=True)
        for i, h in enumerate(headers, start=1):
            fmt = next((f for hh, _, f, *_ in VALUE_COLS if hh == h), None) or next(
                (f for hh, _, f in CALC_COLS if hh == h), None)
            if fmt in ("pct", "x") and last >= p_first:
                c = ws.cell(row=k, column=i, value=f'=IFERROR({fn}({letter[h]}{p_first}:{letter[h]}{last}),"")')
                c.number_format, c.font = st.FMT[fmt], st.font(bold=True)
                if fn == "MEDIAN":
                    medians[h] = f"Comps!${letter[h]}${k}"
    note_row = last + 5
    if subject and last >= p_first:
        note_row = _implied_block(ws, last + 5, first, last + 2, letter) + 2
        medians["_implied"] = f"Comps!$F${last + 7}:$F${last + 9}"     # the three implied values per share

    widths = [w for *_, w in VALUE_COLS] + [13] * len(CALC_COLS) + [30]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "D5"
    ws.cell(row=note_row, column=1, value="If a live close is blank or wrong, check the exchange prefix on the ticker "
                                          "(XNAS: Nasdaq, XNYS: NYSE) or use the override column.").font = st.NOTE_FONT
    return medians


def _implied_block(ws, top, subj, med, letter):
    """What the subject would be worth at the peer median multiples. Returns the last row used."""
    L = lambda h, r: f"${letter[h]}${r}"
    ws.cell(row=top, column=1, value="Implied value at peer median multiples").font = st.font(bold=True)
    heads = ["Multiple", "Peer median", "Company metric", "Implied EV", "Implied equity", "Per share", "vs. price"]
    for i, h in enumerate(heads):
        st.header_cell(ws.cell(row=top + 1, column=1 + i), h, "fcst" if i else "label")
    bridge = f"-{L('Net debt', subj)}-{L('NCI', subj)}"
    lines = [("EV / EBITDA", "EBITDA", True), ("EV / Revenue", "Revenue", True), ("P / E", "Net income", False)]
    for k, (mult, metric, is_ev) in enumerate(lines):
        r = top + 2 + k
        ws.cell(row=r, column=1, value=mult)
        base = f"{L(mult, med)}*{L(metric, subj)}"
        cells = [(f"={L(mult, med)}", "x"), (f"={L(metric, subj)}", "usd"),
                 (f'=IFERROR({base},"")' if is_ev else None, "usd"),
                 (f'=IFERROR({base}{bridge if is_ev else ""},"")', "usd"),
                 (f'=IFERROR(E{r}/{L("Shares (M)", subj)},"")', "price"),
                 (f'=IFERROR(F{r}/{L("Share price", subj)}-1,"")', "pct")]
        for i, (f, fmt) in enumerate(cells, start=2):
            c = ws.cell(row=r, column=i, value=f)
            c.number_format = st.FMT[fmt]
            if f:
                c.font = st.formula_font(f)
    return top + 1 + len(lines)
