"""Portfolio roll-up and valuation log.

The model's values only exist once Excel has calculated the workbook (live prices included), so this opens each
<TICKER>_financials.xlsx in a hidden, read-only Excel, reads the Summary tab under the Bear, Base and Bull
scenarios, and writes:
    valuation_log.csv   one row per company per run (appended: a history of your fair values)
    Portfolio.xlsx      the latest run side by side, plus the full history
Nothing is saved back into the model workbooks. Needs Windows, Excel and pywin32 (the ./run launcher adds it)."""
import csv
import time
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font

from . import xl_style as st
from .layout import INPUT_ROWS, SCENARIOS, SHEET_NAMES

LOG = "valuation_log.csv"
BOOK = "Portfolio.xlsx"
METHODS = {"perp": "DCF: perpetuity growth", "exit": "DCF: exit multiple", "peers": "Peer multiples"}
FIN_METHODS = {"perp": "Equity model: residual income", "exit": "Equity model: exit P/B"}   # same two slots
TARGETS = {"target": "Price target (weighted)", "target_12m": "12-month price target"}
MULT_METHODS = {"fwd_pe": "Forward P/E", "peg": "PEG", "ev_ebitda_fwd": "Forward EV/EBITDA",
                "ev_sales_fwd": "Forward EV/Sales", "p_fcf": "Price / free cash flow"}
FIELDS = ["run", "ticker", "company", "currency", "price", "perp", "exit", "peers", "midpoint", "upside",
          "weighted", "upside_weighted", "w_bear", "w_base", "w_bull", "bear", "bull", "bear_perp", "bull_perp", "bear_exit", "bull_exit", "growth_model", "growth_needed_perp",
          "growth_needed_exit", "wacc", "terminal_growth", "scenario", "checks", "note"] + list(MULT_METHODS) + [
          "mult_avg", "target", "target_12m", "upside_target", "target_weighted", "bear_target", "bull_target"]


def _num(v):
    """A usable number, or None (blank cells, text, and Excel error codes all become None)."""
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return None
    return None if v < -2_000_000_000 else float(v)           # COM returns #N/A, #VALUE! etc. as huge negatives


def parse_summary(get):
    """Read one Summary tab. `get(row, col)` returns a cell value (1-based). Per-share values come back in
    dollars per US-listed share, so foreign filers are comparable with the price."""
    rows = {}
    for r in range(1, 90):                                       # first occurrence of each label in column A
        v = get(r, 1)
        if isinstance(v, str) and v not in rows:
            rows[v] = r
    side = {get(r, 4): r for r in range(4, 20) if isinstance(get(r, 4), str)}   # the valuation-inputs panel
    rate = side.get("WACC", side.get("Cost of equity"))          # financials are valued at the cost of equity
    at = lambda r, c: get(r, c) if r else None
    beside = lambda r: next((v for v in (get(r, 5), get(r, 6)) if v is not None), None) if r else None
    out = {"company": (get(1, 1) or "").rsplit(" - Summary", 1)[0],
           "price": _num(at(rows.get("Share price (US listing)"), 2)), "scenario": at(rows.get("Scenario"), 2),
           "wacc": _num(beside(rate)), "terminal_growth": _num(beside(side.get("Terminal growth"))),
           "checks": beside(side.get("Model checks"))}
    head = rows.get("Method")                    # no method table in a statements-only workbook
    usd_col = 6 if head and get(head, 6) == "$ per US-listed share" else 3
    foreign = bool(head) and usd_col == 6 or any(k.startswith("Share price (") and "per ordinary" in k for k in rows)
    out["currency"] = "USD" if not foreign else next(
        (k.split("(")[1].split(" ")[0] for k in rows if k.startswith("Share price (") and "per ordinary" in k), "?")
    for key, label in list(METHODS.items()) + list(MULT_METHODS.items()) + list(TARGETS.items()):
        label = label if label in rows else FIN_METHODS.get(key)
        out[key] = _num(get(rows[label], usd_col)) if head and label in rows and rows[label] > head else None
    out["financial"] = FIN_METHODS["perp"] in rows
    for key, label in (("growth_model", "Revenue growth per year in the model"),
                       ("growth_needed", "Revenue growth per year the price needs")):
        r = rows.get(label)
        out[key] = (_num(get(r, 2)), _num(get(r, 3))) if r else (None, None)
    return out


def _mid(a, b):
    vals = [v for v in (a, b) if v is not None]
    return sum(vals) / len(vals) if vals else None


def to_row(ticker, by_scenario, selected, run, weights=None):
    """One log row from {scenario: parse_summary()}; `selected` is the scenario the workbook was saved on and
    `weights` the (Bear, Base, Bull) probabilities from its Inputs tab."""
    base = by_scenario.get("Base") or by_scenario[selected]
    row = {k: base.get(k) for k in ("company", "currency", "price", "perp", "exit", "peers", "wacc",
                                    "terminal_growth", "checks") + tuple(MULT_METHODS)}
    row.update({k: base.get(k) for k in TARGETS})
    row["bear_target"] = by_scenario.get("Bear", {}).get("target")
    row["bull_target"] = by_scenario.get("Bull", {}).get("target")
    mults = [v for v in (base.get(k) for k in MULT_METHODS) if v is not None]
    row["mult_avg"] = sum(mults) / len(mults) if mults else None
    row.update(run=run, ticker=ticker, scenario=selected, midpoint=_mid(base.get("perp"), base.get("exit")),
               growth_model=base.get("growth_model", (None, None))[0],
               growth_needed_perp=base.get("growth_needed", (None, None))[0],
               growth_needed_exit=base.get("growth_needed", (None, None))[1])
    for name, key in (("Bear", "bear"), ("Bull", "bull")):
        s = by_scenario.get(name, {})
        row[key], row[f"{key}_perp"], row[f"{key}_exit"] = _mid(s.get("perp"), s.get("exit")), s.get("perp"), s.get("exit")
    price, mid = row["price"], row["midpoint"]
    row["upside"] = mid / price - 1 if price and mid is not None else None
    row["upside_target"] = row["target"] / price - 1 if price and row["target"] is not None else None
    w = [x if isinstance(x, (int, float)) else None for x in (weights or (None, None, None))]
    row["w_bear"], row["w_base"], row["w_bull"] = w
    ok = all(x is not None for x in w) and abs(sum(x or 0 for x in w) - 1) < 0.001
    parts = (row["bear"], mid, row["bull"])
    if ok and all(x is not None for x in parts):
        row["weighted"] = sum(a * b for a, b in zip(w, parts))
        row["upside_weighted"] = row["weighted"] / price - 1 if price else None
    else:
        row["weighted"] = row["upside_weighted"] = None
    targets = (row["bear_target"], row["target"], row["bull_target"])
    row["target_weighted"] = sum(a * b for a, b in zip(w, targets)) if ok and None not in targets else None
    notes = []
    if not price:
        notes.append("no share price (live price blank: check the ticker's exchange prefix or type an override)")
    if mid is None and row["mult_avg"] is None:
        notes.append("no value per share (check Shares outstanding and the gap messages from the build)")
    elif mid is None:
        notes.append("built without the DCF: see the multiples columns")
    if base.get("financial"):
        notes.append("financials template: the two values are residual income and exit P/B, at the cost of equity")
    if row["currency"] != "USD":
        notes.append(f"reports in {row['currency']}; values converted to $ per US-listed share")
    if isinstance(row["checks"], str) and "OK" not in row["checks"]:
        notes.append("model checks flag a difference: see the Checks & Metrics tab")
    row["note"] = "; ".join(notes)
    return row


# ---------------------------------------------------------------------------------------------------------
# Excel (COM) side
# ---------------------------------------------------------------------------------------------------------
PENDING = ("#BUSY!", "#GETTING_DATA", "#CONNECT!")
REJECTED = (-2147418111, -2147417846)        # Excel is busy: retry


def _call(fn, timeout=30):
    import pywintypes
    deadline = time.time() + timeout
    while True:
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.hresult not in REJECTED or time.time() > deadline:
                raise
            time.sleep(0.25)


def _settle(xl, watch, timeout=40, need_number=None, patience=25):
    """Wait until Excel has finished calculating and the live prices have arrived. The price formulas hide
    "still loading" behind IFERROR, so `need_number` (the price cell) is also polled until it holds a number,
    for up to `patience` seconds - after that the price is treated as unavailable."""
    deadline = time.time() + timeout
    time.sleep(0.3)
    if need_number is not None:
        give_up = time.time() + patience
        while time.time() < give_up and _num(_call(lambda: need_number.Value)) is None:
            time.sleep(1)
    while time.time() < deadline:
        busy = _call(lambda: xl.CalculationState) != 0
        pending = any(str(_call(lambda c=c: c.Text)) in PENDING for c in watch)
        if not busy and not pending:
            return True
        time.sleep(0.5)
    return False


def _block(wb, sheet, rng):
    """A range's values as a list of rows (one COM call), with Excel error codes turned into None."""
    v = _call(lambda: wb.Worksheets(sheet).Range(rng).Value)
    rows = v if isinstance(v, tuple) and v and isinstance(v[0], tuple) else ((v,),)
    return [[None if isinstance(x, int) and not isinstance(x, bool) and x < -2_000_000_000 else x for x in row]
            for row in rows]


def read_workbook(xl, path, scenarios=SCENARIOS, extra=None):
    """{scenario: parse_summary()} and the scenario the file was saved on. Opens read-only; saves nothing.
    Excel calculates the workbook as it opens and again each time the scenario cell changes, so no extra
    recalculation is asked for: the scenario the file was saved on is read first, then only the others."""
    wb = _call(lambda: xl.Workbooks.Open(str(path), 0, True))            # UpdateLinks=0, ReadOnly=True
    try:
        names = [_call(lambda i=i: wb.Worksheets(i).Name) for i in range(1, _call(lambda: wb.Worksheets.Count) + 1)]
        if SHEET_NAMES["INP"] not in names:
            raise RuntimeError("built with an older layout: rebuild it first (./run TICKER)")
        summary, a = wb.Worksheets("Summary"), wb.Worksheets(SHEET_NAMES["INP"])
        get = lambda r, c: _call(lambda: summary.Cells(r, c).Value)
        col_a = [row[0] for row in (_call(lambda: summary.Range("A1:A40").Value) or ())]
        price_row = next((i + 1 for i, v in enumerate(col_a) if v == "Share price (US listing)"), 5)
        watch = [summary.Cells(price_row, 2)]
        if _call(lambda: xl.Calculation) != -4105:                       # make sure calculation is automatic
            _call(lambda: setattr(xl, "Calculation", -4105))
        print(" prices", end="", flush=True)
        _settle(xl, watch, need_number=summary.Cells(price_row, 2))
        cell = a.Range(f"B{INPUT_ROWS['Scenario']}")
        weights = tuple(_num(_call(lambda k=k: a.Range(f"B{INPUT_ROWS[k]}").Value)) for k in ("w_bear", "w_base", "w_bull"))
        selected = _call(lambda: cell.Value)
        selected = selected if selected in SCENARIOS else "Base"
        out = {selected: parse_summary(get)}
        for name in scenarios:
            if name in out:
                continue
            print(f" {name}", end="", flush=True)
            _call(lambda n=name: setattr(cell, "Value", n))
            _settle(xl, watch)
            out[name] = parse_summary(get)
        if extra is not None:                    # the write-up also wants whole blocks, read under Base
            if _call(lambda: cell.Value) != "Base":
                _call(lambda: setattr(cell, "Value", "Base"))
                _settle(xl, watch)
            out["_extra"] = extra(lambda sheet, rng: _block(wb, sheet, rng))
        return out, selected, weights
    finally:
        _call(lambda: wb.Close(SaveChanges=False))     # named argument: a bare Close(False) can misfire on the
                                                         # first workbook of a session ("'bool' is not callable")


def collect(paths, scenarios=SCENARIOS):
    """Rows for the log, one per workbook. Raises a clear error when Excel can't be driven from here."""
    try:
        import win32com.client
    except ImportError:
        raise RuntimeError("The roll-up reads values through Excel and needs the pywin32 package on Windows. "
                           "Use the launcher (./run --rollup) or add  --with pywin32  to your uv command.")
    run = datetime.now().strftime("%Y-%m-%d %H:%M")
    xl = win32com.client.DispatchEx("Excel.Application")      # a private instance, never your open window
    xl.Visible, xl.DisplayAlerts = False, False
    rows = []
    try:
        for path in paths:
            ticker = path.stem.replace("_financials", "")
            print(f"  {ticker}: opening", end="", flush=True)
            started = time.time()
            try:
                by_scenario, selected, weights = read_workbook(xl, path, scenarios)
                rows.append(to_row(ticker, by_scenario, selected, run, weights))
                r = rows[-1]
                print((f" -> price {r['price']:,.2f}, DCF midpoint {r['midpoint']:,.2f}" if r["price"] and r["midpoint"] is not None
                       else f" -> incomplete ({r['note']})") + f"   [{time.time() - started:.0f}s]")
            except Exception as e:
                print(f" skipped: {type(e).__name__}: {str(e)[:120]}")
    finally:
        xl.Quit()
    return rows


# ---------------------------------------------------------------------------------------------------------
# Outputs
# ---------------------------------------------------------------------------------------------------------
def append_log(folder, rows):
    path = Path(folder) / LOG
    new = not path.exists()
    if not new:                                   # an older log with fewer columns: rewrite it with today's
        with open(path, newline="", encoding="utf-8") as fh:
            reader = csv.DictReader(fh)
            old_fields, old_rows = reader.fieldnames, list(reader)
        if old_fields != FIELDS:
            with open(path, "w", newline="", encoding="utf-8") as fh:
                w = csv.DictWriter(fh, fieldnames=FIELDS)
                w.writeheader()
                w.writerows({k: r.get(k, "") for k in FIELDS} for r in old_rows)
    with open(path, "a", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=FIELDS)
        if new:
            w.writeheader()
        w.writerows({k: ("" if r.get(k) is None else r.get(k)) for k in FIELDS} for r in rows)
    return path


def read_log(folder):
    path = Path(folder) / LOG
    if not path.exists():
        return []
    with open(path, newline="", encoding="utf-8") as fh:
        out = []
        for r in csv.DictReader(fh):
            for k, v in r.items():
                if k not in ("run", "ticker", "company", "currency", "scenario", "checks", "note"):
                    try:
                        r[k] = float(v) if v != "" else None
                    except (TypeError, ValueError):
                        r[k] = None
            out.append(r)
        return out


COLUMNS = [  # (header, field, format, width)
    ("Ticker", "ticker", None, 9), ("Company", "company", None, 34), ("Price", "price", "usd_price", 11),
    ("Price target", "target", "usd_price", 12), ("Upside to target", "upside_target", "pct", 11),
    ("12-month target", "target_12m", "usd_price", 12),
    ("Target, probability-weighted", "target_weighted", "usd_price", 13),
    ("DCF: perpetuity", "perp", "usd_price", 13), ("DCF: exit multiple", "exit", "usd_price", 13),
    ("Peer multiples", "peers", "usd_price", 13), ("DCF midpoint", "midpoint", "usd_price", 13),
    ("Upside to midpoint", "upside", "pct", 12), ("Bear", "bear", "usd_price", 11), ("Bull", "bull", "usd_price", 11),
    ("Probability-weighted value", "weighted", "usd_price", 13), ("Upside to weighted", "upside_weighted", "pct", 12),
    ("Revenue growth in model", "growth_model", "pct", 13), ("Growth the price needs (exit method)", "growth_needed_exit", "pct", 15),
    ("Forward P/E", "fwd_pe", "usd_price", 11), ("PEG", "peg", "usd_price", 11),
    ("Average of multiples", "mult_avg", "usd_price", 12), ("WACC", "wacc", "pct", 8), ("Terminal growth", "terminal_growth", "pct", 9),
    ("DCF midpoint at the previous run", "prev_midpoint", "usd_price", 13), ("Change in midpoint", "midpoint_change", "pct", 11),
    ("Previous run", "prev_run", None, 16), ("Checks", "checks", None, 16),
    ("Notes", "note", None, 70)]


def write_book(folder, rows, history):
    wb = Workbook()
    ws = wb.active
    ws.title = "Portfolio"
    ws["A1"], ws["A1"].font = "Portfolio valuation roll-up", st.TITLE_FONT
    ws["A2"] = (f"Read from the model workbooks on {rows[0]['run'] if rows else ''}. $ per US-listed share. Price "
                "target = your weighted blend of the methods (Summary tab), Base scenario; the probability-weighted "
                "target blends the Bear / Base / Bull targets. DCF "
                "midpoint = average of the two DCF methods under the Base scenario; Bear and Bull are the same "
                "midpoint under those scenarios; the weighted value uses each workbook's scenario probabilities "
                "(Inputs tab) and needs the full run, not --base-only. Static values: run the roll-up to refresh.")
    ws["A2"].font = st.NOTE_FONT
    for j, (h, _, _, width) in enumerate(COLUMNS, start=1):
        st.header_cell(ws.cell(row=4, column=j), h)
        ws.cell(row=4, column=j).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[ws.cell(row=4, column=j).column_letter].width = width
    ws.row_dimensions[4].height = 44
    up = lambda r: r.get("upside_target") if r.get("upside_target") is not None else r.get("upside")
    ordered = sorted(rows, key=lambda r: (up(r) is None, -(up(r) or 0)))       # most upside to the target first
    for i, r in enumerate(ordered, start=5):
        for j, (_, field, fmt, _) in enumerate(COLUMNS, start=1):
            c = ws.cell(row=i, column=j, value=r.get(field))
            if fmt:
                c.number_format = st.FMT[fmt]
        ws.cell(row=i, column=1).font = Font(bold=True)
        ws.cell(row=i, column=4).font = Font(bold=True)
    if ordered:
        rng = f"E5:E{4 + len(ordered)}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0"], font=Font(color="006100")))
        ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], font=Font(color="9C0006")))
    ws.freeze_panes = "C5"

    hs = wb.create_sheet("History")
    hs.append(FIELDS)
    for c in hs[1]:
        c.fill, c.font = st.FILL["label"], st.HDR_FONT
    for r in history:
        hs.append([r.get(k) for k in FIELDS])
    hs.freeze_panes = "C2"
    path = Path(folder) / BOOK
    wb.save(path)
    return path


def value_changes(rows, history):
    """Set prev_midpoint / midpoint_change / prev_run on each row from the latest earlier log entry of the same
    ticker, and return printable lines for the values that moved by more than half a percent."""
    last = {}
    for h in history:                                   # the log is in run order: the last one seen wins
        if h.get("midpoint") is not None:
            last[h["ticker"]] = h
    lines = []
    for r in rows:
        h = last.get(r["ticker"])
        if not h or r.get("midpoint") is None or not h["midpoint"]:
            continue
        r["prev_midpoint"], r["prev_run"] = h["midpoint"], h["run"]
        r["midpoint_change"] = r["midpoint"] / h["midpoint"] - 1
        if abs(r["midpoint_change"]) >= 0.005:
            price = ""
            if r.get("price") and h.get("price"):
                price = f"; share price {r['price'] / h['price'] - 1:+.1%}"
            lines.append(f"  {r['ticker']:<6} DCF midpoint ${r['midpoint']:,.2f}, was ${h['midpoint']:,.2f} on "
                         f"{h['run']} ({r['midpoint_change']:+.1%}){price}")
    return lines


def rollup(folder, tickers=None, base_only=False):
    folder = Path(folder)
    paths = sorted(folder.glob("*_financials.xlsx"))
    if tickers:
        want = {t.upper() for t in tickers}
        missing = sorted(want - {p.stem.replace("_financials", "").upper() for p in paths})
        paths = [p for p in paths if p.stem.replace("_financials", "").upper() in want]
        if missing:
            print(f"  no workbook yet for: {', '.join(missing)} (build them first)")
    if not paths:
        raise RuntimeError("No <TICKER>_financials.xlsx workbooks found to roll up.")
    rows = collect(paths, ("Base",) if base_only else SCENARIOS)
    if not rows:
        raise RuntimeError("No workbook could be read.")
    moved = value_changes(rows, read_log(folder))
    log = append_log(folder, rows)
    book = write_book(folder, rows, read_log(folder))
    print(f"Saved {book.name} ({len(rows)} companies) and added {len(rows)} rows to {log.name}")
    if moved:
        print("Values that moved since each company's previous roll-up:")
        print("\n".join(moved))
