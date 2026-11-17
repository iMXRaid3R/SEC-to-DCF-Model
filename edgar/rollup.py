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
FIELDS = ["run", "ticker", "company", "currency", "price", "perp", "exit", "peers", "midpoint", "upside",
          "weighted", "upside_weighted", "w_bear", "w_base", "w_bull", "bear", "bull", "bear_perp", "bull_perp", "bear_exit", "bull_exit", "growth_model", "growth_needed_perp",
          "growth_needed_exit", "wacc", "terminal_growth", "scenario", "checks", "note"]


def _num(v):
    """A usable number, or None (blank cells, text, and Excel error codes all become None)."""
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return None
    return None if v < -2_000_000_000 else float(v)           # COM returns #N/A, #VALUE! etc. as huge negatives


def parse_summary(get):
    """Read one Summary tab. `get(row, col)` returns a cell value (1-based). Per-share values come back in
    dollars per US-listed share, so foreign filers are comparable with the price."""
    rows = {get(r, 1): r for r in range(1, 70) if isinstance(get(r, 1), str)}
    side = {get(r, 4): r for r in range(4, 11) if isinstance(get(r, 4), str)}
    rate = side.get("WACC", side.get("Cost of equity"))          # financials are valued at the cost of equity
    out = {"company": (get(1, 1) or "").rsplit(" - Summary", 1)[0],
           "price": _num(get(rows["Share price (US listing)"], 2)), "scenario": get(rows["Scenario"], 2),
           "wacc": _num(get(rate, 5)) if rate else None, "terminal_growth": _num(get(side["Terminal growth"], 5)),
           "checks": get(side["Model checks"], 5)}
    head = rows["Method"]
    usd_col = 6 if get(head, 6) == "$ per US-listed share" else 3
    out["currency"] = "USD" if usd_col == 3 else next(
        (k.split("(")[1].split(" ")[0] for k in rows if k.startswith("Share price (") and "per ordinary" in k), "?")
    for key, label in METHODS.items():
        label = label if label in rows else FIN_METHODS.get(key)
        out[key] = _num(get(rows[label], usd_col)) if label in rows else None
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
                                    "terminal_growth", "checks")}
    row.update(run=run, ticker=ticker, scenario=selected, midpoint=_mid(base["perp"], base["exit"]),
               growth_model=base["growth_model"][0], growth_needed_perp=base["growth_needed"][0],
               growth_needed_exit=base["growth_needed"][1])
    for name, key in (("Bear", "bear"), ("Bull", "bull")):
        s = by_scenario.get(name, {})
        row[key], row[f"{key}_perp"], row[f"{key}_exit"] = _mid(s.get("perp"), s.get("exit")), s.get("perp"), s.get("exit")
    price, mid = row["price"], row["midpoint"]
    row["upside"] = mid / price - 1 if price and mid is not None else None
    w = [x if isinstance(x, (int, float)) else None for x in (weights or (None, None, None))]
    row["w_bear"], row["w_base"], row["w_bull"] = w
    parts = (row["bear"], mid, row["bull"])
    if all(x is not None for x in w) and all(x is not None for x in parts) and abs(sum(w) - 1) < 0.001:
        row["weighted"] = sum(a * b for a, b in zip(w, parts))
        row["upside_weighted"] = row["weighted"] / price - 1 if price else None
    else:
        row["weighted"] = row["upside_weighted"] = None
    notes = []
    if not price:
        notes.append("no share price (live price blank: check the ticker's exchange prefix or type an override)")
    if mid is None:
        notes.append("no value per share (check Shares outstanding and the gap messages from the build)")
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
        watch = [summary.Cells(5, 2)]
        if _call(lambda: xl.Calculation) != -4105:                       # make sure calculation is automatic
            _call(lambda: setattr(xl, "Calculation", -4105))
        print(" prices", end="", flush=True)
        _settle(xl, watch, need_number=summary.Cells(5, 2))
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
    ("DCF: perpetuity", "perp", "usd_price", 13), ("DCF: exit multiple", "exit", "usd_price", 13),
    ("Peer multiples", "peers", "usd_price", 13), ("DCF midpoint", "midpoint", "usd_price", 13),
    ("Upside to midpoint", "upside", "pct", 12), ("Bear", "bear", "usd_price", 11), ("Bull", "bull", "usd_price", 11),
    ("Probability-weighted value", "weighted", "usd_price", 13), ("Upside to weighted", "upside_weighted", "pct", 12),
    ("Revenue growth in model", "growth_model", "pct", 13), ("Growth the price needs (exit method)", "growth_needed_exit", "pct", 15),
    ("WACC", "wacc", "pct", 8), ("Terminal growth", "terminal_growth", "pct", 9), ("Checks", "checks", None, 16),
    ("Notes", "note", None, 70)]


def write_book(folder, rows, history):
    wb = Workbook()
    ws = wb.active
    ws.title = "Portfolio"
    ws["A1"], ws["A1"].font = "Portfolio valuation roll-up", st.TITLE_FONT
    ws["A2"] = (f"Read from the model workbooks on {rows[0]['run'] if rows else ''}. $ per US-listed share. DCF "
                "midpoint = average of the two DCF methods under the Base scenario; Bear and Bull are the same "
                "midpoint under those scenarios; the weighted value uses each workbook's scenario probabilities "
                "(Inputs tab) and needs the full run, not --base-only. Static values: run the roll-up to refresh.")
    ws["A2"].font = st.NOTE_FONT
    for j, (h, _, _, width) in enumerate(COLUMNS, start=1):
        st.header_cell(ws.cell(row=4, column=j), h)
        ws.cell(row=4, column=j).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[ws.cell(row=4, column=j).column_letter].width = width
    ws.row_dimensions[4].height = 44
    ordered = sorted(rows, key=lambda r: (r["upside"] is None, -(r["upside"] or 0)))
    for i, r in enumerate(ordered, start=5):
        for j, (_, field, fmt, _) in enumerate(COLUMNS, start=1):
            c = ws.cell(row=i, column=j, value=r.get(field))
            if fmt:
                c.number_format = st.FMT[fmt]
        ws.cell(row=i, column=1).font = Font(bold=True)
        ws.cell(row=i, column=7).font = Font(bold=True)
    if ordered:
        rng = f"H5:H{4 + len(ordered)}"
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
    log = append_log(folder, rows)
    book = write_book(folder, rows, read_log(folder))
    print(f"Saved {book.name} ({len(rows)} companies) and added {len(rows)} rows to {log.name}")
