"""Keep the user's inputs when a workbook is rebuilt.

Every input cell is registered while the workbook is built (a key that describes it, its address and its
generated default) and the list is saved in a hidden "_inputs" sheet. On the next build the old file is read:
any registered cell whose value no longer equals its default was changed by the user, and is copied into the
cell with the same key in the new workbook - so inputs survive new filings, extra years and layout changes.
Workbooks made before this existed are read by their row labels instead (legacy_inputs)."""
import json
import re
import shutil
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

SHEET = "_inputs"
# input-key prefix -> the tab whose absence means "part switched off" (DCF and beta settings sit on Inputs but
# belong to the DCF part)
PART_TABS = {"DCF": "DCF", "Market": "DCF", "Multiples": "Multiples", "Comps": "Comps", "Build": "Revenue Build"}
# Summary weights belong to the part that produces each method's value
WEIGHT_TABS = {"DCF: perpetuity growth": "DCF", "DCF: exit multiple": "DCF", "Peer multiples": "Comps",
               "Forward P/E": "Multiples", "PEG": "Multiples", "Forward EV/EBITDA": "Multiples",
               "Forward EV/Sales": "Multiples", "Price / free cash flow": "Multiples"}
PARKED = "(parked)"     # inputs from a tab left out of a build, kept in _inputs until the tab is built again


class Registry:
    def __init__(self):
        self.items = {}          # key -> (sheet title, cell address, default value)

    def add(self, ws, cell, *key, mark=True):
        """Register an input cell and give it the input look (Notes lines are free text and stay plain;
        mark=False for a formula cell the user may type over, as in the financials template)."""
        k = "|".join(str(p) for p in key)
        assert k not in self.items, f"duplicate input key {k}"
        self.items[k] = (ws.title, cell.coordinate, cell.value)
        if mark and key[0] != "Notes":
            from .xl_style import mark_input
            mark_input(cell)
        return cell

    def write(self, wb):
        ws = wb.create_sheet(SHEET)
        ws.append(["key", "sheet", "cell", "default (as generated)"])
        for k, (sheet, addr, default) in self.items.items():
            ws.append([k, sheet, addr, json.dumps(default, default=str)])
        ws.sheet_state = "hidden"


def _norm(v):
    """Excel rewrites formulas slightly when it saves: it drops quotes around simple sheet names, the _xlfn.
    prefixes and spaces, and writes 20.0 as 20."""
    if isinstance(v, str) and v.startswith("="):
        v = v.replace(" ", "").replace("'", "").replace("_xlfn.", "").upper()
        return re.sub(r"(?<![\w.])(\d+)\.0+(?![\d])", r"\1", v)
    return v


def _same(a, b):
    if isinstance(a, bool) or isinstance(b, bool):
        return a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= 1e-12 * max(1.0, abs(a), abs(b))
    return _norm(a) == _norm(b)


def _is_formula(v):
    return isinstance(v, str) and v.startswith("=")


def _usable(v):
    return isinstance(v, (int, float, str, date, datetime)) and v != ""


def _registered_edits(old):
    """key -> (value, old sheet, old address) for registered cells the user changed."""
    out = {}
    for key, sheet, addr, default in old[SHEET].iter_rows(min_row=2, max_col=4, values_only=True):
        if sheet == PARKED:
            out[key] = (json.loads(default), PARKED, "")
            continue
        if sheet not in old.sheetnames:
            continue
        v = old[sheet][addr].value
        if _usable(v) and not _same(v, json.loads(default)):
            out[key] = (v, sheet, addr)
    return out


_YEAR_TABS = ("Income Statement", "Equity Model")     # tabs with period-end dates in row 3 over FY headers in row 4


def _fy_by_end(wb):
    """Period end (date) -> the year in its FY....A header, read from the first statement tab."""
    for name in _YEAR_TABS:
        if name in wb.sheetnames:
            ws, out = wb[name], {}
            for c in range(2, ws.max_column + 1):
                end, head = ws.cell(row=3, column=c).value, ws.cell(row=4, column=c).value
                m = re.fullmatch(r"FY(\d{4})A", str(head or ""))
                if m and isinstance(end, (date, datetime)):
                    out[end.date() if isinstance(end, datetime) else end] = int(m.group(1))
            return out
    return {}


def year_shift(old, new):
    """How many years the fiscal-year labels moved between two builds of the same company (1 for a workbook
    that called NVIDIA's year ending January 2026 FY2025 and is now rebuilt as FY2026), else 0."""
    a, b = _fy_by_end(old), _fy_by_end(new)
    shifts = {b[e] - a[e] for e in a.keys() & b.keys()}
    return shifts.pop() if len(shifts) == 1 else 0


def _shift_key(key, n):
    return re.sub(r"FY(\d{4})([AE])", lambda m: f"FY{int(m.group(1)) + n}{m.group(2)}", key) if n else key


LEGACY_DCF = ("Risk-free rate", "Equity risk premium", "Beta", "Pre-tax cost of debt", "Terminal growth",
              "Exit EV/EBITDA multiple", "Mid-year convention (1 = on)")
LEGACY_DCF_DEFAULTS = (0.0425, 0.055, 1.0, 0.05, 0.025, 15.0, 1)


def _legacy_edits(old):
    """Inputs from a workbook built before the _inputs sheet existed: typed numbers found by row label."""
    out = {}
    if "DCF" in old.sheetnames:
        ws = old["DCF"]
        for r in range(4, 30):
            label, v = ws.cell(row=r, column=1).value, ws.cell(row=r, column=2).value
            if label in LEGACY_DCF and isinstance(v, (int, float)) \
                    and not _same(v, LEGACY_DCF_DEFAULTS[LEGACY_DCF.index(label)]):
                out[f"DCF|{label}"] = (v, "DCF", f"B{r}")
    if "Assumptions" in old.sheetnames:
        ws = old["Assumptions"]
        headers = {j: ws.cell(row=4, column=j).value for j in range(2, ws.max_column + 1)}
        for r in range(5, ws.max_row + 1):
            label = ws.cell(row=r, column=1).value
            if label == "Price override ($)" and isinstance(ws.cell(row=r, column=2).value, (int, float)):
                out["Market|Price override"] = (ws.cell(row=r, column=2).value, ws.title, f"B{r}")
            if label == "Ticker" and ":" in str(ws.cell(row=r, column=2).value):
                out["Market|Ticker"] = (ws.cell(row=r, column=2).value, ws.title, f"B{r}")   # e.g. XNAS:ISRG
            if not isinstance(label, str) or label == "Market data":
                continue
            for j, h in headers.items():
                v = ws.cell(row=r, column=j).value
                if isinstance(h, str) and h.endswith("E") and isinstance(v, (int, float)) and v != 0:
                    out[f"Live|{label}|{h}"] = (v, ws.title, ws.cell(row=r, column=j).coordinate)
    return out


def carry_over(old_path, wb, reg, reset=()):
    """Copy the user's inputs from the previous workbook into the new one. Inputs whose key contains one of
    the `reset` words (any case) are left at their new defaults instead. Inputs on a tab this build leaves out
    (a part switched off) are parked in the new _inputs sheet and come back when the tab is built again.
    Returns (kept, dropped, parked): lists of (key, value), (key, reason) and (key, value)."""
    try:
        old = load_workbook(old_path)
    except Exception as e:                                   # unreadable file: build fresh, say so
        return [], [("(whole file)", f"could not read {Path(old_path).name}: {e}")], []
    edits = _registered_edits(old) if SHEET in old.sheetnames else _legacy_edits(old)
    shift = year_shift(old, wb)          # year labels renamed: an override typed for a year follows that year
    edits = {_shift_key(k, shift): v for k, v in edits.items()}
    if shift:
        print(f"  Fiscal years are now labelled as the company names them ({shift:+d} year); "
              "inputs typed for a year moved with it")
    kept, dropped, parked = [], [], []
    for key, (v, old_sheet, old_addr) in edits.items():
        if any(w.lower() in key.lower() for w in reset):
            dropped.append((key, f"reset to its default as asked (was {v})"))
            continue
        if key not in reg.items:
            # an input of a part this build leaves out waits in _inputs; anything else (a forecast year that is
            # now actual, say) is gone for good
            parts_of_key = key.split("|")
            tab = (WEIGHT_TABS.get(parts_of_key[-1]) if parts_of_key[:2] == ["Summary", "Weight"]
                   else PART_TABS.get(parts_of_key[0]))
            if tab and tab not in wb.sheetnames and not _is_formula(v) and SHEET in wb.sheetnames:
                wb[SHEET].append([key, PARKED, "", json.dumps(v, default=str)])
                parked.append((key, v))
            else:
                dropped.append((key, "no matching cell in the new layout (e.g. that forecast year is now actual)"))
            continue
        sheet, addr, default = reg.items[key]
        if _same(v, default):
            continue
        if _is_formula(v) and (sheet, addr) != (old_sheet, old_addr):
            dropped.append((key, f"your formula {v} was not moved because the layout changed - re-enter it"))
            continue
        wb[sheet][addr].value = v
        kept.append((key, v))
    return kept, dropped, parked


def backup(path):
    """Copy an existing workbook to backups/<name>_<timestamp>.xlsx before it is overwritten."""
    path = Path(path)
    if not path.exists():
        return None
    folder = path.resolve().parent / "backups"
    folder.mkdir(exist_ok=True)
    dest = folder / f"{path.stem}_{datetime.now():%Y-%m-%d_%H%M%S}{path.suffix}"
    shutil.copy2(path, dest)
    return dest
