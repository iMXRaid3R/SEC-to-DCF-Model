"""Recalculate a saved workbook in real Excel (Windows COM) and read results back.
Used by the tests and handy for manual checks:  python tests/xl_com.py ISRG_financials.xlsx [price]"""
import sys
import time
from contextlib import contextmanager

import pywintypes
import win32com.client

from edgar.xl_dcf import R, SENS_CENTER, VAL_HDR

PENDING = ("#BUSY!", "#GETTING_DATA")
REJECTED = (-2147418111, -2147417846)      # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER: Excel is busy
XL_DONE = 0                                # Application.CalculationState when idle


def call(fn, timeout=30):
    """Run a COM call, retrying while Excel rejects calls because it is busy."""
    deadline = time.time() + timeout
    while True:
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.hresult not in REJECTED or time.time() > deadline:
                raise
            time.sleep(0.25)


def recalc(wb, watch=(), timeout=60):
    """Recalculate once, then wait until Excel is idle and STOCKHISTORY cells in `watch` have downloaded.
    (Recalculating again while waiting would restart the download.)"""
    xl = wb.Application
    call(xl.CalculateFull)
    deadline = time.time() + timeout
    while time.time() < deadline:
        busy = call(lambda: xl.CalculationState) != XL_DONE
        pending = any(call(lambda r=r: r.Text) in PENDING for r in watch)
        if not busy and not pending:
            return
        time.sleep(0.5)


@contextmanager
def excel_workbook(path):
    xl = win32com.client.DispatchEx("Excel.Application")      # private instance, never the user's window
    xl.Visible, xl.DisplayAlerts = False, False
    wb = None
    try:
        wb = xl.Workbooks.Open(str(path), ReadOnly=True)
        yield wb
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        xl.Quit()


def find_row(ws, label, col=1, max_row=120):
    for r in range(1, max_row):
        v = ws.Cells(r, col).Value
        if isinstance(v, str) and v.split("  (")[0] == label:
            return r
    raise KeyError(label)


def row_values(ws, label, first_col=2, last_col=None):
    r = find_row(ws, label)
    last_col = last_col or ws.UsedRange.Columns.Count
    return [ws.Cells(r, j).Value for j in range(first_col, last_col + 1)]


def summary(wb, price=None):
    """price: written to the override cell so results are repeatable; None keeps whatever is there."""
    a, chk, dcf = wb.Worksheets("Inputs"), wb.Worksheets("Checks & Metrics"), wb.Worksheets("DCF")
    if price is not None:
        a.Cells(find_row(a, "Price override ($)"), 2).Value = price
    live = a.Cells(find_row(a, "Live close (STOCKHISTORY)"), 2)
    recalc(wb, watch=[live])
    out = {"status": chk.Range("D1").Value,
           "live_close": live.Value, "price_date": a.Cells(find_row(a, "Price date"), 2).Value,
           "share_price": a.Cells(find_row(a, "Share price ($)"), 2).Value}
    for label in ("Check: Assets - Liabilities & Equity", "Check: Liabilities + Equity - L&E",
                  "Check: Revenue - COGS - Gross Profit", "Check: Net income IS (incl. NCI) - CF",
                  "Check: BS cash - CF ending cash (forecast)"):
        out[label] = row_values(chk, label)
    out["headers"] = [chk.Cells(4, j).Value for j in range(2, chk.UsedRange.Columns.Count + 1)]
    out["dcf"] = {dcf.Cells(r, 1).Value: (dcf.Cells(r, 2).Value, dcf.Cells(r, 3).Value)
                  for r in range(VAL_HDR + 1, R["implied"] + 1)}
    out["wacc"] = dcf.Range(f"B{R['wacc']}").Value
    out["sens_center"] = tuple(dcf.Range(a).Value for a in SENS_CENTER)
    errors = []
    for ws in wb.Worksheets:
        used = ws.UsedRange
        for row in used.Value or ():
            for v in row if isinstance(row, tuple) else (row,):
                if isinstance(v, int) and v < -2146826000:       # COM error codes (#VALUE!, #REF!, #DIV/0!, ...)
                    errors.append(ws.Name)
    out["error_sheets"] = sorted(set(errors))
    return out


if __name__ == "__main__":
    with excel_workbook(sys.argv[1]) as wb:
        s = summary(wb, float(sys.argv[2]) if len(sys.argv) > 2 else None)
    for k, v in s.items():
        print(f"{k}: {v}")
