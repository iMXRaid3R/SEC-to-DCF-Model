"""Shared Excel styling. Color convention: blue = input (SEC data or user assumption),
black = formula, green = straight link to another sheet."""
import re

from openpyxl.styles import Alignment, Font, PatternFill

BLUE, BLACK, GREEN, GREY = "0000FF", "000000", "008000", "808080"
FILL = {"hist": PatternFill("solid", fgColor="1F3864"), "ltm": PatternFill("solid", fgColor="595959"),
        "ytd": PatternFill("solid", fgColor="7F7F7F"),
        "fcst": PatternFill("solid", fgColor="1E5C4F"), "label": PatternFill("solid", fgColor="1F3864")}
FCST_BG = PatternFill("solid", fgColor="EEF4F1")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")       # cells the user is expected to fill in
HDR_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=13)
NOTE_FONT = Font(italic=True, color=GREY)
CENTER = Alignment(horizontal="center")

# Dollar amounts and share counts are stored exactly and displayed in millions (the ",," scaling).
FMT = {
    "usd": '#,##0.0,,;(#,##0.0,,);"-"',
    "shares": '#,##0.0,,;(#,##0.0,,);"-"',
    "eps": '0.00;(0.00);"-"',
    "pct": '0.0%;-0.0%;"-"',
    "days": '0.0;-0.0;"-"',
    "x": '0.0"x";-0.0"x";"-"',
    "num": '0.00',
    "price": '$#,##0.00',          # per-share amounts in the reporting currency (set_currency changes it)
    "usd_price": '$#,##0.00',      # the US listing's price, always dollars
    "fx": '0.0000',
    "date": "yyyy-mm-dd",
}


CURRENCY = "USD"


def set_currency(code):
    """Per-share amounts show "$" for US-dollar reporters and the ISO code otherwise (e.g. EUR 612.40)."""
    global CURRENCY
    CURRENCY = code
    FMT["price"] = '$#,##0.00' if code == "USD" else f'"{code} "#,##0.00'


def stock_range(ticker_ref, fn):
    """MIN or MAX close over the last 365 days (52-week low / high)."""
    return f'=IFERROR({fn}(_xlfn.STOCKHISTORY({ticker_ref},TODAY()-365,TODAY(),0,0,1)),"")'

_LINK = re.compile(r"^='[^']+'!\$?[A-Z]+\$?\d+$")


def stockhistory(ticker_ref, field):
    """Latest row of Excel's STOCKHISTORY over the last 10 days (covers weekends/holidays).
    field 1 = date, 2 = close. Needs Microsoft 365 + internet; blank otherwise.
    The _xlfn. prefix is how newer Excel functions are stored in the file (without it: #NAME?)."""
    sh = f"_xlfn.STOCKHISTORY({ticker_ref},TODAY()-10,TODAY(),0,0,0,1)"
    return f'=IFERROR(INDEX({sh},ROWS({sh}),{field}),"")'


def font(color=BLACK, bold=False, italic=False):
    return Font(color=color, bold=bold, italic=italic)


def formula_font(formula, bold=False):
    return font(GREEN if _LINK.match(formula) else BLACK, bold)


def header_cell(cell, value, kind="label"):
    cell.value = value
    cell.fill = FILL[kind]
    cell.font = HDR_FONT
    cell.alignment = CENTER
