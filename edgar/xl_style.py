"""Shared Excel styling (the house style).

Meaning is carried by fixed colours that never change: yellow fill = an input, blue type = SEC data, black =
formula, green = a link to another tab, green / red = up / down, gold = the price target.
The accent (title bars, section rules, column heads, charts) can take the company's brand colour: set_theme()
makes it readable under white text and derives the lighter tints from it."""
import re

from openpyxl.styles import Alignment, Font, PatternFill  # noqa: PatternFill is also used by other modules as st.PatternFill

BLUE, BLACK, GREEN, GREY = "0000FF", "000000", "008000", "808080"
FILL = {"hist": PatternFill("solid", fgColor="1F3864"), "ltm": PatternFill("solid", fgColor="595959"),
        "ytd": PatternFill("solid", fgColor="7F7F7F"),
        "fcst": PatternFill("solid", fgColor="1E5C4F"), "label": PatternFill("solid", fgColor="1F3864")}
FCST_BG = PatternFill("solid", fgColor="EEF4F1")
DEFAULT_ACCENT = "1F3864"                                   # navy, when a company has no brand colour
THEME = {}                                                  # filled by set_theme()
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")       # cells the user is expected to fill in
WARN_FILL = PatternFill("solid", fgColor="FDE2E1")        # a range warning
OVERRIDDEN_FILL = PatternFill("solid", fgColor="FCE4D6")  # a formula cell showing a value you overrode elsewhere
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


def _rgb(h):
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _hex(rgb):
    return "".join(f"{max(0, min(255, round(v))):02X}" for v in rgb)


def mix(color, other, t):
    """color moved t (0..1) of the way toward other."""
    a, b = _rgb(color), _rgb(other)
    return _hex(x + (y - x) * t for x, y in zip(a, b))


def luminance(color):
    """WCAG relative luminance, 0 (black) .. 1 (white)."""
    def ch(v):
        v /= 255
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = (ch(v) for v in _rgb(color))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def readable_accent(color):
    """A brand colour adjusted so white text on it is easy to read (contrast at least 4.5:1) and so it is not
    flat black: light colours are darkened step by step, near-black is lifted to a charcoal of the same hue."""
    c = (color or DEFAULT_ACCENT).strip().lstrip("#").upper()
    if not re.fullmatch(r"[0-9A-F]{6}", c):
        c = DEFAULT_ACCENT
    while luminance(c) > 0.18:
        c = mix(c, "000000", 0.08)
    if luminance(c) < 0.012:
        c = mix(c, "FFFFFF", 0.2)
    return c


def set_theme(color=None):
    """Use `color` (hex, e.g. "0064E0") as the accent for this workbook; None = the default navy."""
    global FCST_BG
    acc = readable_accent(color)
    THEME.update(accent=acc, dark=mix(acc, "000000", 0.35), mid=mix(acc, "FFFFFF", 0.4),
                 light=mix(acc, "FFFFFF", 0.7), pale=mix(acc, "FFFFFF", 0.93), brand=bool(color))
    FILL["hist"] = FILL["label"] = PatternFill("solid", fgColor=acc)
    FILL["fcst"] = PatternFill("solid", fgColor=THEME["dark"])
    FCST_BG = PatternFill("solid", fgColor=THEME["pale"])


set_theme()


def row_rules(ws, rows, last_col, bold_rows=()):
    """Hairline under each row across columns 1..last_col; a thin accent rule above subtotal rows."""
    from openpyxl.styles import Border, Side
    hair = Side(style="hair", color="D0D7E2")
    top = Side(style="thin", color=THEME["mid"])
    for r in rows:
        for j in range(1, last_col + 1):
            ws.cell(row=r, column=j).border = Border(bottom=hair, top=top if r in bold_rows else None)


def section(ws, r, c1, c2, text, size=12):
    """A section title in the accent colour with a rule under it across columns c1..c2 (numbers)."""
    from openpyxl.styles import Border, Side
    ws.cell(row=r, column=c1, value=text).font = Font(bold=True, size=size, color=THEME["accent"])
    for j in range(c1, c2 + 1):
        ws.cell(row=r, column=j).border = Border(bottom=Side(style="medium", color=THEME["accent"]))


def hover_note(cell, text, keep=110):
    """Long explanations become a hover note; the cell keeps the first sentence when the text is long."""
    from openpyxl.comments import Comment
    if not text or len(text) <= keep:
        cell.value = text
        return cell
    first = re.split(r"(?<=[.:;])\s", text, maxsplit=1)[0]
    cell.value = ("How it works  (hover)" if keep == 0 else
                  (first if len(first) <= keep else text[:keep].rsplit(" ", 1)[0]) + "  (hover for more)")
    cell.comment = Comment(text, "Model")
    cell.comment.width, cell.comment.height = 360, max(80, 18 * (len(text) // 55 + 1))
    return cell


def title_bar(ws, last_col=None):
    """Turn row 1 into the tab's title bar: accent fill across the tab, white type (the tab title in A1 and
    whatever else row 1 holds, such as the results line), no gridlines."""
    from openpyxl.utils import get_column_letter
    if last_col is None:            # wide enough for what row 1 holds (a formula = the results line, ~120 characters)
        width = lambda j: ws.column_dimensions[get_column_letter(j)].width or 8.43
        need, at = 0.0, 0.0
        for j in range(1, ws.max_column + 1):
            v = ws.cell(row=1, column=j).value
            if v is not None:
                text = 125 if isinstance(v, str) and v.startswith("=") else len(str(v)) * (1.35 if j == 1 else 1.1)
                need = max(need, at + text)
            at += width(j)
        last_col, at = 0, 0.0
        while at < max(need, 100) and last_col < 40:
            last_col += 1
            at += width(last_col)
        last_col = max(last_col, min(ws.max_column, 40))     # and across the whole table
    fill = PatternFill("solid", fgColor=THEME["accent"])
    for j in range(1, last_col + 1):
        c = ws.cell(row=1, column=j)
        if c.fill is not None and c.fill.fgColor is not None and str(c.fill.fgColor.rgb).endswith("FFF2CC"):
            continue                                        # an input living on row 1 (Revenue Build's switch) stays yellow
        c.fill = fill
        if c.value is not None:
            link = c.hyperlink is not None
            c.font = Font(color="FFFFFF", bold=j == 1 or not link, size=14 if j == 1 else 10, underline="single" if link else None)
            c.alignment = Alignment(vertical="center", horizontal=c.alignment.horizontal if j > 1 else None)
    ws.row_dimensions[1].height = 26
    a2 = ws["A2"]
    if isinstance(a2.value, str) and not a2.value.startswith("="):
        a2.font = Font(italic=True, size=9, color="6B7280")
    ws.sheet_view.showGridLines = False


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


INPUT_LEGEND = "Yellow cells are inputs: change them freely, the rebuild keeps them. Everything else is calculated."


def mark_input(cell):
    """The look of a cell that is safe to change: blue type on a yellow fill."""
    cell.fill = INPUT_FILL
    cell.font = Font(color=BLUE, bold=bool(cell.font and cell.font.bold))


LINK_FONT = Font(color="0563C1", underline="single")


def link_to(cell, sheet, addr, text=None):
    """Make `cell` a click-through to sheet!addr (an Excel in-workbook hyperlink)."""
    if text is not None:
        cell.value = text
    from openpyxl.worksheet.hyperlink import Hyperlink
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{sheet}'!{addr}", display=str(cell.value or ""))
    cell.font = Font(color="0563C1", underline="single", bold=bool(cell.font and cell.font.bold))
    return cell


def font(color=BLACK, bold=False, italic=False):
    return Font(color=color, bold=bold, italic=italic)


def formula_font(formula, bold=False):
    return font(GREEN if _LINK.match(formula) else BLACK, bold)


def header_cell(cell, value, kind="label"):
    cell.value = value
    cell.fill = FILL[kind]
    cell.font = HDR_FONT
    cell.alignment = CENTER
