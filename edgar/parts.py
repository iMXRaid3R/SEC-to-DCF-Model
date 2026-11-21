"""Which parts of the model a build includes.

The three linked statements (with the Drivers, Inputs, Checks, Notes, Mapping and Raw tabs) are always built.
Everything else can be switched off:
    revenue_build   the Revenue Build tab (segment build that can replace the revenue growth driver)
    dcf             the DCF tab: WACC, two terminal values, sensitivity grids, reverse DCF
    multiples       the Multiples tab: forward P/E, PEG, EV/EBITDA, EV/Sales, price / free cash flow
    comps           the Comps tab (needs peers in watchlist.json or --peers)
A workbook remembers what it was built with, so a rebuild keeps the same parts unless told otherwise."""
import json

PARTS = ("dcf", "multiples", "comps", "revenue_build")
NAMES = {"dcf": "DCF", "multiples": "Multiples (P/E, PEG, EV/EBITDA, EV/Sales, P/FCF)", "comps": "Comps",
         "revenue_build": "Revenue Build"}
ALIASES = {"revenue": "revenue_build", "rev": "revenue_build", "build": "revenue_build", "peers": "comps",
           "multiple": "multiples", "pe": "multiples", "peg": "multiples"}
KEY = "_build_parts"            # row in the hidden _inputs sheet


def parse(text):
    """'dcf,comps' / 'all' / 'statements' -> set of parts. Raises ValueError for an unknown name."""
    text = text.strip().lower()
    if text in ("all", "full", "everything"):
        return set(PARTS)
    if text in ("statements", "none", "3", "three"):
        return set()
    out = set()
    for word in text.replace(" ", ",").split(","):
        if not word:
            continue
        word = ALIASES.get(word, word.replace("-", "_"))
        if word not in PARTS:
            raise ValueError(f"unknown part '{word}': choose from {', '.join(PARTS)}, all, or statements")
        out.add(word)
    return out


def describe(parts):
    if not parts:
        return "three statements only"
    return "statements + " + ", ".join(NAMES[p].split(" (")[0] for p in PARTS if p in parts)


def previous(path):
    """The parts an existing workbook was built with, or None (older workbook, unreadable, or not there)."""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True)
        try:
            if "_inputs" not in wb.sheetnames:
                return None
            for key, _, _, value in wb["_inputs"].iter_rows(min_row=2, max_col=4, values_only=True):
                if key == KEY:
                    return set(json.loads(value))
        finally:
            wb.close()
    except Exception:
        return None
    return None


def resolve(cli=None, off=(), remembered=None, default=None):
    """What to build: explicit --parts / --only wins, then what the workbook was built with last time, then the
    watchlist.json default (or everything). --no-X switches turn parts off on top of whichever was chosen."""
    if cli is not None:
        parts = set(cli)
    elif remembered is not None:
        parts = set(remembered)
    elif default is not None:
        parts = set(default)
    else:
        parts = set(PARTS)
    return parts - set(off)
