"""Which parts of the model a build includes.

The core model is always built: Summary, Inputs, Scenarios, Drivers, the three linked statements, Ratios, Checks,
Notes, Mapping and Raw. The WACC and Bridge tabs come with the DCF (the Bridge also with Multiples).
Everything else can be switched off:
    revenue_build   the Revenue Build tab (segment build that can replace the revenue growth driver)
    dcf             the DCF tab: WACC, two terminal values, sensitivity grids, reverse DCF
    multiples       the Multiples tab: forward P/E, PEG, EV/EBITDA, EV/Sales, price / free cash flow
    comps           the Comps tab (needs peers in watchlist.json or --peers)
A workbook remembers what it was built with, so a rebuild keeps the same parts unless told otherwise."""
import json

PARTS = ("dcf", "multiples", "comps", "revenue_build")
NAMES = {"dcf": "DCF (with the WACC and Bridge tabs)",
         "multiples": "Multiples: P/E, PEG, EV/EBITDA, EV/Sales, P/FCF (with the Bridge tab)",
         "comps": "Comps (peer multiples)", "revenue_build": "Revenue Build (segment by segment)"}
ALWAYS = "Inputs, Scenarios, Drivers, the three statements, Ratios, Checks and Notes"
ALIASES = {"revenue": "revenue_build", "rev": "revenue_build", "build": "revenue_build", "peers": "comps",
           "multiple": "multiples", "pe": "multiples", "peg": "multiples"}
KEY = "_build_parts"            # rows in the hidden _inputs sheet: what the workbook was built with
YEARS_KEY = "_build_years"      # {"history": n, "forecast": n}
HISTORY, FORECAST = (5, (2, 15)), (10, (3, 20))      # default and allowed range of each


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
        return "core model only (no valuation tabs)"
    return "core model + " + ", ".join(NAMES[p].split(" (")[0].split(":")[0] for p in PARTS if p in parts)


def _stored(path):
    """{key: value} of the build rows in an existing workbook's _inputs sheet ({} when there are none)."""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True)
        try:
            if "_inputs" not in wb.sheetnames:
                return {}
            return {key: json.loads(value) for key, _, _, value in
                    wb["_inputs"].iter_rows(min_row=2, max_col=4, values_only=True) if key in (KEY, YEARS_KEY)}
        finally:
            wb.close()
    except Exception:
        return {}


def previous(path):
    """The parts an existing workbook was built with, or None (older workbook, unreadable, or not there)."""
    got = _stored(path).get(KEY)
    return set(got) if got is not None else None


def previous_years(path):
    """(history years, forecast years) an existing workbook was built with, or (None, None)."""
    got = _stored(path).get(YEARS_KEY) or {}
    return got.get("history"), got.get("forecast")


def years(cli_history=None, cli_forecast=None, remembered=(None, None), default=(None, None)):
    """History and forecast years: what you asked for, else what the workbook had, else watchlist.json, else
    5 and 10. Values outside the allowed ranges are pulled back inside them."""
    out = []
    for i, (fallback, (lo, hi)) in enumerate((HISTORY, FORECAST)):
        v = next((x for x in ((cli_history, cli_forecast)[i], remembered[i], default[i]) if x), fallback)
        out.append(max(lo, min(hi, int(v))))
    return tuple(out)


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
