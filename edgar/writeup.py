"""One-page write-up of a model as Markdown: value by method and scenario, what the price implies, key figures,
the main assumptions and your notes. Read through Excel like the roll-up, so the numbers are the calculated ones.
    ./run --writeup ISRG        ->  writeups/ISRG_2026-11-16.md
Paste it into a pitch, a newsletter draft or a doc. Nothing is saved back into the workbook."""
from datetime import date, datetime
from pathlib import Path

from .layout import INPUT_ROWS, SHEET_NAMES
from .xl_summary import WEIGHTS
from .rollup import MULT_METHODS, SCENARIOS, _mid, _num, to_row

BLOCKS = {"summary": ("Summary", "A1:L120"), "inputs": (SHEET_NAMES["INP"], "A1:B70"),
          "drivers": (SHEET_NAMES["SCN"], "A1:M80"), "notes": ("Notes", "A1:A60")}
KEY_DRIVERS = ("Revenue growth", "Gross margin", "R&D % of revenue", "SG&A % of revenue", "Tax rate",
               "Capex % of revenue", "Return on equity", "Dividends % of net income", "Buybacks % of net income")


def collect_blocks(read):
    """`read(sheet, range)` -> rows. Returns the raw blocks the write-up is rendered from. The financials
    template keeps its drivers on the Equity Model tab, so that is tried when there is no Drivers tab."""
    out = {}
    for k, (sheet, rng) in BLOCKS.items():
        try:
            out[k] = read(sheet, rng)
        except Exception:
            try:
                out[k] = read("Equity Model", rng) if k == "drivers" else []
            except Exception:
                out[k] = []
    return out


# ---- formatting -------------------------------------------------------------------------------------------
def money(v):
    return "n/a" if _num(v) is None else f"${v:,.2f}"


def pct(v, digits=1):
    return "n/a" if _num(v) is None else f"{v:.{digits}%}"


def signed(v):
    return "n/a" if _num(v) is None else f"{v:+.0%}"


def millions(v):
    return "" if _num(v) is None else f"{v / 1e6:,.0f}"


def _day(v):
    return v.strftime("%Y-%m-%d") if isinstance(v, (date, datetime)) else (str(v) if v else "n/a")


def _table(header, rows):
    out = ["| " + " | ".join(header) + " |", "|" + "|".join("---" for _ in header) + "|"]
    return "\n".join(out + ["| " + " | ".join(str(c) for c in r) + " |" for r in rows])


def _find(rows, text, col=0):
    return next((i for i, r in enumerate(rows) if isinstance(r[col], str) and r[col].split("  (")[0] == text), None)


# ---- the page ----------------------------------------------------------------------------------------------
def render(ticker, by_scenario, selected, weights, blocks, today=None):
    """Markdown text. `by_scenario` / `selected` / `weights` are what rollup.read_workbook returns."""
    today = today or date.today()
    row = to_row(ticker, by_scenario, selected, today.isoformat(), weights)
    summary, inputs, drivers, notes = (blocks[k] for k in ("summary", "inputs", "drivers", "notes"))
    inp = {k: inputs[r - 1][1] for k, r in INPUT_ROWS.items() if r - 1 < len(inputs)}
    foreign = row["currency"] != "USD"
    lines = [f"# {row['company']}: valuation summary", ""]
    side = {r[3]: (r[4] if r[4] is not None else r[5]) for r in summary[3:20] if len(r) > 5 and isinstance(r[3], str)}
    pd_row = _find(summary, "Price date")
    lines.append(f"*{today:%B %d, %Y}. Share price {money(row['price'])}"
                 f" ({_day(summary[pd_row][1]) if pd_row is not None else 'n/a'})."
                 f" Valued as of {_day(side.get('Valuation date'))}."
                 + (f" The company reports in {row['currency']}; values below are converted to dollars per "
                    "US-listed share." if foreign else "") + "*")
    lines.append("")

    # value per share
    lines += ["## Value per share", ""]
    scen = [by_scenario.get(s, {}) for s in SCENARIOS]
    fin = bool(by_scenario.get("Base", by_scenario.get(selected, {})).get("financial"))
    names = ("Residual income", "Exit price-to-book") if fin else ("DCF: perpetuity growth", "DCF: exit multiple")
    body = []
    if any(s.get("perp") is not None or s.get("exit") is not None for s in scen):
        body = [[names[0]] + [money(s.get("perp")) for s in scen],
                [names[1]] + [money(s.get("exit")) for s in scen],
                ["**Midpoint of the two**"] + [f"**{money(_mid(s.get('perp'), s.get('exit')))}**" for s in scen],
                ["vs. share price"] + [signed(_mid(s.get("perp"), s.get("exit")) / row["price"] - 1)
                                       if row["price"] and _mid(s.get("perp"), s.get("exit")) is not None else "n/a"
                                       for s in scen]]
    for key, label in MULT_METHODS.items():
        if any(s.get(key) is not None for s in scen):
            body.append([label] + [money(s.get(key)) for s in scen])
    if body:
        lines.append(_table(["Method"] + list(SCENARIOS), body))
    else:
        lines.append("*Built with the core model only: no valuation methods in this workbook.*")
    lines.append("")
    extras = []
    if any(s.get("target") is not None for s in scen):
        lines.append(_table(["", *SCENARIOS], [
            ["**Price target (weighted)**"] + [f"**{money(s.get('target'))}**" for s in scen],
            ["vs. share price"] + [signed(s["target"] / row["price"] - 1) if row["price"] and s.get("target")
                                   is not None else "n/a" for s in scen],
            ["12-month price target"] + [money(s.get("target_12m")) for s in scen]]))
        lines.append("")
        weights = [(r[0], r[6]) for r in summary if len(r) > 6 and r[0] in WEIGHTS
                   and _num(r[6]) is not None and r[0] != "Price target (weighted)" and r[6] > 0]
        if weights:
            extras.append("- Price target weights: " + ", ".join(f"{n} {pct(w, 0)}" for n, w in weights)
                          + " (methods without a value are left out and the rest rescaled).")
        if row["target_weighted"] is not None:
            extras.append(f"- **Probability-weighted price target: {money(row['target_weighted'])}** "
                          f"({signed(row['target_weighted'] / row['price'] - 1) if row['price'] else 'n/a'} vs. price).")
    if row["weighted"] is not None:
        w = ", ".join(f"{n} {pct(x, 0)}" for n, x in zip(SCENARIOS, weights))
        extras.append(f"- Probability-weighted DCF midpoint: {money(row['weighted'])} ({signed(row['upside_weighted'])} "
                      f"vs. price), using {w}.")
    if row["peers"] is not None:
        extras.append(f"- Value at peer median multiples: {money(row['peers'])}.")
    i = _find(summary, "52-week trading range")
    if i is not None and _num(summary[i][1]) is not None:
        k = (row["price"] / summary[i][2]) if foreign and _num(summary[i][2]) else 1
        extras.append(f"- 52-week trading range: {money(summary[i][1] * k)} to {money(summary[i][3] * k)}.")
    lines += extras + [""] if extras else []

    # reverse DCF
    base = by_scenario.get("Base", {})
    gm, gn = base.get("growth_model", (None, None)), base.get("growth_needed", (None, None))
    if gm[0] is not None:
        lines += ["## What the share price implies", "",
                  f"The Base case assumes revenue grows {pct(gm[0])} a year over the forecast. To justify today's "
                  f"price with the model's margins, it would need {pct(gn[1])} a year on the exit-multiple method"
                  f" and {pct(gn[0])} on the perpetuity method.", ""]

    # key figures
    k = _find(summary, "Key figures") if _find(summary, "Key figures") is not None else next(
        (i for i, r in enumerate(summary) if isinstance(r[0], str) and r[0].startswith("Key figures")), None)
    if k is not None:
        head = [h for h in summary[k + 1][1:] if h]
        figs = []
        for r in summary[k + 2:k + 10]:
            if not r[0]:
                break
            is_pct = any(w in r[0] for w in ("growth", "margin", "Return on", "% of"))
            cells = [(pct(v) if is_pct else (f"{v:.2f}" if "EPS" in r[0] or "per share" in r[0] else millions(v)))
                     if _num(v) is not None
                     else "" for v in r[1:1 + len(head)]]
            figs.append([r[0]] + cells)
        unit = summary[k][0].replace("Key figures", "").strip(" ()") or "$ millions"
        lines += [f"## Key figures ({unit}, Base case)", "", _table([""] + head, figs), ""]

    # assumptions
    lines += ["## Main assumptions", ""]
    rate = side.get("WACC", side.get("Cost of equity"))
    beta = inp.get("beta") if _num(inp.get("beta")) is None else format(inp.get("beta"), ".2f")
    mult = "n/a" if _num(inp.get("multiple")) is None else format(inp.get("multiple"), ".1f") + "x"
    lines.append(f"- {'Cost of equity' if fin else 'Discount rate (WACC)'}: {pct(rate)}, from a "
                 f"{pct(inp.get('rf'), 2)} risk-free rate, a {pct(inp.get('erp'), 2)} equity risk premium and a "
                 f"beta of {beta}.")
    if fin:
        lines.append(f"- Valued on equity, not free cash flow: book equity plus the present value of returns above "
                     f"the cost of equity. Terminal growth {pct(inp.get('g'))}; exit multiple {mult} price-to-book.")
    else:
        lines.append(f"- Terminal growth {pct(inp.get('g'))}; exit multiple {mult} EV/EBITDA.")
        lines.append(f"- Terminal cash flow: {inp.get('tv_mode')}. Stock-based compensation: {inp.get('sbc_mode')}. "
                     f"Capex: {inp.get('capex_mode')}. Leases: {inp.get('lease_mode')}.")
    t = next((i for i, r in enumerate(drivers) if isinstance(r[0], str) and r[0].startswith("Scenario inputs")), None)
    if t is not None:
        heads = drivers[t + 2]                                   # Driver | Bear x3 | Base x3 | Bull x3 | Selected x3
        o = _find(drivers, "Your overrides")                     # same shape: a value there replaces the default
        over = {r[0]: r for r in drivers[o + 3:o + 30]} if o is not None else {}
        rows_ = []
        for r in drivers[t + 3:t + 30]:
            if r[0] in KEY_DRIVERS:
                ov = over.get(r[0], [None] * len(r))
                v = [ov[j] if j < len(ov) and _num(ov[j]) is not None else r[j] for j in range(len(r))]
                rows_.append([r[0]] + [pct(x) for x in v[4:7]] + [pct(v[1]), pct(v[7])])
        if rows_:
            lines += ["", _table(["Driver (Base case)", str(heads[4]), str(heads[5]), str(heads[6]),
                                  f"Bear {heads[1]}", f"Bull {heads[7]}"], rows_)]
    lines.append("")

    # notes
    section, buf, wrote = None, [], False
    for r in notes[3:] + [[None]]:
        text = r[0].strip() if isinstance(r[0], str) else ""
        if text in ("Thesis", "Key assumptions and why", "Risks and what would change my mind", "Sources and to-do") \
                or r == [None] and section:
            if section and buf:
                lines += [f"## {section}", ""] + buf + [""]
                wrote = True
            section, buf = text, []
        elif text and section:
            buf.append(text if text.startswith(("-", "*")) else f"- {text}")
    if not wrote:
        lines += ["## Notes", "", "*Nothing on the Notes tab yet. Add your thesis there and it will appear here.*", ""]

    checks = row.get("checks")
    lines += ["---", "", f"*Generated from the model workbook ({'model checks OK' if isinstance(checks, str) and 'OK' in checks else 'model checks flag a difference'}). "
              "The figures follow from the assumptions above; this is not investment advice.*", ""]
    return "\n".join(lines)


def writeups(folder, tickers):
    """Write one Markdown page per ticker into writeups/. Needs Windows + Excel + pywin32, like the roll-up."""
    from .rollup import read_workbook
    try:
        import win32com.client
    except ImportError:
        raise RuntimeError("The write-up reads values through Excel and needs the pywin32 package on Windows. "
                           "Use the launcher (./run --writeup TICKER).")
    folder = Path(folder)
    out_dir = folder / "writeups"
    out_dir.mkdir(exist_ok=True)
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible, xl.DisplayAlerts = False, False
    try:
        for t in tickers:
            path = folder / f"{t.upper()}_financials.xlsx"
            if not path.exists():
                print(f"  {t}: no workbook yet (build it first: ./run {t})")
                continue
            print(f"  {t}: opening", end="", flush=True)
            try:
                by_scenario, selected, weights = read_workbook(xl, path, extra=collect_blocks)
                blocks = by_scenario.pop("_extra")
                out = out_dir / f"{t.upper()}_{date.today():%Y-%m-%d}.md"
                out.write_text(render(t.upper(), by_scenario, selected, weights, blocks), encoding="utf-8")
                print(f" -> writeups/{out.name}")
            except Exception as e:
                print(f" skipped: {type(e).__name__}: {str(e)[:120]}")
    finally:
        xl.Quit()
