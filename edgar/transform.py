"""TRANSFORM: company-facts JSON -> annual + LTM values per standardized line item."""
from dataclasses import dataclass

import pandas as pd

from .mapping import ANNUAL_FORMS, FACT_NAMESPACES, QUARTERLY_FORMS, company_maps, fmt_tag
from .sec import EdgarError


@dataclass
class Financials:
    ticker: str
    name: str
    cols: list            # fiscal-year end dates shown (Timestamps, oldest first)
    values: dict          # label -> {fy_end: value}
    ltm_end: object       # Timestamp of the latest 10-Q, or None when no quarter is newer than the last 10-K
    ltm: dict             # label -> LTM value
    sources: pd.DataFrame  # which tag fed each value (Raw tab)
    shares_out: tuple     # (shares, as-of date, source description)
    symbol: str = ""      # exchange-qualified symbol for Excel's STOCKHISTORY, e.g. XNAS:ISRG ("" = use the ticker)
    currency: str = "USD"  # reporting currency of the filings (EUR for ASML, TWD for TSM, ...)
    maps: tuple = ()      # (IS, BS, CF) tag maps actually used, including overrides/<TICKER>.json
    annual_form: str = "10-K"   # form the annual figures came from (10-K, 20-F, 40-F)
    adr_ratio: float = 1  # ordinary shares per US-listed share (5 for TSM); set from watchlist.json
    rf: tuple = (None, None)    # (10-year Treasury yield, as-of date) when it could be fetched
    taxonomy: str = "us-gaap"   # "ifrs-full" for IFRS filers (changes the lease default in the DCF)
    erp: tuple = (None, "")     # (equity risk premium, where it came from), set from watchlist.json
    ytd: dict = None            # label -> fiscal year-to-date value at ltm_end (flows only), from the latest 10-Q
    ytd_months: int = 0         # months covered by ytd
    dilution: dict = None       # {"options", "strike", "rsus"} from the latest annual report, where tagged
    brand: str = None           # accent colour (hex) for the workbook, from watchlist.json "brand_colors"
    fy_names: dict = None       # fiscal-year end -> the year the company calls it (NVIDIA: Jan 2026 = 2026)
    splits: list = None         # stock splits found in the filings, as (factor, after, by): see split_events

    def looks_financial(self):
        """Banks, insurers and asset managers: no classified balance sheet and no operating-income line.
        The standard model (EBITDA, working capital, net debt) does not describe them."""
        has = lambda label: any(c in self.values.get(label, {}) for c in self.cols)
        return not has("Total Current Assets") and not has("Operating Income") and not has("Cost of Revenue")


def facts_to_df(facts_json):
    """Flatten us-gaap / ifrs-full and dei facts from annual and quarterly filings into one DataFrame."""
    rows = []
    for ns in FACT_NAMESPACES + ("dei",):
        for tag, body in facts_json.get("facts", {}).get(ns, {}).items():
            for unit, entries in body.get("units", {}).items():
                for e in entries:
                    form = e.get("form")
                    if form not in ANNUAL_FORMS and form not in QUARTERLY_FORMS:
                        continue
                    rows.append({
                        "ns": ns, "tag": tag, "unit": unit, "start": e.get("start"), "end": e["end"],
                        "val": e["val"], "filed": e["filed"], "form": form, "accn": e.get("accn"),
                        "fy": e.get("fy"), "fp": e.get("fp"),
                    })
    df = pd.DataFrame(rows)
    if df.empty or not (df["ns"].isin(FACT_NAMESPACES) & df["form"].isin(ANNUAL_FORMS)).any():
        raise EdgarError("No annual financial facts (10-K, 20-F or 40-F) found for this CIK. Likely a newly "
                         "registered entity (e.g. a new holding company after a reorganization) with no annual "
                         "report yet, or a foreign filer that doesn't tag its financials.")
    df["end"] = pd.to_datetime(df["end"], errors="coerce")
    df["start"] = pd.to_datetime(df["start"], errors="coerce")
    df["filed"] = pd.to_datetime(df["filed"])
    # a mistyped date (year 0201 or 9999, say) would overflow the date arithmetic below: drop the fact
    ok = df["end"].between("1980-01-01", "2100-01-01") & (df["start"].isna() | df["start"].between("1980-01-01", "2100-01-01"))
    df = df[ok].copy()
    df["days"] = (df["end"] - df["start"]).dt.days
    return df


# --------------------------------------------------------------------------
# Stock splits: filings made before a split report pre-split share counts and per-share figures. Later filings
# restate the periods they repeat, but older years keep their pre-split numbers, so a 10-for-1 split shows up
# as a tenfold jump in the share count. Splits are found from those restatements and older figures adjusted.
# --------------------------------------------------------------------------
_SHARE_TAGS = ("WeightedAverageNumberOfDilutedSharesOutstanding", "WeightedAverageNumberOfSharesOutstandingBasic",
               "WeightedAverageShares", "AdjustedWeightedAverageShares")
_EPS_TAGS = ("EarningsPerShareDiluted", "EarningsPerShareBasic", "EarningsPerShareBasicAndDiluted",
             "DilutedEarningsLossPerShare", "BasicEarningsLossPerShare")


def _nice(f):
    """A measured ratio rounded to the split the company announced (3.0008 -> 3, 1.5 for 3-for-2)."""
    n = round(f * 2) / 2 if f >= 1 else 1 / (round(2 / f) / 2)
    return n if abs(n / f - 1) < 0.02 else f


def _share_facts(df):
    s = df[df["ns"].isin(FACT_NAMESPACES) & df["tag"].isin(_SHARE_TAGS + _EPS_TAGS) & (df["val"] != 0)]
    return s[(s["unit"] == "shares") | s["unit"].str.endswith("/shares")]


def split_events(df):
    """Stock splits as a list of (factor, after, by): figures filed on or before `after` are pre-split, those
    filed on or after `by` already reflect it; in between, see split_scales. Later filings restate the periods
    they repeat, so a split shows as a share count and an earnings per share that change, inversely, by the same
    ratio. A share count restated alone is a tagging or scale error, not a split, so an event needs both."""
    pairs = []                                    # (after, by, ratio, is_share_count)
    for (tag, start, end, unit), g in _share_facts(df).groupby(["tag", "start", "end", "unit"], dropna=False):
        rows = list(g.sort_values("filed").drop_duplicates("filed", keep="last").itertuples())
        for a, b in zip(rows, rows[1:]):
            shares = tag in _SHARE_TAGS
            r = b.val / a.val if shares else a.val / b.val
            if r > 0 and not 0.8 < r < 1.25 and 1 / 60 < r < 60:
                pairs.append((a.filed, b.filed, r, shares))
    events = []                                   # [after, by, share ratios, eps ratios]
    for after, by, r, shares in sorted(pairs, key=lambda p: p[1] - p[0]):      # tightest windows first
        for e in events:
            ref = e[2] or e[3]
            if after < e[1] and by > e[0] and abs(r / (sum(ref) / len(ref)) - 1) < 0.15:
                e[0], e[1] = max(e[0], after), min(e[1], by)
                (e[2] if shares else e[3]).append(r)
                break
        else:
            events.append([after, by, [r] if shares else [], [] if shares else [r]])
    out = []
    for after, by, sh, eps in events:
        if len(sh) >= 2 and eps:                  # basic and diluted, or two periods; and per-share confirms
            sh = sorted(sh)
            out.append((_nice(sh[len(sh) // 2]), after, by))
    return sorted(out, key=lambda e: e[1])


def split_scales(df, events):
    """accession -> factor turning that filing's share counts into today's shares (per-share figures are divided
    by it). A filing made between the last one known to be pre-split and the first known to be post-split is
    placed by its own share count: if it is about the split ratio below the post-split filing's, it is pre-split."""
    s = _share_facts(df)
    level = (s[s["tag"].isin(_SHARE_TAGS)].sort_values(["end", "filed"]).groupby("accn")["val"].last())
    filed = s.groupby("accn")["filed"].max()
    by_level = {}
    for f, after, by in events:
        post = [level[a] for a in filed[filed == by].index if a in level.index]
        by_level[by] = post[0] if post else None
    scales = {}
    for accn, d in filed.items():
        scale = 1.0
        for f, after, by in events:
            if d <= after:
                pre = True
            elif d >= by:
                pre = False
            else:
                ref, own = by_level[by], level.get(accn)
                pre = bool(ref and own and abs(ref / own / f - 1) < 0.25)
            if pre:
                scale *= f
        scales[accn] = scale
    return scales


def adjust_for_splits(df, events):
    """Restate share counts and per-share figures in today's shares."""
    if not events:
        return df
    scales = split_scales(df, events)

    def by_date(when):                            # a filing with no share figures of its own
        return float(pd.Series([f for f, after, _ in events if when <= after]).prod()) if events else 1.0

    df = df.copy()
    is_sh = df["unit"] == "shares"
    per_sh = df["unit"].str.endswith("/shares")
    rows = is_sh | per_sh
    f = pd.Series([scales[a] if a in scales else by_date(w)
                   for a, w in zip(df.loc[rows, "accn"], df.loc[rows, "filed"])], index=df.index[rows], dtype=float)
    df.loc[is_sh, "val"] = df.loc[is_sh, "val"] * f[is_sh[rows].values].values
    df.loc[per_sh, "val"] = df.loc[per_sh, "val"] / f[per_sh[rows].values].values
    df["split"] = 1.0
    df.loc[rows, "split"] = f.values                 # shown on the Raw tab next to the tag
    return df


def fiscal_year_names(df, cols):
    """The fiscal year each column's company calls it. The date rule (a year ending in January or February
    belongs to the year before, so a 52/53-week year ending 1 January stays in the old year) is right for most
    companies but not all: NVIDIA calls the year ending January 2026 fiscal 2026, Target calls its year ending
    then fiscal 2025. The annual reports say which (the "fy" of the 10-K or 20-F that first covered the year),
    so the most common difference between them and the date rule is applied to every column. None when the
    filings don't say."""
    rule = lambda end: (end - pd.Timedelta(days=45)).year
    a = df[df["ns"].isin(FACT_NAMESPACES) & df["form"].isin(ANNUAL_FORMS) & df["days"].between(350, 380)
           & (df["fp"] == "FY") & df["fy"].notna()]
    diffs = []
    for end in cols:
        s = a[a["end"] == end]
        if not s.empty:
            d = int(s.sort_values("filed")["fy"].iloc[0]) - rule(end)
            if d in (-1, 0, 1):
                diffs.append(d)
    if not diffs:
        return None
    best = max(set(diffs), key=diffs.count)
    if diffs.count(best) * 2 <= len(diffs):
        return None
    return {end: rule(end) + best for end in cols}


def sum_components(sub, parts):
    """One synthetic row per period end where every component tag has a value."""
    pieces = [sub[sub["tag"] == p].set_index("end") for p in parts]
    common = set.intersection(*(set(p.index) for p in pieces))
    rows = []
    for end in sorted(common):
        rs = [p.loc[end] for p in pieces]
        rows.append({"tag": fmt_tag(parts), "unit": rs[0]["unit"], "end": end,
                     "val": sum(r["val"] for r in rs), "filed": max(r["filed"] for r in rs),
                     "form": rs[0]["form"], "accn": ", ".join(dict.fromkeys(r["accn"] for r in rs))})
    return pd.DataFrame(rows, columns=["tag", "unit", "end", "val", "filed", "form", "accn"])


def reporting_currency(df):
    """The currency most annual facts are reported in (ISO code, e.g. USD, EUR, TWD)."""
    a = df[df["ns"].isin(FACT_NAMESPACES) & df["form"].isin(ANNUAL_FORMS) & df["unit"].str.fullmatch("[A-Z]{3}")]
    if a.empty:
        return "USD"
    recent = a[a["end"] >= a["end"].max() - pd.Timedelta(days=800)]
    return recent["unit"].value_counts().idxmax()


def _unit(unit, currency):
    return unit.replace("USD", currency)


def build_statement(df, mapping, currency="USD"):
    """
    Returns (values: label -> {period end: value}, sources DataFrame for the Raw/audit tab).
    Rules:
      - annual (10-K) us-gaap facts only
      - flow items: keep only ~full-year durations (350-380 days) so quarterly/YTD facts are excluded
      - dedupe per (tag, end) keeping the most recently FILED value (handles restatements)
      - per period, take the first mapped tag that has a value (coalesce)
      - a tuple entry sums its component tags, only for periods where all components exist
    """
    df = df[df["ns"].isin(FACT_NAMESPACES) & df["form"].isin(ANNUAL_FORMS)]
    values, sources = {}, []
    for label, kind, unit, tags in mapping:
        unit = _unit(unit, currency)
        values[label] = {}
        flat = {t for e in tags for t in (e if isinstance(e, tuple) else (e,))}
        sub = df[(df["tag"].isin(flat)) & (df["unit"] == unit)].copy()
        if kind == "flow":
            sub = sub[(sub["days"] >= 350) & (sub["days"] <= 380)]
        else:
            sub = sub[sub["start"].isna()]
        if sub.empty:
            continue
        sub = sub.sort_values("filed").drop_duplicates(["tag", "end"], keep="last")
        candidates = []
        for rank, entry in enumerate(tags):
            part = sum_components(sub, entry) if isinstance(entry, tuple) else sub[sub["tag"] == entry]
            candidates.append(part.assign(rank=rank))
        candidates = [c for c in candidates if not c.empty]
        if not candidates:
            continue
        sub = pd.concat(candidates, ignore_index=True)
        sub = sub.sort_values(["end", "rank"]).drop_duplicates("end", keep="first")
        values[label] = {r.end: r.val for r in sub.itertuples()}
        for r in sub.itertuples():
            adj = getattr(r, "split", 1.0)
            adj = 1.0 if adj != adj else adj                        # NaN for summed components
            note = "" if abs(adj - 1) < 1e-9 else f" [restated for stock splits: shares x{adj:g}]"
            sources.append({"Line Item": label, "Period End": r.end.date(), "Tag": r.tag + note,
                            "Value": r.val, "Unit": r.unit, "Form": r.form,
                            "Filed": r.filed.date(), "Accession": r.accn})
    return values, pd.DataFrame(sources)


# --------------------------------------------------------------------------
# LTM (last twelve months) from the latest 10-Q
# --------------------------------------------------------------------------
def _latest(s):
    return s.sort_values("filed").iloc[-1]


def _ltm_fact(g, tag, unit, kind, fy_end, q_end):
    """(value, accession, filed, year-to-date value or None, days in the year-to-date period) for one tag, or None.
    Flows: LTM = last fiscal year + current YTD - prior-year YTD. Shares: current YTD average.
    Instants: the balance at the 10-Q date."""
    s = g[(g["tag"] == tag) & (g["unit"] == unit)]
    if s.empty:
        return None
    if kind == "instant":
        cur = s[(s["end"] == q_end) & s["start"].isna()]
        if cur.empty:
            return None
        r = _latest(cur)
        return r["val"], r["accn"], r["filed"], None, 0
    fy = s[s["form"].isin(ANNUAL_FORMS) & s["days"].between(350, 380) & (s["end"] == fy_end)]
    cur = s[(s["end"] == q_end) & (s["start"] > fy_end) & (s["start"] <= fy_end + pd.Timedelta(days=10))]
    if fy.empty or cur.empty:
        return None
    c = _latest(cur)
    if unit == "shares":
        return c["val"], c["accn"], c["filed"], c["val"], c["days"]
    prior_end = q_end - pd.DateOffset(years=1)
    prior = s[((s["end"] - prior_end).dt.days.abs() <= 7) & ((s["days"] - c["days"]).abs() <= 7)]
    if prior.empty:
        return None
    return _latest(fy)["val"] + c["val"] - _latest(prior)["val"], c["accn"], c["filed"], c["val"], c["days"]


def build_ltm(df, mapping, fy_end, currency="USD"):
    """Returns (ltm_end or None, values, sources, ytd values, months in ytd). Skipped when no 10-Q is newer
    than the last annual report (always the case for 20-F / 40-F filers, whose interim reports aren't in the
    SEC's data). ytd holds the fiscal year-to-date amount of each flow line (per-share lines excluded)."""
    g = df[df["ns"].isin(FACT_NAMESPACES)]
    q = g[g["form"].isin(QUARTERLY_FORMS)]
    if q.empty or q["end"].max() <= fy_end:
        return None, {}, [], {}, 0
    q_end = q["end"].max()
    values, sources, ytd, days = {}, [], {}, 0
    for label, kind, unit, tags in mapping:
        unit = _unit(unit, currency)
        for entry in tags:
            parts = entry if isinstance(entry, tuple) else (entry,)
            got = [_ltm_fact(g, t, unit, kind, fy_end, q_end) for t in parts]
            if all(x is not None for x in got):
                values[label] = sum(x[0] for x in got)
                if kind == "flow" and all(x[3] is not None for x in got) and "/" not in unit:
                    ytd[label] = got[0][3] if unit == "shares" else sum(x[3] for x in got)
                    days = max(days, int(got[0][4]))
                    sources.append({"Line Item": label, "Period End": q_end.date(),
                                    "Tag": fmt_tag(entry) + " [fiscal year to date]", "Value": ytd[label],
                                    "Unit": unit, "Form": "10-Q", "Filed": max(x[2] for x in got).date(),
                                    "Accession": got[0][1]})
                sources.append({"Line Item": label, "Period End": q_end.date(),
                                "Tag": fmt_tag(entry) + (" [LTM: FY + YTD - prior YTD]" if kind == "flow" else " [10-Q]"),
                                "Value": values[label], "Unit": unit, "Form": "10-Q",
                                "Filed": max(x[2] for x in got).date(), "Accession": got[0][1]})
                break
    return q_end, values, sources, ytd, round(days / 30.4)


def _latest_annual(df, tags, unit, fy_end):
    """Latest annual value of the first tag that has one at the fiscal year end (balances only)."""
    a = df[df["ns"].isin(FACT_NAMESPACES) & df["form"].isin(ANNUAL_FORMS) & (df["end"] == fy_end)
           & df["start"].isna() & (df["unit"] == unit)]
    for tag in tags:
        s = a[a["tag"] == tag]
        if not s.empty:
            return float(s.sort_values("filed")["val"].iloc[-1])
    return None


def dilution_facts(df, fy_end, currency):
    """Options and unvested stock units outstanding at the last fiscal year end, where the company tags them
    without a breakdown. Used as defaults for the dilution inputs; blank when not found."""
    px = "ShareBasedCompensationArrangementByShareBasedPaymentAward"
    return {"options": _latest_annual(df, [px + "OptionsOutstandingNumber"], "shares", fy_end),
            "strike": _latest_annual(df, [px + "OptionsOutstandingWeightedAverageExercisePrice"],
                                     f"{currency}/shares", fy_end),
            "rsus": _latest_annual(df, [px + "EquityInstrumentsOtherThanOptionsNonvestedNumber"], "shares", fy_end)}


def shares_outstanding(df, values, cols, ltm):
    """Latest cover-page share count (dei), falling back to the latest diluted weighted-average shares."""
    s = df[(df["ns"] == "dei") & (df["tag"] == "EntityCommonStockSharesOutstanding") & (df["unit"] == "shares")]
    s = s[s["end"] >= cols[-1] - pd.Timedelta(days=400)]      # a count from years ago is worse than none
    if not s.empty:
        s = s[s["filed"] == s["filed"].max()]
        s = s[s["end"] == s["end"].max()].drop_duplicates(["accn", "val"])
        return float(s["val"].sum()), s["end"].iloc[0].date(), "cover page of latest filing"
    if "Diluted Shares" in ltm:
        return float(ltm["Diluted Shares"]), None, "LTM diluted weighted-average shares (no cover-page count)"
    last = values.get("Diluted Shares", {}).get(cols[-1])
    if last is not None:
        return float(last), cols[-1].date(), "diluted weighted-average shares (no cover-page count)"
    return None, None, "not available"


def load_financials(facts, ticker, name, years=5, overrides=None):
    df = facts_to_df(facts)
    splits = split_events(df)
    df = adjust_for_splits(df, splits)
    currency = reporting_currency(df)
    maps = company_maps(overrides)
    all_maps = [row for m in maps for row in m]
    values, sources = {}, []
    for m in maps:
        v, s = build_statement(df, m, currency)
        values.update(v)
        sources.append(s)
    # Fiscal years are defined by full-year flow facts, so a stray balance-sheet date can't add a column.
    fy_ends = sorted({e for label, kind, *_ in all_maps if kind == "flow" for e in values[label]})
    if not fy_ends:
        raise EdgarError("No full-year income statement or cash flow data found in this company's annual reports.")
    cols = fy_ends[-years:]
    ltm_end, ltm, ltm_src, ytd, ytd_months = build_ltm(df, all_maps, cols[-1], currency)

    src = pd.concat([s for s in sources if not s.empty], ignore_index=True)
    src = src[src["Period End"].isin({c.date() for c in cols})]
    if ltm_src:
        src = pd.concat([src, pd.DataFrame(ltm_src)], ignore_index=True)
    annual = df[df["ns"].isin(FACT_NAMESPACES) & df["form"].isin(ANNUAL_FORMS) & (df["end"] == cols[-1])]
    form = annual.sort_values("filed")["form"].iloc[-1].replace("/A", "") if not annual.empty else "10-K"
    return Financials(ticker=ticker, name=name, cols=cols, values=values, ltm_end=ltm_end, ltm=ltm,
                      sources=src, shares_out=shares_outstanding(df, values, cols, ltm),
                      currency=currency, maps=maps, annual_form=form,
                      taxonomy=annual["ns"].value_counts().idxmax() if not annual.empty else "us-gaap",
                      ytd=ytd, ytd_months=ytd_months, dilution=dilution_facts(df, cols[-1], currency),
                      fy_names=fiscal_year_names(df, cols), splits=splits)


# --------------------------------------------------------------------------
# Finding tags: python edgar_etl.py ISRG --find interest income
# --------------------------------------------------------------------------
def find_tags(facts, words, limit=25, end=None, units=None, kind=None):
    """Tags whose name or label contains every word, with their latest annual value. Used to fill gaps:
    pick the tag the company actually reports and add it to overrides/<TICKER>.json."""
    words = [w.lower() for w in words]
    out = []
    for ns in FACT_NAMESPACES:
        for tag, body in facts.get("facts", {}).get(ns, {}).items():
            text = f"{tag} {body.get('label') or ''}".lower()
            if not all(w in text for w in words):
                continue
            best = None
            for unit, entries in body.get("units", {}).items():
                for e in entries:
                    if e.get("form") in ANNUAL_FORMS and (best is None or (e["end"], e["filed"]) > (best[0], best[1])):
                        days = (pd.Timestamp(e["end"]) - pd.Timestamp(e["start"])).days if e.get("start") else None
                        if days is None or 350 <= days <= 380:
                            best = (e["end"], e["filed"], e["val"], unit, "flow" if e.get("start") else "instant")
            if best and (end is None or best[0] == end) and (units is None or best[3] in units) \
                    and (kind is None or best[4] == kind):
                out.append((best[0], tag, body.get("label") or "", best[2], best[3]))
    out.sort(key=lambda r: (r[0], abs(r[3]) if isinstance(r[3], (int, float)) else 0), reverse=True)
    return out[:limit]
