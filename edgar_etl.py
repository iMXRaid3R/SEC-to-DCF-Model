#!/usr/bin/env python3
"""
edgar_etl.py - Build a financial model workbook from SEC EDGAR (XBRL "company facts"):
Drivers (Bear/Base/Bull scenarios), Income Statement, Balance Sheet, Cash Flow (historicals + LTM + linked
forecast), Checks & Metrics, Mapping and Raw - plus, optionally, DCF, Multiples (P/E, PEG, EV/EBITDA, EV/Sales,
P/FCF), Comps and Revenue Build.

Usage:
    python edgar_etl.py                                      # the menu
    python edgar_etl.py ISRG --statements-only               # just the three linked statements
    python edgar_etl.py ISRG --no-dcf --no-revenue-build     # leave parts out (remembered on rebuild)
    pip install requests pandas openpyxl        (or prefix commands with: uv run --with requests --with pandas --with openpyxl)
    python edgar_etl.py ISRG                                 # peers come from watchlist.json when listed there
    python edgar_etl.py ISRG --peers SYK MDT BSX             # choose the peers for the Comps tab
    python edgar_etl.py MSFT --years 8 --forecast-years 7 --out msft.xlsx
    python edgar_etl.py --list holdings                      # build every ticker in a watchlist.json group
    python edgar_etl.py --comps ISRG SYK MDT BSX             # standalone comps table
    python edgar_etl.py ISRG --find accrued                  # list the company's tags to fill a "not reported" line
    python edgar_etl.py --rollup                             # Portfolio.xlsx + valuation_log.csv from the built workbooks
    python edgar_etl.py --fetch all                          # only download (tickers or watchlist.json groups, plus peers)
    python edgar_etl.py AAPL --refresh                       # ignore today's cached download
    python edgar_etl.py AAPL --offline                       # use ./cache as it is, no download

Rebuilding over an existing workbook keeps what you typed: scenario inputs, per-year overrides, DCF inputs,
segment revenue, price overrides. The old file is first copied to ./backups. Use --fresh to start clean.
Add your own tabs in a separate file: sheets the script didn't create are not carried into a rebuild.

SEC requires a descriptive User-Agent with contact info: put your email in a file named .edgar_email next to
this script (or set EDGAR_EMAIL, or pass --email). It is sent only to sec.gov.
Downloads are cached in ./cache for the rest of the day. Handles 10-K filers and foreign filers on 20-F / 40-F
(US GAAP or IFRS, any reporting currency; annual figures only). The layout does not suit banks and insurers.
"""
import argparse
import json
import sys
import time
from datetime import date
from pathlib import Path

from edgar import parts as parts_mod
from edgar import sec
from edgar.comps import build_comps, peer_rows
from edgar.inputs import backup, carry_over
from edgar.sec import DEFAULT_EMAIL, EdgarError, load_company, load_json
from edgar.transform import find_tags, load_financials
from edgar.changes import report as report_changes
from edgar.xl_fin import coverage_report as fin_coverage
from edgar.xl_model import build_workbook, coverage_report, gap_report

WATCHLIST = Path(__file__).resolve().parent / "watchlist.json"
OVERRIDES = Path(__file__).resolve().parent / "overrides"


def load_overrides(ticker):
    """overrides/<TICKER>.json: extra tags for this company's line items, tried before the standard ones.
    {"Accrued Liabilities": [["EmployeeRelatedLiabilitiesCurrent", "OtherAccruedLiabilitiesCurrent"]]}
    (an inner list = add these together). Find tag names with:  python edgar_etl.py TICKER --find accrued"""
    path = OVERRIDES / f"{ticker.upper()}.json"
    if not path.exists():
        return {}
    from edgar.mapping import MAP_BY_LABEL
    data = {k: v for k, v in json.loads(path.read_text(encoding="utf-8")).items() if not k.startswith("_")}
    unknown = [k for k in data if k not in MAP_BY_LABEL]
    if unknown:
        raise EdgarError(f"{path.name}: unknown line item(s) {unknown}. Names must match the Mapping tab, e.g. "
                         f"{list(MAP_BY_LABEL)[:3]}")
    print(f"Using overrides/{path.name}: {', '.join(data)}")
    return data


def load_watchlist():
    """watchlist.json is personal (git-ignored); watchlist.example.json is the starting point shipped with the code."""
    for path in (WATCHLIST, WATCHLIST.with_name("watchlist.example.json")):
        if path.exists():
            if path != WATCHLIST:
                print(f"(no watchlist.json yet: using {path.name}. Copy it to watchlist.json and edit it.)")
            return json.loads(path.read_text(encoding="utf-8"))
    return {}


def _show(v):
    if isinstance(v, float):
        return f"{v:.4g}"
    return str(v)


def build_one(a, ticker, peers, out=None, quiet=False):
    if a.json:
        facts = load_json(a.json)
        name = facts.get("entityName", "Company")
        ticker = (ticker or name.split()[0]).upper()
    else:
        ticker, name, facts = load_company(ticker, a.email, a.refresh)
    if a.find:
        hits = find_tags(facts, a.find)
        print(f"\nTags matching {' + '.join(a.find)} (latest annual value):" if hits else "\nNo matching tags.")
        for end, tag, label, val, unit in hits:
            print(f"  {end}  {val:>18,.0f} {unit:<11} {tag}   - {label[:60]}")
        print(f"\nTo use one, add it to overrides/{ticker}.json, e.g. {{\"Interest Income\": [\"TagName\"]}}")
        return
    fin = load_financials(facts, ticker, name, a.years, load_overrides(ticker))
    if a.since:
        if not Path(a.since).exists():
            raise EdgarError(f"--since: no file {a.since}")
        print(f"Comparing {Path(a.since).name} with today's SEC data (nothing is built or saved):")
        return report_changes(fin, a.since)
    fin.adr_ratio = a.adr.get(ticker, 1)
    fin.brand = a.brands.get(ticker) if a.use_brand else None
    erp = a.erp.get("by_ticker", {}).get(ticker, a.erp.get("default"))
    if erp:
        own = ticker in a.erp.get("by_ticker", {})
        fin.erp = (float(erp), a.erp.get("source", "watchlist.json") + (f" (rate set for {ticker})" if own else ""))
    if not a.json:
        fin.symbol = sec.price_symbol(ticker, a.email, a.refresh)
        fin.rf = sec.risk_free_rate(a.email, a.refresh)
    age = (date.today() - fin.cols[-1].date()).days
    if age > 500:
        print(f"  ! Latest annual figures are for the year ended {fin.cols[-1].date()} ({age // 30} months ago): "
              "the SEC's data doesn't include a newer annual report for this company yet, so the model is stale.")
    if fin.currency != "USD":
        print(f"Reports in {fin.currency} ({fin.annual_form}): the workbook converts the dollar share price; "
              f"ordinary shares per US-listed share = {fin.adr_ratio:g} (watchlist.json \"adr_ratio\")")
    out = Path(out or f"{ticker}_financials.xlsx")
    remembered = parts_mod.previous(out) if out.exists() and not a.fresh else None
    parts = parts_mod.resolve(a.parts, a.off, remembered, a.default_parts)
    if "comps" in parts and not peers:
        parts.discard("comps")
        if a.parts is not None and "comps" in a.parts:
            print(f"  ! No peers for {ticker}: add them under \"peers\" in watchlist.json or use --peers, for a Comps tab")
    rows = peer_rows(peers, a.email, a.refresh, skip={ticker}, adr=a.adr) if peers and "comps" in parts else []
    template = a.template or a.templates.get(ticker)
    wb, model = build_workbook(fin, a.forecast_years, rows, template, parts)
    if model.kind == "fin":
        print("Financial company: built with the financials template (valued on equity, not free cash flow). "
              "Use --template standard to force the standard model.")
        if a.parts is not None or a.off:
            print("  (--parts and the --no- switches apply to the standard model only)")
    else:
        how = ("as last time" if remembered is not None and parts == remembered and a.parts is None and not a.off
               else "")
        print(f"Parts: {parts_mod.describe(parts)}" + (f" ({how}; change with --parts)" if how else ""))
    if out.exists() and not a.fresh:
        kept, dropped, parked = carry_over(out, wb, model.reg, a.reset or ())
        if kept:
            print(f"Kept {len(kept)} of your inputs from the existing {out.name}:")
            for k, v in kept:
                print(f"    {k.replace('|', ' > ')} = {_show(v)}")
        for k, why in dropped:
            print(f"  ! not carried over: {k.replace('|', ' > ')} - {why}")
        if parked:
            tabs = sorted({k.split("|")[0] for k, _ in parked})
            print(f"Kept {len(parked)} input(s) from tabs not in this build ({', '.join(tabs)}): they come back "
                  "when you build with those parts again.")
    saved = backup(out)
    wb.save(out)
    print(f"Saved {out}" + (f"   (previous version: backups/{saved.name})" if saved else ""))
    if a.open:
        open_file(out)
    changed = report_changes(fin, saved, out.resolve().parent) if saved else None
    if model.kind == "fin":
        if not quiet:
            fin_coverage(model)
    else:
        if not quiet:
            coverage_report(model)
        gap_report(model, facts)
    return changed


def fetch(a, wl, peer_map):
    """Download company facts (and the exchange list and Treasury yield) into ./cache without building."""
    groups = {k: v for k, v in wl.items() if isinstance(v, list)}
    tickers = []
    for name in a.fetch:
        tickers += [t for v in groups.values() for t in v] if name == "all" else groups.get(name, [name])
    tickers += [p for t in list(tickers) for p in peer_map.get(t.upper(), [])]
    ok, failed = 0, []
    for t in dict.fromkeys(x.upper() for x in tickers):
        try:
            load_company(t, a.email, a.refresh)
            ok += 1
        except Exception as e:
            failed.append(t)
            print(f"  {t}: {type(e).__name__}: {e}")
        time.sleep(0.15)                    # stay under the SEC's 10 requests/second
    sec.price_symbol("AAPL", a.email, a.refresh)
    rf, day = sec.risk_free_rate(a.email, a.refresh)
    print(f"\nDownloaded or already cached: {ok} companies" + (f"; failed: {', '.join(failed)}" if failed else ""))
    print(f"10-year Treasury yield: {rf:.2%} as of {day}" if rf else "10-year Treasury yield: not available")


def setup(a):
    """First run: the SEC contact email and a personal watchlist, both kept out of git."""
    root = Path(__file__).resolve().parent
    email_file = root / ".edgar_email"
    email = a.email if a.email and a.email != DEFAULT_EMAIL else ""
    if email_file.exists() and not email:
        print("Contact email: already saved in .edgar_email")
    else:
        if not email:
            print("The SEC asks for a contact email with every data request. It is saved in .edgar_email on this\n"
                  "computer, sent only to sec.gov, and never committed to git.")
            email = input("Your email: ").strip()
        if "@" not in email:
            sys.exit("That doesn't look like an email address. Run --setup again.")
        email_file.write_text(email + "\n", encoding="utf-8")
        print("Saved .edgar_email")
    example = WATCHLIST.with_name("watchlist.example.json")
    if WATCHLIST.exists():
        print("Watchlist: watchlist.json already exists")
    elif example.exists():
        WATCHLIST.write_text(example.read_text(encoding="utf-8"), encoding="utf-8")
        print("Created watchlist.json from the example: edit it to list your own tickers and peers")
    print("\nReady. Try:   ./run AAPL        (builds AAPL_financials.xlsx; open it in Excel)\n"
          "Then:        ./run --list holdings   and   ./run --rollup")


def _parts_arg(text):
    try:
        return parts_mod.parse(text)
    except ValueError as e:
        raise argparse.ArgumentTypeError(str(e))


def open_file(path):
    """Open a workbook in Excel (or whatever opens .xlsx files here)."""
    import os
    import subprocess
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(Path(path).resolve()))           # noqa: the Windows-only call
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception as e:
        print(f"  (could not open {path}: {e})")


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("ticker", nargs="?", help="Ticker, e.g. AAPL")
    p.add_argument("--email", default=DEFAULT_EMAIL,
                   help="Your email (SEC User-Agent requirement; default: .edgar_email file or EDGAR_EMAIL)")
    p.add_argument("--years", type=int, default=5, help="Number of historical fiscal years (default 5)")
    p.add_argument("--forecast-years", type=int, default=10, help="Number of forecast years (default 10)")
    p.add_argument("--out", help="Output .xlsx path (default <TICKER>_financials.xlsx, or comps.xlsx)")
    p.add_argument("--json", help="Load a saved companyfacts JSON (.json or .json.gz) instead of calling SEC")
    p.add_argument("--refresh", action="store_true", help="Download again even if cached today")
    p.add_argument("--offline", action="store_true", help="Use ./cache whatever its age; never call sec.gov")
    p.add_argument("--peers", nargs="+", metavar="TICKER",
                   help="Peers for the Comps tab (default: this ticker's list in watchlist.json)")
    p.add_argument("--parts", type=_parts_arg, metavar="LIST",
                   help="Which parts to build besides the three statements: comma-separated from dcf, multiples, "
                        "comps, revenue_build - or 'all', or 'statements' for the three statements only. Default: "
                        "what the workbook was built with last time, else \"default_parts\" in watchlist.json, else all")
    p.add_argument("--statements-only", dest="parts", action="store_const", const=set(),
                   help="Only the three linked statements (with Drivers, Inputs and Checks): no valuation tabs")
    for part, text in (("dcf", "the DCF tab"), ("multiples", "the Multiples tab (P/E, PEG, EV/EBITDA, ...)"),
                       ("comps", "the Comps tab"), ("revenue-build", "the Revenue Build tab")):
        p.add_argument(f"--no-{part}", dest="off", action="append_const", const=part.replace("-", "_"),
                       help=f"Leave out {text}")
    p.add_argument("--no-peers", dest="off", action="append_const", const="comps", help="Same as --no-comps")
    p.add_argument("--open", action="store_true", help="Open the workbook in Excel when it is built")
    p.add_argument("--template", choices=("standard", "financial"),
                   help="Which model to build. Default: financial for banks and insurers (detected from the "
                        "filings) or tickers listed under \"template\" in watchlist.json, otherwise standard")
    p.add_argument("--fresh", action="store_true", help="Don't carry inputs over from an existing workbook")
    p.add_argument("--reset", nargs="+", metavar="WORD",
                   help="Put typed inputs whose name contains one of these words back to their defaults, "
                        "e.g. --reset Beta  or  --reset \"Risk-free\"")
    p.add_argument("--list", metavar="GROUP", help="Build every ticker in this watchlist.json group "
                                                   "(e.g. holdings, watchlist, or all)")
    p.add_argument("--comps", nargs="+", metavar="TICKER", help="Build a standalone comps table for these tickers")
    p.add_argument("--rollup", nargs="*", metavar="TICKER_OR_GROUP",
                   help="Read every built workbook through Excel (Bear/Base/Bull) and write Portfolio.xlsx plus a "
                        "row per company in valuation_log.csv. Optionally limit to tickers or watchlist.json groups")
    p.add_argument("--writeup", nargs="+", metavar="TICKER",
                   help="Write a one-page Markdown summary of each built workbook into writeups/ (values, "
                        "scenarios, assumptions and your Notes tab), read through Excel")
    p.add_argument("--setup", action="store_true",
                   help="First-time setup: save your SEC contact email, create watchlist.json from the example")
    p.add_argument("--base-only", action="store_true",
                   help="With --rollup: read only the Base scenario (about a third of the time; no Bear/Bull range)")
    p.add_argument("--find", nargs="+", metavar="WORD",
                   help="List this company's XBRL tags whose name or label contains every word, to fill a "
                        "'not reported' line via overrides/<TICKER>.json. Builds nothing")
    p.add_argument("--since", metavar="WORKBOOK",
                   help="Only show what changed between an older workbook of this ticker (e.g. one in backups/) "
                        "and today's SEC data. Builds nothing")
    p.add_argument("--fetch", nargs="+", metavar="TICKER_OR_GROUP",
                   help="Only download SEC data into ./cache for these tickers or watchlist.json groups "
                        "(holdings, watchlist, all) and their peers; builds nothing")
    if argv is None and len(sys.argv) == 1:
        from edgar.menu import menu
        main.open_file = open_file
        return menu(main)
    a = p.parse_args(argv)
    if a.forecast_years < 3:
        p.error("--forecast-years must be at least 3")
    if a.setup:
        return setup(a)
    sec.OFFLINE = a.offline
    wl = load_watchlist()
    peer_map = {k.upper(): v for k, v in wl.get("peers", {}).items()}
    a.adr = {k.upper(): v for k, v in wl.get("adr_ratio", {}).items()}
    a.templates = {k.upper(): v for k, v in wl.get("template", {}).items()}
    a.off = a.off or []
    dp = wl.get("default_parts")
    try:
        a.default_parts = None if dp is None else parts_mod.parse(dp if isinstance(dp, str) else ",".join(dp))
    except ValueError as e:
        p.error(f'watchlist.json "default_parts": {e}')
    a.erp = dict(wl.get("equity_risk_premium", {}))
    a.brands = {k.upper(): v for k, v in wl.get("brand_colors", {}).items() if not k.startswith("_")}
    a.use_brand = wl.get("use_brand_colors", True) is not False
    a.erp["by_ticker"] = {k.upper(): v for k, v in a.erp.get("by_ticker", {}).items()}

    def peers_for(ticker):
        if "comps" in a.off:
            return []
        return a.peers or peer_map.get((ticker or "").upper(), [])

    try:
        if a.writeup:
            from edgar.writeup import writeups
            writeups(Path(__file__).resolve().parent, a.writeup)
        elif a.rollup is not None:
            from edgar.rollup import rollup
            groups = {k: v for k, v in wl.items() if isinstance(v, list)}
            names = [t for g in a.rollup for t in (groups.get(g, [g]) if g != "all" else sum(groups.values(), []))]
            rollup(Path(__file__).resolve().parent, names or None, a.base_only)
        elif a.fetch:
            fetch(a, wl, peer_map)
        elif a.comps:
            wb = build_comps(a.comps, a.email, a.refresh)
            out = a.out or "comps.xlsx"
            wb.save(out)
            print(f"Saved {out}")
        elif a.list:
            groups = {k: v for k, v in wl.items() if isinstance(v, list)}
            tickers = [t for v in groups.values() for t in v] if a.list == "all" else groups.get(a.list)
            if not tickers:
                p.error(f"no group '{a.list}' in {WATCHLIST.name} (have: {', '.join(groups) or 'none'})")
            failed, moved = [], []
            for t in dict.fromkeys(tickers):
                print(f"\n=== {t} ===")
                try:
                    if build_one(a, t, peers_for(t), quiet=True) == "new":
                        moved.append(t)
                except EdgarError as e:
                    failed.append((t, str(e)))
                    print(f"  skipped: {e}")
                except PermissionError as e:
                    failed.append((t, f"{e.filename} is open in Excel"))
                    print(f"  skipped: close {e.filename} in Excel and run again")
            if failed:
                print("\nNot built: " + ", ".join(t for t, _ in failed))
            print("\nNew filings since the last build: " + (", ".join(moved) + "  (reports saved in changes/)"
                                                            if moved else "none"))
        elif a.json or a.ticker:
            build_one(a, a.ticker, peers_for(a.ticker), a.out)
        else:
            p.error("give a ticker, --list GROUP, --json FILE, or --comps TICKER ...")
    except (EdgarError, RuntimeError) as e:
        sys.exit(str(e))
    except PermissionError as e:
        sys.exit(f"Could not write {e.filename} - close it in Excel and run again.")


if __name__ == "__main__":
    main()
