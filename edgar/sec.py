"""EXTRACT: SEC EDGAR requests, with a same-day on-disk cache."""
import gzip
import json
import os
from datetime import date
from pathlib import Path

import requests

def _default_email():
    """The contact email the SEC asks for in the User-Agent header. It is personal, so it is kept out of the
    code: set the EDGAR_EMAIL environment variable, or put the address in a file named .edgar_email next to
    edgar_etl.py (git-ignored)."""
    env = os.environ.get("EDGAR_EMAIL", "").strip()
    if env:
        return env
    f = Path(__file__).resolve().parent.parent / ".edgar_email"
    return f.read_text(encoding="utf-8").strip() if f.exists() else ""


DEFAULT_EMAIL = _default_email()     # sent only to sec.gov in the User-Agent header
CACHE_DIR = Path(__file__).resolve().parent.parent / "cache"


class EdgarError(Exception):
    """A problem the user should see as a one-line message (bad ticker, no 10-K data, ...)."""


def sec_get(url, email):
    if "sec.gov" in url and not email:
        raise EdgarError("The SEC asks for a contact email with every download. Put yours in a file named "
                         ".edgar_email next to edgar_etl.py, set EDGAR_EMAIL, or pass --email.")
    agent = f"Personal research {email}" if email else "Personal research"
    r = requests.get(url, headers={"User-Agent": agent}, timeout=30)
    r.raise_for_status()
    return r.json()


OFFLINE = False     # set by --offline: use whatever is in ./cache, however old, and never call sec.gov


def _cache_file(name):
    """Cache files are gzipped (name.gz, ~15x smaller); plain files from older versions are still read."""
    gz, plain = CACHE_DIR / (name + ".gz"), CACHE_DIR / name
    return gz if gz.exists() or not plain.exists() else plain


def _read(path):
    if path.suffix == ".gz":
        with gzip.open(path, "rt", encoding="utf-8") as fh:
            return fh.read()
    return path.read_text(encoding="utf-8")


def _cached_text(name, url, email, refresh):
    """(text, came_from_cache). Reuses cache/<name>[.gz] if it was downloaded today (or at any age when
    OFFLINE); otherwise downloads it and saves it gzipped."""
    path = _cache_file(name)
    if OFFLINE:
        if not path.exists():
            raise EdgarError(f"--offline: {name} is not in the cache folder yet (run once without --offline).")
        return _read(path), True
    if not refresh and path.exists() and date.fromtimestamp(path.stat().st_mtime) == date.today():
        return _read(path), True
    if "sec.gov" in url and not email:
        raise EdgarError("The SEC asks for a contact email with every download. Put yours in a file named "
                         ".edgar_email next to edgar_etl.py, set EDGAR_EMAIL, or pass --email.")
    agent = f"Personal research {email}" if email else "Personal research"
    r = requests.get(url, headers={"User-Agent": agent}, timeout=30)
    r.raise_for_status()
    CACHE_DIR.mkdir(exist_ok=True)
    with gzip.open(CACHE_DIR / (name + ".gz"), "wt", encoding="utf-8") as fh:
        fh.write(r.text)
    return r.text, False


def _cached_get(name, url, email, refresh):
    text, cached = _cached_text(name, url, email, refresh)
    return json.loads(text), cached


def ticker_to_cik(ticker, email, refresh=False):
    data, _ = _cached_get("company_tickers.json", "https://www.sec.gov/files/company_tickers.json", email, refresh)
    want = ticker.upper().replace(".", "-")          # SEC lists BRK.B as BRK-B
    for row in data.values():
        if row["ticker"].upper() == want:
            return int(row["cik_str"]), row["title"]
    raise EdgarError(f"Ticker {ticker} not found in SEC ticker list.")


def fetch_company_facts(cik, email, refresh=False):
    return _cached_get(f"CIK{cik:010d}.json", f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json",
                       email, refresh)


def load_company(ticker, email, refresh=False):
    """Returns (ticker, name, facts_json)."""
    ticker = ticker.upper()
    cik, name = ticker_to_cik(ticker, email, refresh)
    facts, cached = fetch_company_facts(cik, email, refresh)
    how = ("using cached company facts" if OFFLINE else "using company facts cached today") if cached \
        else "downloaded company facts"
    print(f"{name} (CIK {cik:010d}) - {how}")
    return ticker, name, facts


def load_json(path):
    """Load a saved companyfacts JSON (plain or .gz) for offline runs and tests."""
    opener = gzip.open if str(path).endswith(".gz") else open
    with opener(path, "rt", encoding="utf-8") as fh:
        return json.load(fh)


EXCHANGE_CODES = {"NASDAQ": "XNAS", "NYSE": "XNYS"}      # SEC exchange name -> the prefix Excel understands


def price_symbol(ticker, email, refresh=False):
    """Exchange-qualified symbol for Excel's STOCKHISTORY (XNAS:ISRG). A bare ticker can resolve to a foreign
    listing with a very different price. Falls back to the bare ticker if the SEC's exchange list is
    unavailable or the exchange isn't Nasdaq/NYSE."""
    plain = ticker.upper().replace("-", ".")             # Excel writes BRK-B as BRK.B
    try:
        data, _ = _cached_get("company_tickers_exchange.json",
                              "https://www.sec.gov/files/company_tickers_exchange.json", email, refresh)
        fields = data["fields"]
        it, ie = fields.index("ticker"), fields.index("exchange")
        want = ticker.upper().replace(".", "-")
        for row in data["data"]:
            if str(row[it]).upper() == want:
                code = EXCHANGE_CODES.get(str(row[ie] or "").upper())
                return f"{code}:{plain}" if code else plain
    except Exception:
        pass
    return plain


def risk_free_rate(email, refresh=False):
    """(rate, as-of date) for the 10-year Treasury yield: FRED (series DGS10), else the Treasury's own daily
    yield-curve file. (None, None) with a printed reason if neither can be fetched or read. The email is not sent."""
    reasons = []
    try:
        text, _ = _cached_text("DGS10.csv", "https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS10", "", refresh)
        for line in reversed(text.strip().splitlines()[1:]):
            day, _, value = line.partition(",")
            try:
                return round(float(value) / 100, 4), day.strip() + ", FRED"
            except ValueError:            # FRED writes "." on days with no quote
                continue
        reasons.append("FRED: no numeric rows")
    except Exception as e:
        reasons.append(f"FRED: {type(e).__name__}: {str(e)[:80]}")
    try:
        year = date.today().year
        url = ("https://home.treasury.gov/resource-center/data-chart-center/interest-rates/daily-treasury-rates.csv/"
               f"{year}/all?type=daily_treasury_yield_curve&field_tdr_date_value={year}&page&_format=csv")
        text, _ = _cached_text("treasury_yields.csv", url, "", refresh)
        rows = [r.split(",") for r in text.strip().splitlines()]
        col = [h.strip().strip('"') for h in rows[0]].index("10 Yr")
        newest = rows[1]                  # the file lists the latest day first
        return round(float(newest[col]) / 100, 4), newest[0].strip('"') + ", US Treasury"
    except Exception as e:
        reasons.append(f"Treasury: {type(e).__name__}: {str(e)[:80]}")
    if not OFFLINE:
        print("  10-year Treasury yield not downloaded (" + "; ".join(reasons) + ")")
    return None, None
