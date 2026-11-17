# EDGAR financial models

Builds an Excel valuation model for a public company from its SEC filings: three linked statements, a ten-year
forecast, a DCF, comparable companies and a one-page summary. One command per company, and a roll-up across
a portfolio.

The data comes from the SEC's free XBRL "company facts" API. Prices, exchange rates and betas are live Excel
formulas (`STOCKHISTORY`), so the workbooks need Microsoft 365.

This is a personal research tool. Its output is a starting point for your own assumptions, not investment advice.

## Setup

1. Install [uv](https://docs.astral.sh/uv/). It fetches Python and the packages the first time you run.
2. The SEC asks for a contact email with every request. Put yours in a file named `.edgar_email` next to
   `edgar_etl.py` (one line, git-ignored), or set the `EDGAR_EMAIL` environment variable.
3. Copy `watchlist.example.json` to `watchlist.json` and list your own tickers and peers. Your copy is
   git-ignored, so it stays private.

## Everyday commands

Run these from the project folder. In Git Bash use `./run`; in cmd or PowerShell use `run`.

| Command | What it does |
|---|---|
| `./run ISRG` | Build or rebuild `ISRG_financials.xlsx` |
| `./run --list holdings` | Build every ticker in a `watchlist.json` group (`holdings`, `watchlist`, `all`) |
| `./run --rollup` | Read every built workbook through Excel and write `Portfolio.xlsx` and `valuation_log.csv` |
| `./run --rollup --base-only` | Same, Base scenario only (about a third of the time) |
| `./run ISRG --find accrued` | List the company's XBRL tags matching a word, to fill an empty line |
| `./run ISRG --reset Beta` | Put a typed input back to its default |
| `./run --fetch all` | Download data only, build nothing |
| `./run --comps ISRG SYK MDT` | Standalone comparables table |

Other options: `--peers T1 T2`, `--no-peers`, `--years N`, `--forecast-years N`, `--offline`, `--refresh`,
`--fresh`. Run `./run --help` for the full list.

Close a workbook in Excel before rebuilding it.

## What's in a workbook

| Tab | Purpose |
|---|---|
| Summary | Price, value per share by method with ranges, what the price implies, key figures |
| Inputs | Every setting: scenario, price and currency, dilution, discount rate, terminal value, method choices |
| Drivers | Yearly growth and margin paths, and the Bear / Base / Bull table that sets them |
| Revenue Build | Optional segment build that replaces the single revenue growth driver |
| Income Statement, Balance Sheet, Cash Flow | Five years of history, LTM, year to date, ten forecast years |
| DCF | Free cash flow, two terminal-value methods, sensitivity grids, reverse DCF |
| Comps | Peers, their multiples, and the value they imply |
| Checks & Metrics | Balance checks and ratios |
| Notes | Your thesis and sources (column A is kept across rebuilds) |
| Mapping, Raw | Which filing tag fed each number |

Colours: blue is an input, black a formula, green a link to another tab.

## Things worth knowing

- **Your inputs survive rebuilds.** Anything you type into an input cell is carried into the next build and
  listed in the console. The previous file is copied to `backups/` first. Tabs you add yourself are not carried.
- **Scenarios.** The Scenario dropdown on Inputs switches the whole model. A workbook shows one scenario at a
  time; the roll-up reads all three and applies the probabilities set on Inputs.
- **Year one uses reported results.** When a 10-Q is newer than the last 10-K, the first forecast year is the
  months already reported plus a forecast for the rest.
- **Valued as of today.** The elapsed part of year one is excluded and net debt comes from the latest balance sheet.
- **Defaults are conventions, not views.** Risk-free rate: the current 10-year Treasury yield. Equity risk
  premium: set in `watchlist.json` (currently Damodaran's January 2026 figures). Beta: five-year monthly against
  SPY with the Blume adjustment. Base-case drivers: three-year averages, with year-one growth capped at 30%.
- **Filling gaps.** After a build, any empty line the valuation depends on is listed with the company's closest
  tags. Add the right one to `overrides/<TICKER>.json`; see `overrides/ISRG.json` for the format.

## What it handles and what it doesn't

- **Handles:** US companies filing 10-K and 10-Q; foreign filers on 20-F or 40-F under US GAAP or IFRS, in any
  reporting currency (annual data only), with the dollar share price converted through a live rate and an ADR
  ratio set in `watchlist.json`.
- **Doesn't fit:** banks, insurers and asset managers. They build, but the layout (EBITDA, working capital, net
  debt) is wrong for them.
- **Gaps in the SEC data:** some companies have no usable share count (multiple share classes; type it on
  Inputs), some don't tag a line such as depreciation, and a foreign filer's latest year can be missing.
- **Roll-up:** Windows with Excel only, because it reads calculated values through Excel.

## Tests

```
uv run --with pytest --with requests --with pandas --with openpyxl pytest tests/test_model.py
```

These run offline on saved filings in `tests/fixtures`. `tests/test_excel.py` recalculates workbooks in real
Excel (Windows, add `--with pywin32`).

## Layout of the code

```
edgar_etl.py        command line
edgar/sec.py        downloads and cache (SEC, Treasury yield)
edgar/mapping.py    line item -> XBRL tags (US GAAP and IFRS)
edgar/transform.py  filings -> annual, LTM and year-to-date values
edgar/layout.py     every row of every sheet and its formulas
edgar/xl_*.py       sheet writers (model, inputs, DCF, summary, revenue build)
edgar/comps.py      comparables
edgar/inputs.py     keeps typed inputs across rebuilds
edgar/rollup.py     portfolio roll-up through Excel
```

## Licence

MIT. See `LICENSE`.
