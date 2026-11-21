# Changelog

To move from an older version, see "Upgrading from an older version" in the README: update the code, then
rebuild your workbooks (`./run --list all`). Your typed inputs carry over.

## 2026-11-21

- **Stock splits.** Years reported only before a split kept their pre-split share counts and earnings per share,
  so NVIDIA's history showed a tenfold jump in shares. Splits are now found in the filings (a later filing that
  restates both the share count and EPS of an earlier period by the same ratio) and older figures restated in
  today's shares. The build lists each split; the Raw tab marks every restated figure.
- **Fiscal years are named the way the company names them.** NVIDIA's year ending January 2026 is now FY2026
  (it was FY2025). Companies whose year ends in March to December are unchanged. Inputs typed for a particular
  year move with it when you rebuild.
- A company whose filings contain a mistyped date (Oracle, for one) no longer fails to build.

## 2026-11-20

### Using it
- **Menu.** `./run` with nothing after it (or double-clicking `run.bat`) opens a menu: build a company, build a
  watchlist group, roll-up, write-up, what changed, setup. Each choice prints the command it runs.
- **Choose what to build.** The core model (Inputs, Scenarios, Drivers, the three statements, Ratios, Checks,
  Notes) is always built; the DCF, Multiples, Comps and Revenue Build are optional (`--no-dcf`,
  `--statements-only`, `--parts dcf,multiples`, ...). A workbook remembers its parts.
- **Choose the length.** `--years N` (2-15 years of history) and `--forecast-years N` (3-20), also asked in
  the menu and remembered per workbook.
- **`--open`** opens the workbook in Excel after building.

### New tabs
- **Multiples:** forward P/E, PEG, EV/EBITDA, EV/Sales and price / free cash flow on next-twelve-months figures,
  each with a target, a range and an On / Off switch, beside today's multiple on your forecast, on reported
  results, and the peer median.
- **WACC:** cost of equity (risk-free rate, equity risk premium, beta measured from prices), cost of debt with
  last year's actual rate for reference, market-value weights.
- **Bridge:** enterprise value to equity value to value per share: debt (leases optional), cash, other owners'
  interests, other assets and claims, and the diluted share count. Every valuation reads it.
- **Scenarios:** the Bear / Base / Bull table, your overrides of it, and the scenario probabilities. Drivers
  now holds only the yearly paths and your per-year overrides.
- **Ratios:** growth, margins, returns (ROE, ROIC, ROA), cash flow, leverage and per-share figures, history and
  forecast. **Checks** (formerly Checks & Metrics) is now just the balance checks, each with a tick or a cross.

### Valuation
- **Price target** on the Summary: your weights on each method, a low-high range, a 12-month target, and a
  probability-weighted target across scenarios in the roll-up.
- **Free-cash-flow yield** on the Summary (reported, on your forecast, and after stock-based compensation).
- Price / free cash flow defaults to 0% weight, so the DCF isn't counted twice.

### Inputs
- **Overrides never delete a formula.** Every suggested value (beta, exit multiple, scenario defaults,
  per-year drivers, segment growth, dilution, filed share count) stays in its own cell with a yellow override
  cell beside it. Clear the override to go back. Yellow cells never hold formulas.
- **Inputs sit next to what they change**, and the Inputs tab lists every setting with a link to it.
- A results line at the top of every tab with inputs shows the price target and values as you edit.

### Look
- **Summary dashboard:** headline cards, panels, the valuation table with weights, football-field and
  revenue-and-margin charts, your thesis from Notes, what the price implies, today's multiples, key figures.
- **House style** on every tab: title bars, no gridlines, row rules, hover notes, in the company's brand colour
  (`brand_colors` in `watchlist.json`; meaning colours such as yellow inputs stay fixed).
- **Comps** is now tall: one column per company, one row per metric.
- The income statement shows growth and margins under its key lines; the DCF shows its results first, with the
  two sensitivity grids side by side.

### Earnings season
- **What changed since the last build:** every rebuild reports the new quarter against a year earlier, the
  trailing figures and margins before and after, the balance sheet, and restated years; saved in `changes/`.
  `--since FILE` compares with an older workbook without building. The roll-up shows how each value moved.

### Fixes
- Default formulas saved by Excel (e.g. `20.0` written back as `20`) are no longer mistaken for your edits.
- The launcher works on Mac and Linux; tests run on GitHub on every push.
- Layout: labels no longer cut off on the Scenarios, Multiples, DCF and Comps tabs; the Drivers tab's top row
  no longer repeats the scenario; notes that still pointed to the old tabs now name Scenarios and Bridge.
- README screenshots now show NVIDIA in the current layout; examples in the menu and help use NVDA.

## 2026-11-16

- First public version: statements, DCF, comps, Summary, roll-up and write-up, the financials template for
  banks and insurers, `--setup`.
