# Notes for AI coding agents

This project generates Excel valuation workbooks from SEC filings. Read `README.md` first for what it does and
how it is run. This file covers what is not obvious from the code.

## How to work here

- **Run the offline tests after every change:**
  `uv run --with pytest --with requests --with pandas --with openpyxl pytest tests/test_model.py`
- **Formulas are written, never evaluated, by Python.** openpyxl writes formula text; values exist only after
  a spreadsheet application calculates them. To check a change, build a workbook from `tests/fixtures` or
  `cache/` with `--offline`, recalculate it (LibreOffice headless works for everything except the live-price
  functions), and read the values back. Set a price override first, since `STOCKHISTORY` is Excel-only.
- **"It recalculates in LibreOffice" is not "it works in Excel".** Say which one you checked. The roll-up and
  anything involving `STOCKHISTORY`, dynamic arrays or dropdowns can only be confirmed in real Excel.
- **Do not rebuild or overwrite the user's `*_financials.xlsx` files** unless asked. They hold typed inputs and
  may be open in Excel. Build test output somewhere else (`--out`).
- **Never put the user's email in code, commits or commands.** It lives in `.edgar_email` (git-ignored) or
  `EDGAR_EMAIL` and is sent only to sec.gov.
- A sandbox usually cannot reach sec.gov. Use `--offline` with the cache, or ask the user to run `--fetch`.

## Architecture in one paragraph

`transform.load_financials` turns a company-facts JSON into a `Financials` object (annual values, LTM, year to
date, shares, currency). `layout.py` declares every row of every sheet as a `Row` with formula builders;
`Grid` assigns columns (history, LTM, YTD, forecast) and row numbers; `Ctx` turns a row label into a cell
reference, so formulas refer to rows by label. The `xl_*` writers place those formulas. Settings live on the
Inputs tab (`layout.INPUT_SPEC`); the DCF tab links to them.

`changes.py` compares the previous workbook's Raw tab with the new `fin.sources` on every rebuild (the Raw tab
is the record of what a build was fed, so keep its columns and the `[fiscal year to date]` / `[LTM: ...]` /
`[10-Q]` tag suffixes stable). `rollup.value_changes` reports value-per-share moves from `valuation_log.csv`.
`.github/workflows/tests.yml` runs `tests/test_model.py` on every push: keep those tests offline.

## Rules that will bite you

- **Row and setting labels are identifiers.** Formulas, the input registry, the roll-up parser and the tests
  find things by label. Renaming a label is a breaking change; search for every use.
- **Registry keys must stay stable.** Each input cell is registered under a key (for example
  `DCF|Risk-free rate`, `Scenario|Revenue growth|Base|Year 1`). `inputs.carry_over` uses keys to move typed
  values into a rebuilt workbook, so inputs can change sheet and cell freely but must keep their key.
- **No Excel Data Tables.** A data table that depends on `STOCKHISTORY` makes Excel recalculate for a minute
  or more after every edit. Scenario comparison is done by the roll-up instead.
- **No `INDEX(array, SEQUENCE(...))` in written formulas.** Excel inserts an implicit-intersection `@` when it
  loads them from a file and they break. Use `TAKE` / `DROP`. Newer functions need the `_xlfn.` prefix and
  `LET` variables the `_xlpm.` prefix.
- **A link to an empty cell returns 0, not blank.** Links to optional inputs are written as
  `=IF(src="","",src)`.
- **Amounts are stored in full and displayed in millions** through number formats. Per-share values use
  `FMT["price"]`, which `xl_style.set_currency` changes for non-dollar reporters; call it back to `"USD"` in tests.
- **Fallback formulas can create circular references.** `Row.fallback` fires when its `needs` are available,
  including through other fallbacks. Check the three equity lines (parent, noncontrolling, total) before adding one.
- **Balance-sheet integrity.** Forecast cash is the plug; every other balance-sheet change must flow through
  the cash flow statement. `Checks & Metrics` must read "ALL CHECKS OK" for the fixtures and for a sample of
  real companies after any change to `layout.py`.

## Data quirks already handled (do not undo)

- Tags are tried in order per period; a tuple of tags means "sum these, only when all are present".
- Foreign filers: forms 20-F / 40-F count as annual, IFRS tags follow the US GAAP ones in each list, and the
  reporting currency is detected from the facts.
- A bare ticker can resolve to a foreign listing in `STOCKHISTORY`; symbols get an `XNAS:` / `XNYS:` prefix
  from the SEC's exchange list.
- A cover-page share count older than about a year is rejected (multi-class companies report it by class).
- Redeemable noncontrolling interests sit between liabilities and equity; they have their own balance-sheet
  line and are deducted in the DCF.
- Multiples on negative earnings or EBITDA are left blank so they do not distort peer medians.

## Two templates

`build_workbook` builds the standard model, or `xl_fin.build_financial` when the company looks like a bank or
insurer (`Financials.looks_financial`) or `--template financial` is given. The financials workbook is valued on
equity (residual income, exit price-to-book) and has its own row map in `xl_fin.py`, but it shares the Inputs
tab, the registry and the Summary labels, so the roll-up and write-up read both. Keep those Summary labels and
the Inputs row map (`layout.INPUT_SPEC`) in step across the two.

## Known gaps

- The financials template treats all shareholders' equity as common (preferred stock is not separated) and
  needs a parent-equity line in the SEC data.
- Quarterly columns do not exist; year one uses year-to-date totals only.
- The DCF sensitivity grids are stacked, and the statements have no growth or margin rows.
- `tests/test_excel.py`, the roll-up's and the write-up's Excel automation have no coverage outside Windows.
