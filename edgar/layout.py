"""
LAYOUT: every row of every model sheet, with its historical and forecast formulas.

All sheets share one column grid (historical FYs, optional LTM, forecast years), so a period sits in the
same Excel column on every sheet. Labels are unique across sheets, so a formula refers to another row
just by its label:  c.v("Revenue") is this column, c.p("Revenue") the prior period, c.pn() the same wrapped in N().

Row kinds:
  Row   - pulled from SEC (label must be in mapping.py). `fallback` fills a historical gap with a formula,
          only when every label in `needs` is available for that column.
  Calc  - not pulled: `plug` is the historical formula (e.g. "Other" lines that make subtotals foot).
Every row may define `fcst`, its formula in forecast columns.

Forecast mechanics (balances by construction):
  cash is the plug: BS cash = CF ending cash; every other BS change flows through the CF.
  Interest uses prior-year balances, so there are no circular references.
"""
from dataclasses import dataclass
from typing import Callable, Optional

import pandas as pd
from openpyxl.utils import get_column_letter

SHEET_NAMES = {"A": "Drivers", "IS": "Income Statement", "BS": "Balance Sheet", "CF": "Cash Flow",
               "DCF": "DCF", "CHK": "Checks & Metrics", "SUM": "Summary", "INP": "Inputs"}
FIRST_ROW = 5

# Revenue Build sheet (optional segment build; rows are fixed, columns follow the shared grid)
RB_SHEET = "Revenue Build"
RB_N = 6                              # segment slots
RB_SEG_FIRST = FIRST_ROW              # segment revenue rows ($ millions, as typed)
RB_TOTAL = RB_SEG_FIRST + RB_N        # total of the segments
RB_MODEL = RB_TOTAL + 1               # reported / model revenue, for comparison
RB_UNALLOC = RB_TOTAL + 2             # reported - segments (history only)
RB_GROW_FIRST = RB_TOTAL + 4          # segment growth rows
RB_GROW_TOTAL = RB_GROW_FIRST + RB_N
RB_SWITCH = "$F$1"                    # "On" / "Off"

SCENARIOS = ("Bear", "Base", "Bull")
SCENARIO_CELL = "$F$1"                # Drivers: shows the scenario chosen on the Inputs tab (the formulas read this)
BASIS_CELL = "$I$1"                   # Drivers: shows what the base case starts from (also set on Inputs)
BASIS_OPTIONS = ("3-yr avg", "Last FY", "LTM")

# ---- Inputs tab: every setting in one place, column B. Section headers are ("#", title).
INPUT_SPEC = [
    ("#", "Scenario"), "Scenario", "Basis", "Use YTD", "w_bear", "w_base", "w_bull", "w_total",
    ("#", "Share price and currency"), "Ticker", "Live close", "Price date", "Price override", "FX rate",
    "FX override", "Shares per listing", "Share price", "Shares outstanding", "Market cap", "52w low", "52w high",
    ("#", "Dilution"), "Options", "Option strike", "RSUs", "Diluted shares",
    ("#", "Discount rate"), "rf", "erp", "Beta benchmark", "Beta years", "Beta frequency", "Computed beta",
    "Adjusted beta", "beta", "kd",
    ("#", "Terminal value"), "g", "multiple", "tv_mode", "roic_t",
    ("#", "Method choices"), "mid", "sbc_mode", "capex_mode", "lease_mode", "val_override",
    ("#", "Equity bridge adjustments"), "other_assets", "other_claims",
]


def _input_rows():
    rows, headers, r = {}, {}, 4
    for item in INPUT_SPEC:
        if isinstance(item, tuple):
            r += 1 if r > 4 else 0          # blank line before each section after the first
            headers[item[1]] = r
        else:
            rows[item] = r
        r += 1
    return rows, headers


INPUT_ROWS, INPUT_HEADERS = _input_rows()


def input_ref(key, sheet=None):
    """Absolute reference to a setting on the Inputs tab (local when written from that tab)."""
    cell = f"$B${INPUT_ROWS[key]}"
    return cell if sheet == "INP" else f"{SHEET_NAMES['INP']}!{cell}"


USE_YTD = f'{input_ref("Use YTD")}="On"'

# Driver groups, shown as headings on the Drivers tab.
DRIVER_GROUPS = {
    "Growth and margins": ("Revenue growth", "Gross margin", "R&D % of revenue", "SG&A % of revenue",
                           "Other opex % of revenue"),
    "Below the operating line": ("Interest rate on debt", "Yield on cash & investments",
                                 "Other non-operating income ($)", "Tax rate", "Other below-tax items ($)",
                                 "NCI share of net income", "Diluted share count change"),
    "Capital intensity": ("Depreciation % of revenue", "Capex % of revenue", "SBC % of revenue"),
    "Working capital": ("DSO (days)", "DIO (days)", "DPO (days)"),
    "Capital returns and financing": ("Dividend payout (% of NI)", "Buybacks (% of NI)",
                                      "Net debt issuance / (repayment)"),
}


@dataclass
class Row:
    label: str
    fcst: Optional[Callable] = None
    plug: Optional[Callable] = None
    fallback: Optional[Callable] = None
    needs: tuple = ()
    data: bool = True
    fmt: str = "usd"
    bold: bool = False
    indent: int = 0
    note: str = ""
    default: str = "avg"      # Assumptions drivers: "avg" = from history (3-yr avg / last FY / LTM), "zero" = 0
    err_default: float = 0    # driver default when there is no history to average
    # Assumptions drivers: how Bear/Bull default relative to Base. ("rel", x): Bull = Base*(1+x), Bear = Base*(1-x)
    # (negative x for costs). ("growth", x): Base +/- the larger of 1 point and x of the base rate.
    spread: Optional[tuple] = None
    cap: Optional[float] = None   # upper limit for the Bear/Bull default (e.g. gross margin <= 100%)
    base_cap: Optional[float] = None   # upper limit for the Base year-1 default taken from history


def Calc(label, **kw):
    return Row(label, data=False, **kw)


def _blank(*refs):
    return "OR(" + ",".join(f'{r}=""' for r in refs) + ")" if len(refs) > 1 else f'{refs[0]}=""'


def _first_with_ytd(c):
    """True for the first forecast year when the grid has a year-to-date column (a 10-Q newer than the 10-K)."""
    return c.g.ytd is not None and c.col is c.g.fcst[0]


def _revenue_fcst(c):
    """Prior revenue x (1 + growth). Year 1 with reported year-to-date figures: the year-to-date revenue plus
    the same remaining months of last year grown at the year-1 rate (LTM - YTD is exactly those months).
    The Revenue Build tab takes over when it is switched on and has last-FY segment revenue typed in."""
    last = c.g.last_fy.L
    rb = f"'{RB_SHEET}'!"
    growth = f"{c.pn('Revenue')}*(1+{c.v('Revenue growth')})"
    if _first_with_ytd(c):
        ytd, ltm = c.ref("Revenue", c.g.ytd), c.ref("Revenue", c.g.ltm)
        growth = (f"IF(AND({USE_YTD},ISNUMBER({ytd}),ISNUMBER({ltm})),"
                  f"{ytd}+({ltm}-{ytd})*(1+{c.v('Revenue growth')}),{growth})")
    return (f'=IF(AND({rb}{RB_SWITCH}="On",SUM({rb}${last}${RB_SEG_FIRST}:${last}${RB_TOTAL - 1})>0),'
            f"{rb}{c.col.L}{RB_TOTAL}*1000000,{growth})")


def _ytd(label, full):
    """Forecast formula for a line that follows revenue. `full(c)` is the full-year expression. In year 1 with
    year-to-date figures it becomes: reported year to date + the full-year expression x the share of the
    year's revenue still to come - so the months already reported are actuals, not forecast."""
    def f(c):
        expr = full(c)
        if not _first_with_ytd(c):
            return "=" + expr
        y, yrev = c.ref(label, c.g.ytd), c.ref("Revenue", c.g.ytd)
        return (f"=IF(AND({USE_YTD},ISNUMBER({y}),ISNUMBER({yrev})),"
                f"{y}+({expr})*(1-{yrev}/{c.v('Revenue')}),{expr})")
    return f


# --------------------------------------------------------------------------
# Income statement
# --------------------------------------------------------------------------
IS_ROWS = [
    Row("Revenue", bold=True, note="Forecast: growth driver, or the Revenue Build tab when it is switched on",
        fcst=lambda c: _revenue_fcst(c)),
    Row("Cost of Revenue", note="Fallback: Revenue - Gross Profit",
        fallback=lambda c: f"={c.v('Revenue')}-{c.v('Gross Profit')}", needs=("Revenue", "Gross Profit"),
        fcst=_ytd("Cost of Revenue", lambda c: f"{c.v('Revenue')}*(1-{c.v('Gross margin')})")),
    Row("Gross Profit", bold=True, note="Fallback: Revenue - Cost of Revenue",
        fallback=lambda c: f"={c.v('Revenue')}-{c.v('Cost of Revenue')}", needs=("Revenue", "Cost of Revenue"),
        fcst=lambda c: f"={c.v('Revenue')}-{c.v('Cost of Revenue')}"),
    Row("R&D Expense", fcst=_ytd("R&D Expense", lambda c: f"{c.v('Revenue')}*{c.v('R&D % of revenue')}")),
    Row("SG&A Expense", fcst=_ytd("SG&A Expense", lambda c: f"{c.v('Revenue')}*{c.v('SG&A % of revenue')}")),
    Calc("Other Operating Expense", note="Plug: Revenue - COGS - R&D - SG&A - Operating Income",
         plug=lambda c: f'=IF({_blank(c.v("Revenue"), c.v("Operating Income"))},"",{c.v("Revenue")}'
                        f'-{c.n("Cost of Revenue")}-{c.n("R&D Expense")}-{c.n("SG&A Expense")}-{c.v("Operating Income")})',
         fcst=_ytd("Other Operating Expense", lambda c: f"{c.v('Revenue')}*{c.v('Other opex % of revenue')}")),
    Row("Operating Income", bold=True,
        note="Fallback when the company doesn't report it: Pre-Tax Income + Interest Expense - Interest Income",
        fallback=lambda c: f"={c.v('Pre-Tax Income')}+{c.n('Interest Expense')}-{c.n('Interest Income')}",
        needs=("Pre-Tax Income",),
        fcst=lambda c: f"={c.v('Revenue')}-{c.v('Cost of Revenue')}-{c.v('R&D Expense')}-{c.v('SG&A Expense')}"
                       f"-{c.v('Other Operating Expense')}"),
    Calc("EBITDA", indent=1, note="Operating Income + D&A (from the cash flow statement)",
         plug=lambda c: f'=IF({_blank(c.v("Operating Income"), c.v("D&A"))},"",{c.v("Operating Income")}+{c.v("D&A")})',
         fcst=lambda c: f"={c.v('Operating Income')}+{c.v('D&A')}"),
    Row("Interest Expense",
        fcst=lambda c: f"={c.v('Interest rate on debt')}*({c.pn('Short-Term Borrowings')}"
                       f"+{c.pn('Current Portion of LT Debt')}+{c.pn('Long-Term Debt')})"),
    Row("Interest Income",
        fcst=lambda c: f"={c.v('Yield on cash & investments')}*({c.pn('Cash & Equivalents')}"
                       f"+{c.pn('Short-Term Investments')}+{c.pn('Long-Term Investments')})"),
    Calc("Other Non-Operating Income", note="Plug: Pre-Tax - Operating Income + Interest Expense - Interest Income",
         plug=lambda c: f'=IF({_blank(c.v("Pre-Tax Income"), c.v("Operating Income"))},"",{c.v("Pre-Tax Income")}'
                        f'-{c.v("Operating Income")}+{c.n("Interest Expense")}-{c.n("Interest Income")})',
         fcst=lambda c: f"={c.v('Other non-operating income ($)')}"),
    Row("Pre-Tax Income", bold=True,
        fcst=lambda c: f"={c.v('Operating Income')}-{c.v('Interest Expense')}+{c.v('Interest Income')}"
                       f"+{c.v('Other Non-Operating Income')}"),
    Row("Income Tax Expense", fcst=lambda c: f"={c.v('Pre-Tax Income')}*{c.v('Tax rate')}"),
    Calc("Other Below-Tax Items", note="Plug: equity-method income, discontinued operations, etc.",
         plug=lambda c: f'=IF({_blank(c.v("Net Income"), c.v("Pre-Tax Income"))},"",{c.v("Net Income")}'
                        f'+{c.n("Net Income to NCI")}-{c.v("Pre-Tax Income")}+{c.n("Income Tax Expense")})',
         fcst=lambda c: f"={c.v('Other below-tax items ($)')}"),
    Row("Net Income to NCI",
        fcst=lambda c: f"=({c.v('Pre-Tax Income')}-{c.v('Income Tax Expense')}+{c.v('Other Below-Tax Items')})"
                       f"*{c.v('NCI share of net income')}"),
    Row("Net Income", bold=True, note="Attributable to common shareholders",
        fcst=lambda c: f"={c.v('Pre-Tax Income')}-{c.v('Income Tax Expense')}+{c.v('Other Below-Tax Items')}"
                       f"-{c.v('Net Income to NCI')}"),
    Row("Diluted Shares", fmt="shares", note="Fallback: Net Income / Diluted EPS (companies with several share classes)",
        fallback=lambda c: f"={c.v('Net Income')}/{c.v('Diluted EPS')}", needs=("Net Income", "Diluted EPS"),
        fcst=lambda c: f"={c.pn('Diluted Shares')}*(1+{c.v('Diluted share count change')})"),
    Row("Diluted EPS", fmt="eps", note="Fallback: Net Income / Diluted Shares",
        fallback=lambda c: f"={c.v('Net Income')}/{c.v('Diluted Shares')}", needs=("Net Income", "Diluted Shares"),
        fcst=lambda c: f'=IFERROR({c.v("Net Income")}/{c.v("Diluted Shares")},"")'),
]

# --------------------------------------------------------------------------
# Balance sheet
# --------------------------------------------------------------------------
def _flat(label):
    return lambda c: f"={c.pn(label)}"


PARENT_EQ = "Shareholders' Equity (Parent)"
MEZZ = "Redeemable NCI & Mezzanine Equity"


def _with_revenue(label):
    return lambda c: f'=IFERROR({c.pn(label)}*{c.v("Revenue")}/{c.p("Revenue")},0)'


BS_ROWS = [
    Row("Cash & Equivalents", fcst=lambda c: f"={c.v('Ending Cash')}"),
    Row("Short-Term Investments", fcst=_flat("Short-Term Investments")),
    Row("Accounts Receivable", fcst=lambda c: f"={c.v('Revenue')}*{c.v('DSO (days)')}/365"),
    Row("Inventory", fcst=lambda c: f"={c.v('Cost of Revenue')}*{c.v('DIO (days)')}/365"),
    Calc("Other Current Assets", note="Plug: Total Current Assets - listed current assets",
         plug=lambda c: f'=IF({_blank(c.v("Total Current Assets"))},"",{c.v("Total Current Assets")}'
                        f'-{c.n("Cash & Equivalents")}-{c.n("Short-Term Investments")}-{c.n("Accounts Receivable")}'
                        f'-{c.n("Inventory")})',
         fcst=_flat("Other Current Assets")),
    Row("Total Current Assets", bold=True,
        fcst=lambda c: f"={c.v('Cash & Equivalents')}+{c.v('Short-Term Investments')}+{c.v('Accounts Receivable')}"
                       f"+{c.v('Inventory')}+{c.v('Other Current Assets')}"),
    Row("PP&E, Net", fcst=lambda c: f"={c.pn('PP&E, Net')}+{c.v('Capex')}-({c.v('D&A')}-{c.v('Amortization of Intangibles')})"),
    Row("Goodwill", fcst=_flat("Goodwill")),
    Row("Intangibles, Net", fcst=lambda c: f"={c.pn('Intangibles, Net')}-{c.v('Amortization of Intangibles')}"),
    Row("Long-Term Investments", note="Marketable securities etc.; counted as cash in net debt",
        fcst=_flat("Long-Term Investments")),
    Calc("Other Non-Current Assets", note="Plug: Total Assets - Total Current Assets - listed non-current assets",
         plug=lambda c: f'=IF({_blank(c.v("Total Assets"), c.v("Total Current Assets"))},"",{c.v("Total Assets")}'
                        f'-{c.v("Total Current Assets")}-{c.n("PP&E, Net")}-{c.n("Goodwill")}-{c.n("Intangibles, Net")}'
                        f'-{c.n("Long-Term Investments")})',
         fcst=_flat("Other Non-Current Assets")),
    Row("Total Assets", bold=True,
        fcst=lambda c: f"={c.v('Total Current Assets')}+{c.v('PP&E, Net')}+{c.v('Goodwill')}+{c.v('Intangibles, Net')}"
                       f"+{c.v('Long-Term Investments')}+{c.v('Other Non-Current Assets')}"),
    Row("Accounts Payable", fcst=lambda c: f"={c.v('Cost of Revenue')}*{c.v('DPO (days)')}/365"),
    Row("Accrued Liabilities", note="Forecast grows with revenue", fcst=_with_revenue("Accrued Liabilities")),
    Row("Deferred Revenue (Current)", note="Forecast grows with revenue", fcst=_with_revenue("Deferred Revenue (Current)")),
    Row("Short-Term Borrowings", fcst=_flat("Short-Term Borrowings")),
    Row("Current Portion of LT Debt", fcst=_flat("Current Portion of LT Debt")),
    Calc("Other Current Liabilities", note="Plug: Total Current Liabilities - listed current liabilities",
         plug=lambda c: f'=IF({_blank(c.v("Total Current Liabilities"))},"",{c.v("Total Current Liabilities")}'
                        f'-{c.n("Accounts Payable")}-{c.n("Accrued Liabilities")}-{c.n("Deferred Revenue (Current)")}'
                        f'-{c.n("Short-Term Borrowings")}-{c.n("Current Portion of LT Debt")})',
         fcst=_flat("Other Current Liabilities")),
    Row("Total Current Liabilities", bold=True,
        fcst=lambda c: f"={c.v('Accounts Payable')}+{c.v('Accrued Liabilities')}+{c.v('Deferred Revenue (Current)')}"
                       f"+{c.v('Short-Term Borrowings')}+{c.v('Current Portion of LT Debt')}"
                       f"+{c.v('Other Current Liabilities')}"),
    Row("Long-Term Debt", fcst=lambda c: f"={c.pn('Long-Term Debt')}+{c.v('Debt Issued')}-{c.v('Debt Repaid')}"),
    Row("Operating Lease Liabilities", fcst=_flat("Operating Lease Liabilities")),
    Calc("Other Non-Current Liabilities", note="Plug: Total Liabilities - Total Current Liabilities - listed items",
         plug=lambda c: f'=IF({_blank(c.v("Total Liabilities"), c.v("Total Current Liabilities"))},"",'
                        f'{c.v("Total Liabilities")}-{c.v("Total Current Liabilities")}-{c.n("Long-Term Debt")}'
                        f'-{c.n("Operating Lease Liabilities")})',
         fcst=_flat("Other Non-Current Liabilities")),
    Row("Total Liabilities", bold=True, note="Fallback: Total L&E - Total Equity",
        fallback=lambda c: f"={c.v('Total Liabilities & Equity')}-{c.v('Total Equity')}",
        needs=("Total Liabilities & Equity", "Total Equity"),
        fcst=lambda c: f"={c.v('Total Current Liabilities')}+{c.v('Long-Term Debt')}"
                       f"+{c.v('Operating Lease Liabilities')}+{c.v('Other Non-Current Liabilities')}"),
    Calc(MEZZ, note="Plug: Total L&E - Total Liabilities - Total Equity (redeemable noncontrolling interests and "
                    "other temporary equity that sits between the two). Forecast: held flat; deducted in the DCF",
         plug=lambda c: f'=IF({_blank(c.v("Total Liabilities & Equity"), c.v("Total Liabilities"), c.v("Total Equity"))},'
                        f'"",{c.v("Total Liabilities & Equity")}-{c.v("Total Liabilities")}-{c.v("Total Equity")})',
         fcst=_flat(MEZZ)),
    Row(PARENT_EQ, note="Fallback: Total Equity - Noncontrolling Interest",
        fallback=lambda c: f"={c.v('Total Equity')}-{c.n('Noncontrolling Interest')}", needs=("Total Equity",),
        fcst=lambda c: f"={c.pn(PARENT_EQ)}+{c.v('Net Income')}"
                       f"-{c.v('Dividends Paid')}-{c.v('Share Repurchases')}+{c.v('Stock-Based Comp')}"),
    Row("Noncontrolling Interest", fcst=lambda c: f"={c.pn('Noncontrolling Interest')}+{c.v('Net Income to NCI')}"),
    Row("Total Equity", bold=True, note="Fallback: Parent equity + NCI",
        fallback=lambda c: f"={c.v(PARENT_EQ)}+{c.n('Noncontrolling Interest')}", needs=(PARENT_EQ,),
        fcst=lambda c: f"={c.v(PARENT_EQ)}+{c.v('Noncontrolling Interest')}"),
    Row("Total Liabilities & Equity", bold=True, note="Fallback: Total Liabilities + Total Equity",
        fallback=lambda c: f"={c.v('Total Liabilities')}+{c.v('Total Equity')}",
        needs=("Total Liabilities", "Total Equity"),
        fcst=lambda c: f"={c.v('Total Liabilities')}+{c.v(MEZZ)}+{c.v('Total Equity')}"),
]

# --------------------------------------------------------------------------
# Cash flow
# --------------------------------------------------------------------------
_WC_ASSETS = ("Accounts Receivable", "Inventory", "Other Current Assets")
_WC_LIABS = ("Accounts Payable", "Accrued Liabilities", "Deferred Revenue (Current)", "Other Current Liabilities")


def _wc_change(c):
    up = lambda l: f"({c.v(l)}-{c.pn(l)})"
    return "=-(" + "+".join(up(l) for l in _WC_ASSETS) + ")+(" + "+".join(up(l) for l in _WC_LIABS) + ")"


CF_ROWS = [
    Row("Net Income (CF)", note="Consolidated net income (incl. NCI)",
        fcst=lambda c: f"={c.v('Net Income')}+{c.v('Net Income to NCI')}"),
    Row("D&A", fcst=_ytd("D&A", lambda c: f"{c.v('Revenue')}*{c.v('Depreciation % of revenue')}"
                                          f"+{c.v('Amortization of Intangibles')}")),
    Row("Amortization of Intangibles", indent=1, note="Of which (included in D&A). Forecast: last FY amount until intangibles run out",
        fcst=lambda c: f"=MIN({c.pn('Intangibles, Net')},{c.last('Amortization of Intangibles', n=True)})"),
    Row("Stock-Based Comp", fcst=_ytd("Stock-Based Comp", lambda c: f"{c.v('Revenue')}*{c.v('SBC % of revenue')}")),
    Calc("Working Capital & Other", note="Plug: CFO - Net Income - D&A - SBC. Forecast: change in working capital",
         plug=lambda c: f'=IF({_blank(c.v("Cash from Operations"))},"",{c.v("Cash from Operations")}'
                        f'-{c.n("Net Income (CF)")}-{c.n("D&A")}-{c.n("Stock-Based Comp")})',
         fcst=_wc_change),
    Row("Cash from Operations", bold=True,
        fcst=lambda c: f"={c.v('Net Income (CF)')}+{c.v('D&A')}+{c.v('Stock-Based Comp')}+{c.v('Working Capital & Other')}"),
    Row("Capex", note="Shown as a positive payment",
        fcst=_ytd("Capex", lambda c: f"{c.v('Revenue')}*{c.v('Capex % of revenue')}")),
    Row("Acquisitions", note="Shown as a positive payment", fcst=lambda c: "=0"),
    Calc("Other Investing", note="Plug: CFI + Capex + Acquisitions",
         plug=lambda c: f'=IF({_blank(c.v("Cash from Investing"))},"",{c.v("Cash from Investing")}'
                        f'+{c.n("Capex")}+{c.n("Acquisitions")})',
         fcst=lambda c: "=0"),
    Row("Cash from Investing", bold=True,
        fcst=lambda c: f"=-{c.v('Capex')}-{c.v('Acquisitions')}+{c.v('Other Investing')}"),
    Row("Debt Issued", fcst=lambda c: f"=MAX({c.v('Net debt issuance / (repayment)')},0)"),
    Row("Debt Repaid", note="Shown as a positive payment", fcst=lambda c: f"=MAX(-{c.v('Net debt issuance / (repayment)')},0)"),
    Row("Dividends Paid", note="Shown as a positive payment",
        fcst=lambda c: f"=MAX(0,{c.v('Net Income')}*{c.v('Dividend payout (% of NI)')})"),
    Row("Share Repurchases", note="Shown as a positive payment",
        fcst=lambda c: f"=MAX(0,{c.v('Net Income')}*{c.v('Buybacks (% of NI)')})"),
    Calc("Other Financing", note="Plug: CFF - Debt Issued + Debt Repaid + Dividends + Buybacks",
         plug=lambda c: f'=IF({_blank(c.v("Cash from Financing"))},"",{c.v("Cash from Financing")}'
                        f'-{c.n("Debt Issued")}+{c.n("Debt Repaid")}+{c.n("Dividends Paid")}+{c.n("Share Repurchases")})',
         fcst=lambda c: "=0"),
    Row("Cash from Financing", bold=True,
        fcst=lambda c: f"={c.v('Debt Issued')}-{c.v('Debt Repaid')}-{c.v('Dividends Paid')}-{c.v('Share Repurchases')}"
                       f"+{c.v('Other Financing')}"),
    Calc("Net Change in Cash", note="CFO + CFI + CFF (excludes FX effects)",
         plug=lambda c: f'=IF({_blank(c.v("Cash from Operations"))},"",{c.v("Cash from Operations")}'
                        f'+{c.n("Cash from Investing")}+{c.n("Cash from Financing")})',
         fcst=lambda c: f"={c.v('Cash from Operations')}+{c.v('Cash from Investing')}+{c.v('Cash from Financing')}"),
    Calc("Ending Cash", bold=True, note="Forecast only: prior cash + net change; feeds the balance sheet",
         fcst=lambda c: f"={c.pn('Cash & Equivalents')}+{c.v('Net Change in Cash')}"),
]

# --------------------------------------------------------------------------
# Assumptions (drivers). Historical columns show the actual ratio. Forecast columns fade in a straight line
# from the selected scenario's year-1 value to its final-year value (the Scenario inputs table below the
# drivers); typing a number over any forecast cell overrides that year.
# --------------------------------------------------------------------------
def _ratio(num, den, n_num=False, mult=""):
    def f(c):
        a = c.n(num) if n_num else c.v(num)
        cond = _blank(c.v(den)) if n_num else _blank(c.v(num), c.v(den))
        return f'=IF({cond},"",IFERROR({a}/{c.v(den)}{mult},""))'
    return f


def _prev_needed(f):
    return lambda c: f(c) if c.prev is not None else None


DRIVER_ROWS = [
    Calc("Revenue growth", fmt="pct", spread=("growth", 0.25), base_cap=0.30,
         note="Base year-1 default is capped at 30%: history above that rarely repeats. Overwrite it if you disagree",
         plug=_prev_needed(lambda c: f'=IF({_blank(c.v("Revenue"), c.p("Revenue"))},"",IFERROR({c.v("Revenue")}/{c.p("Revenue")}-1,""))')),
    Calc("Gross margin", fmt="pct", err_default=1, spread=("rel", 0.02), cap=1, note="100% when the company reports no cost of revenue",
         plug=_ratio("Gross Profit", "Revenue")),
    Calc("R&D % of revenue", fmt="pct", plug=_ratio("R&D Expense", "Revenue", n_num=True)),
    Calc("SG&A % of revenue", fmt="pct", spread=("rel", -0.02), plug=_ratio("SG&A Expense", "Revenue", n_num=True)),
    Calc("Other opex % of revenue", fmt="pct", plug=_ratio("Other Operating Expense", "Revenue")),
    Calc("Interest rate on debt", fmt="pct", note="Interest expense / prior-year total debt",
         plug=_prev_needed(lambda c: f'=IF({_blank(c.v("Interest Expense"))},"",IFERROR({c.v("Interest Expense")}/('
                                     f'{c.pn("Short-Term Borrowings")}+{c.pn("Current Portion of LT Debt")}+{c.pn("Long-Term Debt")}),""))')),
    Calc("Yield on cash & investments", fmt="pct", note="Interest income / prior-year cash + ST and LT investments",
         plug=_prev_needed(lambda c: f'=IF({_blank(c.v("Interest Income"))},"",IFERROR({c.v("Interest Income")}/('
                                     f'{c.pn("Cash & Equivalents")}+{c.pn("Short-Term Investments")}'
                                     f'+{c.pn("Long-Term Investments")}),""))')),
    Calc("Other non-operating income ($)", fmt="usd", note="$ amount (includes interest income when not tagged separately)",
         plug=lambda c: f"={c.v('Other Non-Operating Income')}"),
    Calc("Tax rate", fmt="pct", plug=_ratio("Income Tax Expense", "Pre-Tax Income")),
    Calc("Other below-tax items ($)", fmt="usd", note="$ amount: equity-method income, discontinued operations",
         plug=lambda c: f"={c.v('Other Below-Tax Items')}"),
    Calc("NCI share of net income", fmt="pct",
         plug=lambda c: f'=IF({_blank(c.v("Net Income"))},"",IFERROR({c.n("Net Income to NCI")}/({c.v("Net Income")}'
                        f'+{c.n("Net Income to NCI")}),""))'),
    Calc("Diluted share count change", fmt="pct",
         plug=_prev_needed(lambda c: f'=IF({_blank(c.v("Diluted Shares"), c.p("Diluted Shares"))},"",'
                                     f'IFERROR({c.v("Diluted Shares")}/{c.p("Diluted Shares")}-1,""))')),
    Calc("Depreciation % of revenue", fmt="pct", note="(D&A - amortization of intangibles) / revenue",
         plug=lambda c: f'=IF({_blank(c.v("D&A"), c.v("Revenue"))},"",IFERROR(({c.v("D&A")}'
                        f'-{c.n("Amortization of Intangibles")})/{c.v("Revenue")},""))'),
    Calc("Capex % of revenue", fmt="pct", plug=_ratio("Capex", "Revenue", n_num=True)),
    Calc("SBC % of revenue", fmt="pct", plug=_ratio("Stock-Based Comp", "Revenue", n_num=True)),
    Calc("DSO (days)", fmt="days", plug=_ratio("Accounts Receivable", "Revenue", mult="*365")),
    Calc("DIO (days)", fmt="days", plug=_ratio("Inventory", "Cost of Revenue", mult="*365")),
    Calc("DPO (days)", fmt="days", plug=_ratio("Accounts Payable", "Cost of Revenue", mult="*365")),
    Calc("Dividend payout (% of NI)", fmt="pct", plug=_ratio("Dividends Paid", "Net Income", n_num=True)),
    Calc("Buybacks (% of NI)", fmt="pct", plug=_ratio("Share Repurchases", "Net Income", n_num=True)),
    Calc("Net debt issuance / (repayment)", fmt="usd", default="zero", note="$ amount; forecast defaults to 0",
         plug=lambda c: f'=IF(AND({c.v("Debt Issued")}="",{c.v("Debt Repaid")}=""),"",'
                        f'{c.n("Debt Issued")}-{c.n("Debt Repaid")})'),
]

# --------------------------------------------------------------------------
# Checks & metrics
# --------------------------------------------------------------------------
def _both(f):
    return {"plug": f, "fcst": f}


CHECK_ROWS = [
    Calc("Check: Assets - Liabilities & Equity", **_both(
        lambda c: f'=IF({_blank(c.v("Total Assets"), c.v("Total Liabilities & Equity"))},"n/a",'
                  f'{c.v("Total Assets")}-{c.v("Total Liabilities & Equity")})')),
    Calc("Check: Liabilities + Equity - L&E", **_both(
        lambda c: f'=IF({_blank(c.v("Total Liabilities"), c.v("Total Equity"), c.v("Total Liabilities & Equity"))},"n/a",'
                  f'{c.v("Total Liabilities")}+{c.n(MEZZ)}+{c.v("Total Equity")}-{c.v("Total Liabilities & Equity")})')),
    Calc("Check: Revenue - COGS - Gross Profit", **_both(
        lambda c: f'=IF({_blank(c.v("Revenue"), c.v("Cost of Revenue"), c.v("Gross Profit"))},"n/a",'
                  f'{c.v("Revenue")}-{c.v("Cost of Revenue")}-{c.v("Gross Profit")})')),
    Calc("Check: Net income IS (incl. NCI) - CF", **_both(
        lambda c: f'=IF({_blank(c.v("Net Income"), c.v("Net Income (CF)"))},"n/a",'
                  f'{c.v("Net Income")}+{c.n("Net Income to NCI")}-{c.v("Net Income (CF)")})')),
    Calc("Check: BS cash - CF ending cash (forecast)",
         fcst=lambda c: f"={c.v('Cash & Equivalents')}-{c.v('Ending Cash')}"),
]
METRIC_ROWS = [
    Calc("Operating margin", fmt="pct", **_both(_ratio("Operating Income", "Revenue"))),
    Calc("EBITDA margin", fmt="pct", **_both(_ratio("EBITDA", "Revenue"))),
    Calc("Net margin", fmt="pct", **_both(_ratio("Net Income", "Revenue"))),
    Calc("Free cash flow", note="CFO - Capex", **_both(
        lambda c: f'=IF({_blank(c.v("Cash from Operations"))},"",{c.v("Cash from Operations")}-{c.n("Capex")})')),
    Calc("FCF margin", fmt="pct", **_both(_ratio("Free cash flow", "Revenue"))),
    Calc("FCF conversion", fmt="x", note="FCF / Net Income", **_both(_ratio("Free cash flow", "Net Income"))),
    Calc("Net debt", note="Total debt - cash - ST and LT investments", **_both(
        lambda c: f"={c.n('Short-Term Borrowings')}+{c.n('Current Portion of LT Debt')}+{c.n('Long-Term Debt')}"
                  f"-{c.n('Cash & Equivalents')}-{c.n('Short-Term Investments')}-{c.n('Long-Term Investments')}")),
    Calc("Net debt / EBITDA", fmt="x", **_both(_ratio("Net debt", "EBITDA"))),
]

ROW_BY_LABEL = {r.label: r for rows in (IS_ROWS, BS_ROWS, CF_ROWS, DRIVER_ROWS, CHECK_ROWS, METRIC_ROWS) for r in rows}


# --------------------------------------------------------------------------
# Column grid + formula context
# --------------------------------------------------------------------------
@dataclass
class Col:
    kind: str        # "hist" | "ltm" | "fcst"
    header: str
    end: object      # period end (Timestamp) for hist/ltm
    j: int           # Excel column index

    @property
    def L(self):
        return get_column_letter(self.j)


def fy_year(end):
    """Fiscal-year label: a year ending in Jan-Feb belongs to the prior calendar year (retail convention)."""
    return (end - pd.Timedelta(days=45)).year


class Grid:
    def __init__(self, fy_ends, ltm_end, n_fcst, ytd_months=0):
        self.cols, j = [], 2
        for e in fy_ends:
            self.cols.append(Col("hist", f"FY{fy_year(e)}A", e, j))
            j += 1
        if ltm_end is not None:
            self.cols.append(Col("ltm", f"LTM {ltm_end:%b-%y}", ltm_end, j))
            j += 1
            if ytd_months:
                self.cols.append(Col("ytd", f"YTD {ytd_months}M", ltm_end, j))
                j += 1
        base = fy_year(fy_ends[-1])
        for k in range(1, n_fcst + 1):
            self.cols.append(Col("fcst", f"FY{base + k}E", None, j))
            j += 1
        self.hist = [c for c in self.cols if c.kind == "hist"]
        self.fcst = [c for c in self.cols if c.kind == "fcst"]
        self.last_fy = self.hist[-1]

        self.rows = {}    # label -> (sheet key, row number)
        for key, rows in (("IS", IS_ROWS), ("BS", BS_ROWS), ("CF", CF_ROWS), ("CHK", CHECK_ROWS)):
            self._place(key, [r.label for r in rows], FIRST_ROW)
        # Drivers tab: the driver rows in groups, each under a heading row
        self.driver_headings, r = {}, FIRST_ROW
        for title, labels in DRIVER_GROUPS.items():
            self.driver_headings[title] = r
            self._place("A", labels, r + 1)
            r += len(labels) + 2
        assert all(d.label in self.rows for d in DRIVER_ROWS), "every driver needs a group in DRIVER_GROUPS"
        self.driver_last = r - 2
        # Assumptions tab, below the live driver rows: scenario inputs, selected-scenario results, market data
        self.scen_title = self.driver_last + 2
        self.scen_first = self.scen_title + 3                    # after the title and two header rows
        self.scen_rows = {r.label: self.scen_first + i for i, r in enumerate(DRIVER_ROWS)}
        self.val_title = self.scen_first + len(DRIVER_ROWS) + 1
        self.ltm = next((c for c in self.cols if c.kind == "ltm"), None)
        self.ytd = next((c for c in self.cols if c.kind == "ytd"), None)
        self.metric_header = FIRST_ROW + len(CHECK_ROWS) + 1
        self._place("CHK", [r.label for r in METRIC_ROWS], self.metric_header + 1)

    def _place(self, key, labels, start):
        for i, label in enumerate(labels):
            assert label not in self.rows, f"duplicate row label {label}"
            self.rows[label] = (key, start + i)

    def prev(self, col):
        if col.kind == "hist":
            i = self.hist.index(col)
            return self.hist[i - 1] if i else None
        if col.kind in ("ltm", "ytd"):
            return None
        i = self.fcst.index(col)
        return self.fcst[i - 1] if i else self.last_fy


class Ctx:
    """Formula-building context for one cell: which sheet we're on and which column."""

    def __init__(self, grid, sheet, col):
        self.g, self.sheet, self.col = grid, sheet, col
        self.prev = grid.prev(col) if col is not None else None

    def ref(self, label, col, abs_col=False, abs_row=False):
        sheet, row = self.g.rows[label]
        cell = f"{'$' if abs_col else ''}{col.L}{'$' if abs_row else ''}{row}"
        return cell if sheet == self.sheet else f"'{SHEET_NAMES[sheet]}'!{cell}"

    def v(self, label):
        return self.ref(label, self.col)

    def n(self, label):
        return f"N({self.v(label)})"

    def p(self, label):
        return self.ref(label, self.prev)

    def pn(self, label):
        return f"N({self.p(label)})"

    def last(self, label, n=False):
        r = self.ref(label, self.g.last_fy, abs_col=True)
        return f"N({r})" if n else r

    def market(self, label):
        """Absolute reference to a single-value setting on the Inputs tab (column B)."""
        return input_ref(label, self.sheet)
