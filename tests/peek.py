"""Print selected rows (in $M) from a recalculated model:  python tests/peek.py FILE"""
import sys

from xl_com import excel_workbook, row_values

ROWS = {
    "Assumptions": ["Revenue growth", "Gross margin", "Tax rate", "Capex % of revenue", "DSO (days)",
                    "Dividend payout (% of NI)", "Buybacks (% of NI)"],
    "Income Statement": ["Revenue", "Operating Income", "EBITDA", "Interest Income", "Net Income", "Diluted EPS"],
    "Cash Flow": ["Cash from Operations", "Capex", "Working Capital & Other", "Share Repurchases", "Ending Cash"],
    "Balance Sheet": ["Cash & Equivalents", "PP&E, Net", "Intangibles, Net", "Total Assets", "Total Equity"],
}


def fmt(v, label):
    if v is None or isinstance(v, str):
        return f"{v or '':>9}"
    if "(days)" in label:
        return f"{v:>9.1f}"
    if abs(v) < 5 or label.endswith("EPS"):
        return f"{v:>9.3f}" if abs(v) < 5 else f"{v:>9.2f}"
    return f"{v / 1e6:>9,.0f}"


with excel_workbook(sys.argv[1]) as wb:
    wb.Application.CalculateFull()
    hdr = [wb.Worksheets("Income Statement").Cells(4, j).Value for j in range(2, 13)]
    print(f"{'':<28}" + "".join(f"{h or '':>9}" for h in hdr))
    for sheet, labels in ROWS.items():
        ws = wb.Worksheets(sheet)
        for label in labels:
            vals = row_values(ws, label, 2, len(hdr) + 1)
            print(f"{label[:27]:<28}" + "".join(fmt(v, label) for v in vals))
