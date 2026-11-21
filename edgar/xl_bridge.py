"""Bridge tab: from enterprise value to equity value to value per share, in one place.

    Share count              shares outstanding (Inputs), options (treasury method), stock units -> diluted shares
    Enterprise value -> equity  debt (with leases when you say so), cash and investments, noncontrolling and
                             redeemable interests, other assets and claims -> net claims ahead of common equity
    At today's price         equity value and enterprise value, for comparison with the valuations

The DCF, the Multiples tab and the WACC's debt weight all read from here. Built whenever the DCF or the Multiples
tab is. Settings keep the registry keys they had before (on Inputs and on the DCF tab), so typed values carry
over."""
from openpyxl.styles import Alignment, Font

from . import xl_inputs
from . import xl_style as st
from .layout import LEASE_MODES, Ctx

SHEET = "Bridge"
K = {"sh_title": 4, "sh_hdr": 5, "shares": 6, "options": 7, "strike": 8, "rsus": 9, "price": 10, "itm": 11,
     "diluted": 12,
     "ev_title": 14, "ev_hdr": 15, "bs_date": 16, "borrow": 17, "leases": 18, "lease_mode": 19, "debt": 20,
     "cash": 21, "net_debt": 22, "nci": 23, "other_assets": 24, "other_claims": 25, "claims": 26,
     "now_title": 28, "mcap": 29, "ev": 30}


def ref(key, col="B"):
    """Absolute reference to a Bridge cell from another sheet, e.g. ref("claims")."""
    return f"{SHEET}!${col}${K[key]}"


def write_bridge(ws, m):
    g, fin = m.grid, m.fin
    foreign = fin.currency != "USD"
    cur = fin.currency
    last = Ctx(g, "BRG", g.last_fy)
    bs_col = g.ltm if g.ltm is not None and m.has_data("Total Assets", g.ltm) else g.last_fy
    bs = Ctx(g, "BRG", bs_col)
    B = lambda k: f"$B${K[k]}"
    ws["A1"], ws["A1"].font = f"{fin.name} ({fin.ticker}) - Bridge to equity", st.TITLE_FONT
    ws["A2"] = ("From enterprise value to value per share: the claims that come before common shareholders, and "
                "the share count they are divided by. The DCF and Multiples tabs read from here. "
                + ("Millions except per share." if not foreign else f"{cur} millions; per share in {cur} per "
                                                                    "ordinary share."))
    ws["A2"].font = st.NOTE_FONT

    def calc(key, label, f, fmt, note="", bold=False):
        ws.cell(row=K[key], column=1, value=label).font = st.font(bold=bold)
        c = ws.cell(row=K[key], column=2, value=f)
        c.number_format, c.alignment = st.FMT[fmt], Alignment(horizontal="right")
        c.font = st.formula_font(f, bold) if isinstance(f, str) else st.font(st.BLUE, bold)
        if note:
            st.hover_note(ws.cell(row=K[key], column=6), note).font = st.NOTE_FONT
        return c

    # ---- share count
    st.section(ws, K["sh_title"], 1, 6, "Share count")
    xl_inputs.override_header(ws, K["sh_hdr"])
    calc("shares", "Shares outstanding", f"=N({last.market('Shares outstanding')})", "shares",
         "From the Inputs tab (the filed count, or your override there)")
    d = fin.dilution or {}
    filed = lambda v, f: f"Filed in the latest annual report: {f.format(v)}. " if v else ""
    xl_inputs.setting(ws, m, K["options"], "Options", "Stock options outstanding", d.get("options"), "shares",
                      filed(d.get("options"), "{:,.0f}") + "In full. Blank = no dilution from options",
                      reg=("Inputs", "Stock options outstanding"))
    strike = f"Average exercise price ({'$' if not foreign else cur})"
    xl_inputs.setting(ws, m, K["strike"], "Option strike", strike, d.get("strike"), "price",
                      filed(d.get("strike"), "{:,.2f}") + "Options only dilute to the extent the share price is "
                                                          "above this", reg=("Inputs", strike))
    xl_inputs.setting(ws, m, K["rsus"], "RSUs", "Unvested stock units (RSUs / PSUs)", d.get("rsus"), "shares",
                      filed(d.get("rsus"), "{:,.0f}") + "In full, counted one for one. Usually the larger item for "
                                                        "tech companies",
                      reg=("Inputs", "Unvested stock units (RSUs / PSUs)"))
    calc("price", "Share price" + ("" if not foreign else f" ({cur} per ordinary share)"),
         f"=N({last.market('Share price')})", "price", "Inputs tab")
    calc("itm", "Shares added by options (treasury method)",
         f'=IF(AND(N({B("options")})>0,{B("price")}>0),MAX(0,N({B("options")})*(1-N({B("strike")})/{B("price")})),0)',
         "shares", "Options x (1 - exercise price / share price): the exercise money buys back part of the shares")
    calc("diluted", "Diluted shares used in the valuation",
         f'=IF(N({B("shares")})=0,"",{B("shares")}+N({B("rsus")})+{B("itm")})', "shares",
         "Shares outstanding + stock units + options in the money. Every value per share divides by this", bold=True)

    # ---- enterprise value to equity
    st.section(ws, K["ev_title"], 1, 6, "From enterprise value to equity value")
    xl_inputs.override_header(ws, K["ev_hdr"])
    calc("bs_date", "Balance sheet date", bs_col.end.date(), "date",
         ("Latest 10-Q" if bs_col.kind == "ltm" else "Last fiscal year end (no newer 10-Q)")
         + ": the source of the figures below")
    ws.cell(row=K["bs_date"], column=2).font = st.font()
    calc("borrow", "Borrowings", f"={bs.n('Short-Term Borrowings')}+{bs.n('Current Portion of LT Debt')}"
                                 f"+{bs.n('Long-Term Debt')}", "usd",
         "Short-term borrowings + current portion of long-term debt + long-term debt")
    calc("leases", "Operating lease liabilities", f"={bs.n('Operating Lease Liabilities')}", "usd")
    ifrs = fin.taxonomy == "ifrs-full"
    xl_inputs.setting(ws, m, K["lease_mode"], "lease_mode", "Lease liabilities",
                      LEASE_MODES[1] if ifrs else LEASE_MODES[0], "num",
                      "IFRS filers: lease payments are outside operating cash flow, so leases count as debt (default "
                      "for them). US GAAP: operating lease cost is already in EBIT, so they don't", LEASE_MODES)
    calc("debt", "Total debt", f'={B("borrow")}+IF({B("lease_mode")}="{LEASE_MODES[1]}",{B("leases")},0)', "usd",
         "Borrowings, plus lease liabilities when the choice above says so. Also the debt in the WACC's weights",
         bold=True)
    calc("cash", "Cash & investments", f"={bs.n('Cash & Equivalents')}+{bs.n('Short-Term Investments')}"
                                       f"+{bs.n('Long-Term Investments')}", "usd",
         "Cash + short-term and long-term investments")
    calc("net_debt", "Net debt", f"={B('debt')}-{B('cash')}", "usd")
    calc("nci", "Noncontrolling & redeemable interests",
         f"={bs.n('Noncontrolling Interest')}+{bs.n('Redeemable NCI & Mezzanine Equity')}", "usd",
         "Other owners' share of subsidiaries: theirs, not common shareholders'")
    xl_inputs.setting(ws, m, K["other_assets"], "other_assets", "Other non-operating assets", None, "usd",
                      "Optional, in full (e.g. 5000000000): assets the forecast cash flows don't capture and net "
                      "debt doesn't count, such as equity stakes. Added to equity value")
    xl_inputs.setting(ws, m, K["other_claims"], "other_claims", "Other claims on the business", None, "usd",
                      "Optional, in full: obligations net debt misses, such as pension deficits or preferred stock. "
                      "Deducted from equity value")
    c = calc("claims", "Net claims ahead of common equity",
             f"={B('net_debt')}+{B('nci')}+N({B('other_claims')})-N({B('other_assets')})", "usd",
             "Net debt + other owners' interests + other claims - other non-operating assets. Equity value = "
             "enterprise value - this", bold=True)
    c.fill = st.PatternFill("solid", fgColor=st.THEME["pale"])

    # ---- at today's price
    st.section(ws, K["now_title"], 1, 6, "At today's share price")
    calc("mcap", "Equity value (diluted shares x price)", f'=IF({B("diluted")}="","",{B("price")}*{B("diluted")})',
         "usd")
    calc("ev", "Enterprise value (equity value + net claims)", f'=IF({B("mcap")}="","",{B("mcap")}+{B("claims")})',
         "usd", "What the market pays for the business: compare with the DCF's enterprise value", bold=True)

    st.row_rules(ws, [K[k] for k in ("shares", "price", "itm", "diluted", "bs_date", "borrow", "leases", "debt",
                                      "cash", "net_debt", "nci", "claims", "mcap", "ev")], 6)
    for col, w in zip("ABCDEF", (46, 16, 16, 16, 30, 90)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "B4"
