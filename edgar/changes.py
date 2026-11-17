"""What changed since the last build.

Every workbook carries the SEC figures it was built from on its Raw tab. On a rebuild the previous Raw tab is
compared with the new data, so after a 10-Q or 10-K lands you get a short list of what moved:
    the newly reported period against the same period a year earlier (new YTD less old YTD),
    trailing-twelve-month figures and the ratios the driver defaults are taken from, before and after,
    the balance sheet, and any earlier-year figures that were restated or newly filled.
Printed after the build and saved to changes/<TICKER>_<date>.md when a new period arrived. Values per share
need Excel to calculate, so their movement is reported by the roll-up (see rollup.value_changes)."""
from datetime import date
from pathlib import Path

import pandas as pd

YTD_MARK, LTM_MARKS = "[fiscal year to date]", ("[LTM: FY + YTD - prior YTD]", "[10-Q]")
FLOWS = ("Revenue", "Net Interest Income", "Provision for Credit Losses", "Gross Profit", "Operating Income",
         "Pre-Tax Income", "Net Income", "Cash from Operations", "Free cash flow")
RATIOS = (  # (name, numerator, denominator)
    ("Gross margin", "Gross Profit", "Revenue"), ("R&D % of revenue", "R&D Expense", "Revenue"),
    ("SG&A % of revenue", "SG&A Expense", "Revenue"), ("Operating margin", "Operating Income", "Revenue"),
    ("Tax rate", "Income Tax Expense", "Pre-Tax Income"), ("Net margin", "Net Income", "Revenue"),
    ("Capex % of revenue", "Capex", "Revenue"), ("SBC % of revenue", "Stock-Based Comp", "Revenue"),
    ("FCF margin", "Free cash flow", "Revenue"))
BALANCE = (("Cash and investments", ("Cash & Equivalents", "Short-Term Investments", "Long-Term Investments")),
           ("Debt", ("Short-Term Borrowings", "Current Portion of LT Debt", "Long-Term Debt")),
           ("Deposits", ("Deposits",)), ("Total assets", ("Total Assets",)),
           ("Shareholders' equity", ("Shareholders' Equity (Parent)",)))


class Snapshot:
    """The reported figures behind one build: annual values, the latest 10-Q's trailing-twelve-month and
    year-to-date values, and when the newest of them was filed."""

    def __init__(self, src):
        self.annual, self.ltm, self.ytd = {}, {}, {}
        self.q_end = self.fy_end = self.filed = None
        if src is None or src.empty:
            return
        latest_filed = {}
        for r in src.itertuples(index=False):
            label, end, tag, val = r[0], pd.Timestamp(r[1]), str(r[2]), r[3]
            if not isinstance(val, (int, float)) or pd.isna(val):
                continue
            if tag.endswith(YTD_MARK):
                self.ytd[label] = float(val)
            elif tag.endswith(LTM_MARKS):
                self.ltm[label] = float(val)
                self.q_end = end if self.q_end is None else max(self.q_end, end)
            else:
                self.annual.setdefault(label, {})[end] = float(val)
                continue
            latest_filed[end] = max(latest_filed.get(end, pd.Timestamp(r[6])), pd.Timestamp(r[6]))
        flows = [e for label in ("Revenue", "Net Income", "Pre-Tax Income") for e in self.annual.get(label, {})]
        ends = flows or [e for series in self.annual.values() for e in series]
        self.fy_end = max(ends) if ends else None
        if self.q_end is not None:
            self.filed = latest_filed.get(self.q_end)
        elif self.fy_end is not None:
            dates = [pd.Timestamp(r[6]) for r in src.itertuples(index=False) if pd.Timestamp(r[1]) == self.fy_end]
            self.filed = max(dates) if dates else None

    @property
    def end(self):
        return self.q_end if self.q_end is not None else self.fy_end

    @property
    def label(self):
        if self.q_end is not None:
            return f"LTM {self.q_end:%b-%y}"
        return f"FY to {self.fy_end:%b-%y}" if self.fy_end is not None else "no data"

    def ttm(self, label):
        """Trailing twelve months of a flow line, or the latest balance of a balance-sheet line."""
        if self.q_end is not None:
            return _derived(self.ltm.get, label)
        return _derived(lambda k: self.annual.get(k, {}).get(self.fy_end), label)

    def cumulative(self, label, since):
        """Flow reported since fiscal-year end `since` (the other snapshot's last full year)."""
        ytd = _derived(self.ytd.get, label) if self.q_end is not None else 0.0
        if self.fy_end == since:
            return ytd
        full = _derived(lambda k: self.annual.get(k, {}).get(self.fy_end), label)
        return None if full is None or ytd is None else full + ytd


def _derived(get, label):
    """A line's value, building the few that aren't reported as such from the ones that are."""
    v = get(label)
    if v is not None:
        return v
    if label == "Gross Profit":
        rev, cost = get("Revenue"), get("Cost of Revenue")
        return rev - cost if rev is not None and cost is not None else None
    if label == "Free cash flow":
        cfo, capex = get("Cash from Operations"), get("Capex")
        return cfo - abs(capex) if cfo is not None and capex is not None else None
    return None


def read_raw(path):
    """The Raw tab of a built workbook as a DataFrame (None if there isn't one)."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if "Raw" not in wb.sheetnames:
            return None
        rows = list(wb["Raw"].iter_rows(values_only=True))
    finally:
        wb.close()
    if len(rows) < 2 or rows[0][0] != "Line Item":
        return None
    return pd.DataFrame([r[:8] for r in rows[1:] if r and r[0]], columns=list(rows[0][:8]))


def _months(a, b):
    return round((b - a).days / 30.4)


def compare(old, new, currency="USD"):
    """Sections [(title, header, rows)] describing what moved between two Snapshots; [] if nothing did."""
    if old.end is None or new.end is None:
        return []
    out, unit = [], ("$" if currency == "USD" else currency + " ")
    money = lambda v: "" if v is None else f"{v / 1e6:,.0f}"
    pct = lambda v: "" if v is None else f"{v:+.1%}"
    share = lambda v: "" if v is None else f"{v:.1%}"
    pts = lambda a, b: "" if a is None or b is None else f"{(b - a) * 100:+.1f} pts"
    growth = lambda a, b: (b / a - 1) if a and b is not None and a > 0 and b > 0 else None
    ratio = lambda get, n, d: (get(n) / get(d)) if get(n) is not None and get(d) else None

    months = _months(old.end, new.end)
    if months > 0 and new.fy_end is not None and old.fy_end is not None \
            and 0 <= (new.fy_end - old.fy_end).days <= 380 and months <= 12:
        # the period reported since the last build, and the same months a year earlier:
        # trailing-12-month change = new period - year-ago period
        def period(label):
            a, b = old.cumulative(label, old.fy_end), new.cumulative(label, old.fy_end)
            t0, t1 = old.ttm(label), new.ttm(label)
            if None in (a, b, t0, t1):
                return None, None
            now = b - a
            return now, now - (t1 - t0)
        span = "Quarter" if months == 3 else f"{months} months"
        rows, got = [], {}
        for label in FLOWS:
            now, ago = period(label)
            if now is None:
                continue
            got[label] = (now, ago)
            rows.append([label, money(now), money(ago), pct(growth(ago, now))])
        ni, sh = got.get("Net Income"), new.ytd.get("Diluted Shares") or new.ttm("Diluted Shares")
        e0, e1 = old.ttm("Diluted EPS"), new.ttm("Diluted EPS")
        if ni and sh and e0 is not None and e1 is not None:
            eps = ni[0] / sh
            ago = eps - (e1 - e0)
            rows.append(["Diluted EPS (approx.)", f"{eps:.2f}", f"{ago:.2f}", pct(growth(ago, eps))])
        for name, n, d in RATIOS[:1] + RATIOS[3:4]:
            if n in got and "Revenue" in got and got["Revenue"][0] and got["Revenue"][1]:
                a, b = got[n][1] / got["Revenue"][1], got[n][0] / got["Revenue"][0]
                rows.append([name, share(b), share(a), pts(a, b)])
        if rows:
            out.append((f"{span} ended {new.end:%b %d, %Y} ({unit.strip()} millions)",
                        ["", f"{span} to {new.end:%b-%y}", "Same period a year earlier", "Change"], rows))

    if months > 0 or old.ltm != new.ltm:
        rows = []
        for label in ("Revenue", "Net Interest Income", "Net Income", "Free cash flow"):
            a, b = old.ttm(label), new.ttm(label)
            if a is not None and b is not None:
                rows.append([label, money(a), money(b), pct(growth(a, b))])
        a, b = old.ttm("Diluted EPS"), new.ttm("Diluted EPS")
        if a is not None and b is not None:
            rows.append(["Diluted EPS", f"{a:.2f}", f"{b:.2f}", pct(growth(a, b))])
        for name, n, d in RATIOS:
            a, b = ratio(old.ttm, n, d), ratio(new.ttm, n, d)
            if a is not None and b is not None:
                rows.append([name, share(a), share(b), pts(a, b)])
        rows = [r for r in rows if months > 0 or r[1] != r[2]]
        if rows:
            out.append((f"Trailing twelve months ({unit.strip()} millions): what the driver defaults are taken from",
                        ["", old.label, new.label, "Change"], rows))
        rows = []
        for name, parts in BALANCE:
            a = [old.ttm(p) for p in parts]
            b = [new.ttm(p) for p in parts]
            if any(x is not None for x in a) and any(x is not None for x in b):
                a, b = sum(x or 0 for x in a), sum(x or 0 for x in b)
                if months > 0 or a != b:
                    rows.append([name, money(a), money(b), f"{(b - a) / 1e6:+,.0f}" if b != a else "unchanged"])
        a, b = old.ttm("Diluted Shares"), new.ttm("Diluted Shares")
        if a and b and (months > 0 or a != b):
            rows.append(["Diluted shares (millions)", f"{a / 1e6:,.1f}", f"{b / 1e6:,.1f}", pct(b / a - 1)])
        if rows:
            out.append((f"Balance sheet ({unit.strip()} millions)",
                        ["", f"{old.end:%b-%y}", f"{new.end:%b-%y}", "Change"], rows))

    # earlier years: restated figures, and lines that gained or lost a value (new tags, overrides)
    rows = []
    for label in sorted(set(old.annual) | set(new.annual)):
        a, b = old.annual.get(label, {}), new.annual.get(label, {})
        for end in sorted(set(a) & set(b)):
            if abs(b[end] - a[end]) > max(0.005 * abs(a[end]), 1e-9):
                per_share = abs(a[end]) < 1e4
                f = (lambda v: f"{v:.2f}") if per_share else money
                rows.append([f"{label}, FY to {end:%b-%y}", f(a[end]), f(b[end]), pct(growth(abs(a[end]), abs(b[end])))])
    if rows:
        more = len(rows) - 12
        rows = rows[:12] + ([[f"... and {more} more", "", "", ""]] if more > 0 else [])
        out.append(("Earlier years restated", ["", "Was", "Now", "Change"], rows))
    common = set(old.annual.get("Revenue", {})) & set(new.annual.get("Revenue", {}))
    has = lambda s, label: any(e in common for e in s.annual.get(label, {}))
    gained = sorted(k for k in new.annual if has(new, k) and not has(old, k))
    lost = sorted(k for k in old.annual if has(old, k) and not has(new, k))
    if common and (gained or lost):
        rows = [["Now filled", ", ".join(gained)]] * bool(gained) + [["Now blank", ", ".join(lost)]] * bool(lost)
        out.append(("Lines that changed", ["", ""], rows))
    return out


def headline(old, new, annual_form="10-K"):
    if new.end is None or old.end is None:
        return ""
    filed = f", filed {new.filed:%b %d, %Y}" if new.filed is not None else ""
    if new.fy_end is not None and old.fy_end is not None and new.fy_end > old.fy_end:
        what = f"New annual report: fiscal year to {new.fy_end:%b %d, %Y}"
        if new.q_end is not None:
            what += f", and a quarterly report to {new.q_end:%b %d, %Y}"
        return what + filed + ". The three-year averages behind the driver defaults have moved on a year."
    if new.end > old.end:
        return f"New quarterly report: period to {new.end:%b %d, %Y}{filed} (the last build ran to {old.end:%b %d, %Y})."
    return "No new period since the last build, but some reported figures differ."


def as_text(title, sections):
    lines = [title]
    for name, header, rows in sections:
        table = [header] + rows
        widths = [max(len(str(r[i])) for r in table) for i in range(len(header))]
        lines += ["", "  " + name]
        for r in table:
            lines.append("    " + str(r[0]).ljust(widths[0]) + "".join("  " + str(c).rjust(w) for c, w in zip(r[1:], widths[1:])))
    return "\n".join(lines)


def as_markdown(ticker, name, title, sections, today):
    lines = [f"# {name} ({ticker}): what changed", "", f"*Built {today:%B %d, %Y}.* {title}"]
    for sec, header, rows in sections:
        lines += ["", f"## {sec}", "", "| " + " | ".join(header) + " |",
                  "|---|" + "---:|" * (len(header) - 1)]
        lines += ["| " + " | ".join(str(c) for c in r) + " |" for r in rows]
    lines += ["", "---", "", "*Your typed inputs were kept as they were: check the Drivers tab against the figures "
              "above, then run the roll-up to see how the value per share moved.*"]
    return "\n".join(lines) + "\n"


def report(fin, old_path, folder=None, today=None):
    """Compare the workbook at old_path with freshly loaded `fin`. Prints the report; saves it under
    <folder>/changes/ when a new period arrived. Returns "new", "revised" or None."""
    try:
        raw = read_raw(old_path)
    except Exception as e:                      # a damaged or foreign workbook must never stop a build
        print(f"  (no change report: could not read {Path(old_path).name}: {type(e).__name__})")
        return None
    if raw is None:
        return None
    old, new = Snapshot(raw), Snapshot(fin.sources)
    if old.end is None or new.end is None:
        return None
    if new.end < old.end:
        print(f"  ! The previous workbook ran to {old.end.date()} but this data stops at {new.end.date()}: "
              "the cache may be older than the workbook. Run with --refresh.")
        return None
    sections = compare(old, new, fin.currency)
    if not sections:
        print(f"No change in the reported figures since the last build (latest period: {new.end.date()}).")
        return None
    title = headline(old, new, fin.annual_form)
    print("\n" + as_text("What changed since the last build. " + title, sections))
    kind = "new" if new.end > old.end else "revised"
    if kind == "new" and folder is not None:
        today = today or date.today()
        out = Path(folder) / "changes"
        out.mkdir(exist_ok=True)
        path = out / f"{fin.ticker}_{today}.md"
        path.write_text(as_markdown(fin.ticker, fin.name, title, sections, today), encoding="utf-8")
        print(f"\n  Saved changes/{path.name}. Your typed inputs were kept: check the Drivers tab against these "
              "figures.")
    return kind
