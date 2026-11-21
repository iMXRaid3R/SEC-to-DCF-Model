"""The menu shown when the script runs with no arguments (./run, .\\run.bat in PowerShell, or double-clicking run.bat).

Each choice is turned into the same command-line arguments you could type yourself, and the command is shown
before it runs, so the menu doubles as a way to learn the shortcuts."""
import json
from pathlib import Path

from . import parts as P
from . import sec

ROOT = Path(__file__).resolve().parent.parent
CHOICES = [("1", "Build or update a company"), ("2", "Build every company in a watchlist group"),
           ("3", "Portfolio roll-up (reads every workbook through Excel)"),
           ("4", "One-page write-up of a company (through Excel)"),
           ("5", "What changed since an older version of a workbook"),
           ("6", "Setup: SEC contact email and watchlist"), ("q", "Quit")]


def _ask(prompt, default=""):
    try:
        got = input(prompt).strip()
    except EOFError:
        raise KeyboardInterrupt
    return got or default


def _yes(prompt, default=True):
    got = _ask(prompt + (" [Y/n] " if default else " [y/N] ")).lower()
    return default if not got else got.startswith("y")


def _watchlist():
    for name in ("watchlist.json", "watchlist.example.json"):
        f = ROOT / name
        if f.exists():
            try:
                return json.loads(f.read_text(encoding="utf-8"))
            except ValueError:
                return {}
    return {}


def _number(prompt, default, bounds):
    """A whole number within bounds; Enter keeps the default, anything else asks again."""
    lo, hi = bounds
    while True:
        got = _ask(prompt)
        if not got:
            return default
        if got.isdigit() and lo <= int(got) <= hi:
            return int(got)
        print(f"  A whole number from {lo} to {hi}, please.")


def pick_parts(current, peers):
    """Toggle list. Returns the chosen set of parts, or None to go back."""
    chosen = set(current)
    while True:
        print(f"\nAlways built: {P.ALWAYS}. Optional parts:")
        for i, part in enumerate(P.PARTS, start=1):
            extra = ""
            if part == "comps":
                extra = f"   peers: {', '.join(peers)}" if peers else "   (no peers listed in watchlist.json: skipped)"
            print(f"  [{'x' if part in chosen else ' '}] {i}  {P.NAMES[part]}{extra}")
        got = _ask("Numbers switch a part on/off, s = core model only, a = all, b = back, Enter = build: ").lower()
        if not got:
            return chosen
        if got == "b":
            return None
        if got == "s":
            chosen = set()
        elif got == "a":
            chosen = set(P.PARTS)
        else:
            for ch in got.replace(",", " ").split():
                if ch.isdigit() and 1 <= int(ch) <= len(P.PARTS):
                    chosen ^= {P.PARTS[int(ch) - 1]}


def _run(main, args):
    print("\n> ./run " + " ".join(f'"{x}"' if " " in x else x for x in args) + "\n")
    try:
        main(args)
    except SystemExit as e:                     # argparse errors and sys.exit("message") end up here
        if e.code not in (None, 0):
            print(e.code if isinstance(e.code, str) else f"(stopped with code {e.code})")
        return False
    return True


def menu(main):
    print("EDGAR financial models")
    while True:
        email = [] if sec.DEFAULT_EMAIL else (["--email", sec._default_email()] if sec._default_email() else [])
        if not sec.DEFAULT_EMAIL and not email:
            print("\nFirst run: the SEC asks for a contact email before anything can be downloaded.")
            _run(main, ["--setup"])
            continue
        print()
        for key, text in CHOICES:
            print(f"  {key}  {text}")
        choice = ""
        try:
            choice = _ask("Choose: ").lower()
            if choice in ("q", "quit", "exit"):
                return
            wl = _watchlist()
            if choice == "1":
                ticker = _ask("Ticker (e.g. NVDA): ").upper()
                if not ticker:
                    continue
                path = ROOT / f"{ticker}_financials.xlsx"
                default = P.previous(path) if path.exists() else None
                dp = wl.get("default_parts")
                if default is None:
                    default = P.parse(dp if isinstance(dp, str) else ",".join(dp)) if dp else set(P.PARTS)
                peers = {k.upper(): v for k, v in wl.get("peers", {}).items()}.get(ticker, [])
                if path.exists():
                    print(f"{path.name} exists: your typed inputs are kept and the old file goes to backups/.")
                chosen = pick_parts(default, peers)
                if chosen is None:
                    continue
                hist, fcst = P.years(remembered=P.previous_years(path) if path.exists() else (None, None),
                                     default=(wl.get("default_history_years"), wl.get("default_forecast_years")))
                hist = _number(f"Years of history ({P.HISTORY[1][0]}-{P.HISTORY[1][1]}, Enter = {hist}): ", hist,
                               P.HISTORY[1])
                fcst = _number(f"Years of forecast ({P.FORECAST[1][0]}-{P.FORECAST[1][1]}, Enter = {fcst}): ", fcst,
                               P.FORECAST[1])
                args = [ticker, "--parts", ",".join(sorted(chosen)) or "statements",
                        "--years", str(hist), "--forecast-years", str(fcst)] + email
                if _run(main, args) and (ROOT / path.name).exists() and _yes("Open it in Excel?"):
                    main.open_file(path)
            elif choice == "2":
                groups = {k: v for k, v in wl.items() if isinstance(v, list)}
                if not groups:
                    print("No groups in watchlist.json yet (run Setup, then edit watchlist.json).")
                    continue
                print("Groups: " + ", ".join(f"{k} ({len(v)})" for k, v in groups.items()) + ", all")
                g = _ask("Which group? ", next(iter(groups)))
                print("Each workbook keeps the parts it was built with; new ones get everything.")
                _run(main, ["--list", g] + email)
            elif choice == "3":
                base = _yes("Base scenario only? (about a third of the time, no Bear/Bull)", False)
                _run(main, ["--rollup"] + (["--base-only"] if base else []))
            elif choice == "4":
                tickers = _ask("Ticker(s), separated by spaces: ").upper().split()
                if tickers:
                    _run(main, ["--writeup"] + tickers)
            elif choice == "5":
                ticker = _ask("Ticker: ").upper()
                old = sorted((ROOT / "backups").glob(f"{ticker}_financials_*.xlsx"), reverse=True)[:8]
                if not old:
                    print(f"No earlier versions of {ticker}_financials.xlsx in backups/ yet.")
                    continue
                for i, f in enumerate(old, start=1):
                    print(f"  {i}  {f.name}")
                k = _ask("Compare today's SEC data with which one? [1] ", "1")
                if k.isdigit() and 1 <= int(k) <= len(old):
                    _run(main, [ticker, "--since", str(old[int(k) - 1].relative_to(ROOT))] + email)
            elif choice == "6":
                _run(main, ["--setup"])
                sec.DEFAULT_EMAIL = sec._default_email()
            else:
                print("Type one of the numbers above, or q.")
        except KeyboardInterrupt:          # Ctrl+C backs out of the current step, at the menu it quits
            print()
            if not choice:
                return
        except Exception as e:                 # keep the menu alive whatever one command does
            print(f"\nThat stopped with an error: {type(e).__name__}: {e}")
