---
name: simple-channel-breakout-backtest
description: Backtests a simple long-only channel breakout strategy (Donchian-style) on one stock or ETF from a CSV or Excel price file and produces an HTML tearsheet (optional PDF). Use only for channel breakout tests - not for moving average, momentum or other strategies.
---

# Simple Channel Breakout Backtest

Runs one fixed strategy: a long-only channel breakout on a single security. The backtest math lives in the bundled Python script below, so every run uses identical rules. Never change the rules, add optimisation, or test other parameter sets unless the user explicitly asks for a separate run.

## Step 1 - Get the price file
- The user points to (Claude Code) or attaches (chat) a `.csv` or `.xlsx` file. If none is given, ask for it.
- The file MUST contain the columns `date`, `open`, `high`, `low`, `close` (case-insensitive, extra columns are fine and ignored). If any of the five is missing, stop and tell the user which ones are missing. Do not guess or rebuild missing columns.
- Dates must be dd/mm/yyyy (native Excel date cells are also fine). All dates in the output are shown as dd/mm/yyyy.
- If an Excel workbook has several sheets, ask which sheet to use.

## Step 2 - Ask for the parameters (always ask, never assume)
Ask for all of these in one message (use the multiple-choice question tool if available). Only commission and slippage have defaults.
1. **Entry basis** - break above the highest **high** or the highest **close** of the previous X bars?
2. **X** - entry lookback in bars (whole number).
3. **Exit basis** - break below the lowest **low** or the lowest **close** of the previous Y bars?
4. **Y** - exit lookback in bars (whole number).
5. **Starting capital**.
6. **Commission** - flat amount per order, charged on entry and on exit. Default 0.
7. **Slippage** - percent of the fill price, against you on every fill. Default 0%.
8. **Output** - HTML only (default) or HTML + PDF.
Optional: a label for the report (e.g. the ticker); otherwise the file name is used.

Fixed rules (do not ask, state them in the confirmation): long only; 100% of available equity per trade in whole shares only (leftover cash stays uninvested; an entry that cannot afford one share is skipped and flagged); signal on the bar's close, fill at the next bar's open; channels use the previous X / Y bars only (the signal bar is excluded); an open position at the end is marked at the last close; benchmark is buy-and-hold on the same file, bought at the first bar's open with the same costs.

Repeat the chosen settings back in one short line, then run.

## Step 3 - Run the script
1. Check `pandas`, `numpy`, `matplotlib` and `openpyxl` are importable; install any that are missing (`pip install <pkg>`, add `--break-system-packages` if pip asks for it).
2. If `scb_backtest.py` is not already saved next to this SKILL.md (or in the working directory), write the script from the "Bundled script" section below to `scb_backtest.py` **exactly as written** - do not edit or "improve" it.
3. Run:
```
python scb_backtest.py --file "<path>" [--sheet "<sheet>"] [--name "<label>"] \
  --entry-basis high|close --x <X> --exit-basis low|close --y <Y> \
  --capital <capital> --commission <commission> --slippage <slippage_pct> [--pdf] --outdir <output folder>
```
   `--slippage` is in percent: `0.05` means 0.05%.
4. The script prints JSON. If `"status": "error"`, show the user the message in plain words and stop - do not work around it.

## Step 4 - Present the results
- Deliver the HTML tearsheet (and the PDF if requested, plus the trades CSV). In Claude Code, say where the files were saved; in chat, share the files.
- In the reply, give 3-5 lines: CAGR, max drawdown, Sharpe and number of trades for the strategy vs buy-and-hold, plus any data warnings. Do not re-list every metric; the tearsheet has them.
- Do not suggest "better" parameters or run extra variations unless asked.

## Tearsheet contents (produced by the script)
Header with label and generation date; strategy rules as selected; performance vs buy-and-hold (final equity, total return, CAGR, annualised volatility, Sharpe, Sortino, max drawdown, Calmar, longest drawdown); trade statistics (closed trades, open trade flag, win rate, profit factor, average trade/win/loss, best/worst trade, average bars held, exposure); equity curve vs buy-and-hold; drawdown chart; price chart with entry/exit channels and trade markers; monthly returns heatmap with yearly totals; full trade list. Ratios use a zero risk-free rate, annualised from the data frequency (daily/weekly/monthly detected automatically).

## Bundled script - scb_backtest.py
```python
#!/usr/bin/env python3
"""Simple channel breakout backtest -> HTML tearsheet (+ optional PDF).

Long only, single ticker, 100% of available equity per trade (whole shares only).
Signals on a bar's close, fills at the next bar's open.
"""
import argparse, base64, io, json, os, sys
from datetime import datetime

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

DF = "%d/%m/%Y"  # every date shown to the user uses dd/mm/yyyy
REQUIRED = ["date", "open", "high", "low", "close"]


def fail(msg):
    print(json.dumps({"status": "error", "message": msg}))
    sys.exit(1)


# ---------------------------------------------------------------- data
def load(path, sheet):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        xl = pd.ExcelFile(path)
        if sheet is None:
            if len(xl.sheet_names) > 1:
                fail("Workbook has several sheets: %s. Re-run with --sheet." % xl.sheet_names)
            sheet = xl.sheet_names[0]
        raw = xl.parse(sheet)
    elif ext in (".csv", ".txt"):
        raw = pd.read_csv(path, dtype=str)  # read as text so dates are never guessed
    else:
        fail("Unsupported file type '%s'. Use .csv or .xlsx." % ext)

    cols = {str(c).strip().lower(): c for c in raw.columns}
    missing = [c for c in REQUIRED if c not in cols]
    if missing:
        fail("Missing required column(s): %s. Found: %s. The file must contain date, open, high, low, close."
             % (", ".join(missing), [str(c) for c in raw.columns]))
    df = pd.DataFrame({c: raw[cols[c]] for c in REQUIRED})

    def to_date(v):
        if isinstance(v, (pd.Timestamp, datetime)):
            return pd.Timestamp(v)
        try:
            return pd.Timestamp(datetime.strptime(str(v).strip()[:10], DF))
        except Exception:
            return pd.NaT
    df["date"] = df["date"].map(to_date)
    bad = df.index[df["date"].isna()].tolist()
    if bad:
        fail("Could not read dates as dd/mm/yyyy in %d row(s), e.g. file rows %s." % (len(bad), [i + 2 for i in bad[:10]]))
    for c in REQUIRED[1:]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    bad = df.index[df[REQUIRED[1:]].isna().any(axis=1)].tolist()
    if bad:
        fail("Non-numeric or blank prices in %d row(s), e.g. file rows %s." % (len(bad), [i + 2 for i in bad[:10]]))
    if df["date"].duplicated().any():
        fail("Duplicate dates found, e.g. %s." % df.loc[df["date"].duplicated(), "date"].dt.strftime(DF).tolist()[:5])
    df = df.sort_values("date").reset_index(drop=True)
    warn = []
    odd = ((df.high < df[["open", "close"]].max(axis=1)) | (df.low > df[["open", "close"]].min(axis=1))).sum()
    if odd:
        warn.append("%d bar(s) where high/low do not contain open/close - check the data." % odd)
    return df, warn


# ---------------------------------------------------------------- engine
def run(df, p):
    o, h, l, c = (df[k].to_numpy(float) for k in ("open", "high", "low", "close"))
    n = len(df)
    ent_src = df["high"] if p.entry_basis == "high" else df["close"]
    ext_src = df["low"] if p.exit_basis == "low" else df["close"]
    upper = ent_src.rolling(p.x).max().shift(1).to_numpy()   # previous X bars, current bar excluded
    lower = ext_src.rolling(p.y).min().shift(1).to_numpy()   # previous Y bars, current bar excluded
    s, comm = p.slippage / 100.0, p.commission

    cash, shares, pending, skipped = p.capital, 0, None, 0
    eq, pos, trades, cur = np.zeros(n), np.zeros(n, bool), [], None
    for t in range(n):
        if pending == "buy":
            px = o[t] * (1 + s); qty = int(np.floor((cash - comm) / px)) if cash > comm else 0
            if qty >= 1:  # whole shares only; leftover cash stays uninvested
                cost = qty * px + comm; cash -= cost; shares = qty
                cur = dict(signal_date=df.date[t - 1], entry_date=df.date[t], entry_price=px,
                           shares=shares, cost=cost, entry_i=t)
            else:
                skipped += 1
        elif pending == "sell":
            px = o[t] * (1 - s); proceeds = shares * px - comm; cash += proceeds
            cur.update(exit_signal_date=df.date[t - 1], exit_date=df.date[t], exit_price=px,
                       proceeds=proceeds, bars=t - cur["entry_i"], status="Closed")
            trades.append(cur); shares, cur = 0, None
        pending = None
        eq[t] = cash + shares * c[t]; pos[t] = shares > 0
        if t < n - 1:
            if shares == 0 and not np.isnan(upper[t]) and c[t] > upper[t]:
                pending = "buy"
            elif shares > 0 and not np.isnan(lower[t]) and c[t] < lower[t]:
                pending = "sell"
    if cur is not None:  # still open at the end: mark to last close, no exit costs
        cur.update(exit_signal_date=pd.NaT, exit_date=df.date[n - 1], exit_price=c[-1],
                   proceeds=shares * c[-1], bars=n - 1 - cur["entry_i"], status="Open (marked at last close)")
        trades.append(cur)
    for tr in trades:
        tr["pnl"] = tr["proceeds"] - tr["cost"]; tr["ret"] = tr["pnl"] / tr["cost"]

    # buy & hold: buy at first bar open with the same costs, hold to the end
    bpx = o[0] * (1 + s); bh_sh = int(np.floor((p.capital - comm) / bpx))
    if bh_sh < 1: fail("Starting capital is too small to buy even one share at the first bar's open (%.2f)." % bpx)
    bh = (p.capital - comm - bh_sh * bpx) + bh_sh * c
    return pd.Series(eq, df.date), pd.Series(bh, df.date), pos, trades, upper, lower, skipped


# ---------------------------------------------------------------- metrics
def periods_per_year(dates):
    d = dates.diff().dt.days.median()
    return 252 if d <= 4 else 52 if d <= 10 else 12 if d <= 40 else 4 if d <= 120 else 1


def stats(eq, cap, ppy):
    r = eq.pct_change().fillna(eq.iloc[0] / cap - 1)
    yrs = max((eq.index[-1] - eq.index[0]).days / 365.25, 1e-9)
    tot = eq.iloc[-1] / cap - 1
    cagr = (eq.iloc[-1] / cap) ** (1 / yrs) - 1 if eq.iloc[-1] > 0 else -1
    sd = r.std(); dn = np.sqrt((np.minimum(r, 0) ** 2).mean())
    peak = np.maximum.accumulate(np.r_[cap, eq.to_numpy()])[1:]
    dd = eq / peak - 1
    # longest drawdown duration (calendar days, peak to recovery or end)
    longest, start = 0, None
    for d, v in dd.items():
        if v < 0 and start is None: start = d
        if v >= 0 and start is not None: longest = max(longest, (d - start).days); start = None
    if start is not None: longest = max(longest, (dd.index[-1] - start).days)
    mdd = dd.min()
    return dict(total_return=tot, cagr=cagr, ann_vol=sd * np.sqrt(ppy),
                sharpe=r.mean() / sd * np.sqrt(ppy) if sd > 0 else np.nan,
                sortino=r.mean() / dn * np.sqrt(ppy) if dn > 0 else np.nan,
                max_dd=mdd, calmar=cagr / abs(mdd) if mdd < 0 else np.nan,
                longest_dd_days=longest, final_equity=eq.iloc[-1]), dd


def trade_stats(trades, pos):
    closed = [t for t in trades if t["status"] == "Closed"]
    rets = np.array([t["ret"] for t in closed]); pnl = np.array([t["pnl"] for t in closed])
    w, lo = pnl[pnl > 0], pnl[pnl <= 0]
    return dict(n_trades=len(closed), open_trade=len(trades) - len(closed),
                win_rate=(pnl > 0).mean() if len(pnl) else np.nan,
                profit_factor=w.sum() / abs(lo.sum()) if lo.sum() < 0 else (np.inf if len(w) else np.nan),
                avg_trade=rets.mean() if len(rets) else np.nan,
                avg_win=rets[rets > 0].mean() if (rets > 0).any() else np.nan,
                avg_loss=rets[rets <= 0].mean() if (rets <= 0).any() else np.nan,
                best=rets.max() if len(rets) else np.nan, worst=rets.min() if len(rets) else np.nan,
                avg_bars=np.mean([t["bars"] for t in closed]) if closed else np.nan,
                exposure=pos.mean())


def monthly_table(eq, cap):
    try: m = eq.resample("ME").last()
    except ValueError: m = eq.resample("M").last()
    r = m.pct_change(); r.iloc[0] = m.iloc[0] / cap - 1
    tab = pd.DataFrame({"y": r.index.year, "m": r.index.month, "r": r.values}).pivot(index="y", columns="m", values="r")
    tab = tab.reindex(columns=range(1, 13))
    y = eq.groupby(eq.index.year).last()
    yr = y.pct_change(); yr.iloc[0] = y.iloc[0] / cap - 1
    tab["Year"] = yr.values
    return tab


# ---------------------------------------------------------------- charts
NAVY, GOLD, GREY, RED, GREEN = "#1f3a5f", "#c8962e", "#8a8f98", "#b23a3a", "#2e7d4f"
plt.rcParams.update({"font.size": 9, "axes.spines.top": False, "axes.spines.right": False,
                     "axes.grid": True, "grid.alpha": .25})


def charts(df, eq, bh, dd, bdd, trades, upper, lower, mt, name):
    figs = {}
    f, ax = plt.subplots(figsize=(10, 3.8))
    ax.plot(eq.index, eq, color=NAVY, lw=1.4, label="Strategy"); ax.plot(bh.index, bh, color=GOLD, lw=1.1, label="Buy & hold")
    ax.set_title("Equity curve"); ax.legend(frameon=False); figs["equity"] = f
    f, ax = plt.subplots(figsize=(10, 2.6))
    ax.fill_between(dd.index, dd * 100, 0, color=NAVY, alpha=.35, label="Strategy")
    ax.plot(bdd.index, bdd * 100, color=GOLD, lw=1, label="Buy & hold")
    ax.set_title("Drawdown (%)"); ax.legend(frameon=False); figs["drawdown"] = f
    f, ax = plt.subplots(figsize=(10, 4.2))
    ax.plot(df.date, df.close, color=GREY, lw=1, label="Close")
    ax.plot(df.date, upper, color=GREEN, lw=.7, ls="--", label="Entry channel")
    ax.plot(df.date, lower, color=RED, lw=.7, ls="--", label="Exit channel")
    ax.scatter([t["entry_date"] for t in trades], [t["entry_price"] for t in trades], marker="^", color=GREEN, s=36, zorder=3, label="Entry")
    cl = [t for t in trades if t["status"] == "Closed"]
    ax.scatter([t["exit_date"] for t in cl], [t["exit_price"] for t in cl], marker="v", color=RED, s=36, zorder=3, label="Exit")
    ax.set_title("%s - price, channels and trades" % name); ax.legend(frameon=False, ncol=5, fontsize=8); figs["price"] = f
    f, ax = plt.subplots(figsize=(10, max(1.6, .32 * len(mt) + 1)))
    v = mt.to_numpy(float) * 100; lim = np.nanmax(np.abs(v[:, :12])) if np.isfinite(v[:, :12]).any() else 1
    ax.imshow(np.ma.masked_invalid(v[:, :12]), cmap="RdYlGn", vmin=-lim, vmax=lim, aspect="auto")
    for i in range(v.shape[0]):
        for j in range(12):
            if np.isfinite(v[i, j]): ax.text(j, i, "%.1f" % v[i, j], ha="center", va="center", fontsize=7)
    ax.set_xticks(range(12), ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])
    ax.set_yticks(range(len(mt)), mt.index); ax.grid(False); ax.set_title("Monthly returns (%)"); figs["monthly"] = f
    for f in figs.values(): f.tight_layout()
    return figs


def b64(fig):
    buf = io.BytesIO(); fig.savefig(buf, format="png", dpi=130); return base64.b64encode(buf.getvalue()).decode()


# ---------------------------------------------------------------- formatting
def pct(v): return "-" if v is None or not np.isfinite(v) else "%.2f%%" % (v * 100)
def num(v, d=2): return "-" if v is None or not np.isfinite(v) else ("%.*f" % (d, v) if v != np.inf else "inf")
def money(v): return "-" if not np.isfinite(v) else "{:,.2f}".format(v)
def dt(v): return "-" if v is None or pd.isna(v) else pd.Timestamp(v).strftime(DF)


def build_rows(p, df, s, b, ts):
    rules = [
        ("Instrument / file", "%s (%s)" % (p.name, os.path.basename(p.file))),
        ("Period", "%s to %s (%d bars)" % (dt(df.date.iloc[0]), dt(df.date.iloc[-1]), len(df))),
        ("Entry", "Buy at next bar open when close > highest %s of the previous %d bars" % (p.entry_basis, p.x)),
        ("Exit", "Sell at next bar open when close < lowest %s of the previous %d bars" % (p.exit_basis, p.y)),
        ("Direction / sizing", "Long only; 100% of available equity per trade, whole shares only (leftover cash stays uninvested)"),
        ("Starting capital", money(p.capital)),
        ("Commission", "%s flat per order (charged on entry and on exit)" % money(p.commission)),
        ("Slippage", "%.4g%% of the open price, against you on every fill" % p.slippage),
        ("Benchmark", "Buy & hold: buy at the first bar open (same costs), hold to the end"),
    ]
    perf = [("Final equity", money(s["final_equity"]), money(b["final_equity"])),
            ("Total return", pct(s["total_return"]), pct(b["total_return"])),
            ("CAGR", pct(s["cagr"]), pct(b["cagr"])),
            ("Annualised volatility", pct(s["ann_vol"]), pct(b["ann_vol"])),
            ("Sharpe ratio (rf = 0)", num(s["sharpe"]), num(b["sharpe"])),
            ("Sortino ratio (rf = 0)", num(s["sortino"]), num(b["sortino"])),
            ("Max drawdown", pct(s["max_dd"]), pct(b["max_dd"])),
            ("Calmar ratio", num(s["calmar"]), num(b["calmar"])),
            ("Longest drawdown (days)", str(s["longest_dd_days"]), str(b["longest_dd_days"]))]
    trade = [("Closed trades", str(ts["n_trades"])), ("Open trade at end", "Yes" if ts["open_trade"] else "No"),
             ("Win rate", pct(ts["win_rate"])), ("Profit factor", num(ts["profit_factor"])),
             ("Average trade", pct(ts["avg_trade"])), ("Average win", pct(ts["avg_win"])),
             ("Average loss", pct(ts["avg_loss"])), ("Best trade", pct(ts["best"])), ("Worst trade", pct(ts["worst"])),
             ("Average bars held", num(ts["avg_bars"], 1)), ("Exposure (time in market)", pct(ts["exposure"]))]
    return rules, perf, trade


def trade_rows(trades):
    return [(str(i + 1), dt(t["signal_date"]), dt(t["entry_date"]), num(t["entry_price"], 4), dt(t["exit_signal_date"]),
             dt(t["exit_date"]), num(t["exit_price"], 4), str(int(t["shares"])), money(t["pnl"]), pct(t["ret"]),
             str(t["bars"]), t["status"]) for i, t in enumerate(trades)]


TRADE_HDR = ["#", "Entry signal", "Entry date", "Entry price", "Exit signal", "Exit date", "Exit price",
             "Shares", "P&L", "Return", "Bars", "Status"]


def html(p, rules, perf, trade, trows, mt, figs, warns):
    def tbl(rows, hdr=None, cls=""):
        h = "<tr>%s</tr>" % "".join("<th>%s</th>" % x for x in hdr) if hdr else ""
        return "<table class='%s'>%s%s</table>" % (cls, h, "".join("<tr>%s</tr>" % "".join("<td>%s</td>" % x for x in r) for r in rows))
    def mcell(v):
        if not np.isfinite(v): return "<td></td>"
        a = min(abs(v) / 0.10, 1); col = "rgba(46,125,79,%.2f)" % a if v >= 0 else "rgba(178,58,58,%.2f)" % a
        return "<td style='background:%s'>%.1f</td>" % (col, v * 100)
    mhtml = "<table class='mon'><tr><th></th>%s<th>Year</th></tr>%s</table>" % (
        "".join("<th>%s</th>" % m for m in ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]),
        "".join("<tr><th>%s</th>%s</tr>" % (y, "".join(mcell(v) for v in r)) for y, r in zip(mt.index, mt.to_numpy(float))))
    img = lambda k: "<img src='data:image/png;base64,%s'/>" % b64(figs[k])
    w = "".join("<p class='warn'>Data warning: %s</p>" % x for x in warns)
    return """<!doctype html><html><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'>
<title>%s Channel Breakout</title><style>
:root{--ink:#1d2430;--mute:#5d6672;--line:#dfe3e8;--navy:#1f3a5f;--bg:#fff;--panel:#f6f8fa}
body{font-family:Georgia,'Times New Roman',serif;color:var(--ink);background:var(--bg);margin:0}
.wrap{max-width:1040px;margin:0 auto;padding:28px 16px}
header{border-bottom:3px solid var(--navy);padding-bottom:10px;margin-bottom:18px}
h1{margin:0;font-size:26px;color:var(--navy)} .sub{color:var(--mute);font-family:Arial,sans-serif;font-size:13px}
h2{font-size:16px;color:var(--navy);border-bottom:1px solid var(--line);padding-bottom:4px;margin-top:28px}
table{border-collapse:collapse;width:100%%;font-family:Arial,sans-serif;font-size:12.5px}
td,th{border-bottom:1px solid var(--line);padding:5px 8px;text-align:left} th{background:var(--panel);font-weight:600}
.grid{display:grid;grid-template-columns:1.4fr 1fr;gap:22px} .num td:not(:first-child),.num th:not(:first-child){text-align:right}
.mon td,.mon th{text-align:center;font-size:11.5px;padding:4px} img{width:100%%;height:auto}
.scroll{overflow-x:auto} .trades td,.trades th{white-space:nowrap;font-size:11.5px}
.warn{background:#fff4e0;border-left:3px solid #c8962e;padding:6px 10px;font-family:Arial,sans-serif;font-size:12.5px}
.note{color:var(--mute);font-family:Arial,sans-serif;font-size:11.5px}
@media(max-width:760px){.grid{grid-template-columns:1fr}}
</style></head><body><div class='wrap'>
<header><h1>%s &mdash; Simple Channel Breakout</h1><div class='sub'>Backtest tearsheet &middot; generated %s</div></header>%s
<h2>Strategy rules</h2>%s
<div class='grid'><div><h2>Performance vs buy &amp; hold</h2>%s</div><div><h2>Trade statistics</h2>%s</div></div>
<h2>Equity curve</h2>%s<h2>Drawdown</h2>%s<h2>Price, channels and trades</h2>%s
<h2>Monthly returns (%%)</h2><div class='scroll'>%s</div>
<h2>Trade list</h2><div class='scroll'>%s</div>
<p class='note'>Signals are evaluated on each bar's close and filled at the next bar's open. Channels use the previous X / Y bars only (the signal bar is excluded). Whole shares only; any leftover cash stays uninvested. Ratios use a zero risk-free rate and are annualised from the data frequency. Past performance in a backtest does not guarantee future results.</p>
</div></body></html>""" % (p.name, p.name, datetime.now().strftime(DF), w, tbl(rules, cls=""),
                            tbl(perf, ["Metric", "Strategy", "Buy & hold"], "num"), tbl(trade, ["Metric", "Value"], "num"),
                            img("equity"), img("drawdown"), img("price"), mhtml,
                            tbl(trows, TRADE_HDR, "trades") if trows else "<p class='note'>No trades were triggered.</p>")


def pdf(path, p, rules, perf, trade, trows, figs):
    def table_page(pp, title, blocks):
        f = plt.figure(figsize=(11.69, 8.27)); f.suptitle(title, x=.05, ha="left", fontsize=15, color=NAVY, weight="bold")
        y = .9
        for sub, rows, hdr, widths in blocks:
            h = .028 * (len(rows) + (1 if hdr else 0))
            f.text(.05, y, sub, fontsize=11, color=NAVY, weight="bold")
            ax = f.add_axes([.05, y - h - .01, .9, h]); ax.axis("off")
            t = ax.table(cellText=rows, colLabels=hdr, colWidths=widths, loc="upper left", cellLoc="left")
            t.auto_set_font_size(False); t.set_fontsize(8)
            for (r, _), cell in t.get_celld().items():
                cell.set_edgecolor("#dfe3e8")
                if hdr and r == 0: cell.set_facecolor("#f6f8fa")
            y -= h + .07
        pp.savefig(f); plt.close(f)
    with PdfPages(path) as pp:
        table_page(pp, "%s - Simple Channel Breakout" % p.name,
                   [("Strategy rules", rules, None, [.22, .78]),
                    ("Performance vs buy & hold", perf, ["Metric", "Strategy", "Buy & hold"], [.4, .3, .3])])
        table_page(pp, "Trade statistics", [("", trade, ["Metric", "Value"], [.5, .5])])
        for k in ("equity", "drawdown", "price", "monthly"):
            figs[k].set_size_inches(11.69, 8.27 if k in ("price", "monthly") else 5); pp.savefig(figs[k])
        for i in range(0, max(len(trows), 1), 25):
            chunk = trows[i:i + 25] or [["-"] * len(TRADE_HDR)]
            table_page(pp, "Trade list" + (" (cont.)" if i else ""), [("", chunk, TRADE_HDR, None)])


# ---------------------------------------------------------------- main
def main():
    a = argparse.ArgumentParser()
    a.add_argument("--file", required=True); a.add_argument("--sheet")
    a.add_argument("--name", help="label for the report, defaults to file name")
    a.add_argument("--entry-basis", choices=["high", "close"], required=True)
    a.add_argument("--x", type=int, required=True)
    a.add_argument("--exit-basis", choices=["low", "close"], required=True)
    a.add_argument("--y", type=int, required=True)
    a.add_argument("--capital", type=float, required=True)
    a.add_argument("--commission", type=float, default=0.0)
    a.add_argument("--slippage", type=float, default=0.0, help="percent, e.g. 0.05 = 0.05%%")
    a.add_argument("--pdf", action="store_true"); a.add_argument("--outdir", default=".")
    p = a.parse_args()
    if p.x < 1 or p.y < 1: fail("X and Y must be whole numbers of at least 1.")
    if p.capital <= 0: fail("Starting capital must be positive.")
    if p.commission < 0 or p.slippage < 0: fail("Commission and slippage cannot be negative.")
    if p.commission >= p.capital: fail("Commission is larger than starting capital.")
    p.name = p.name or os.path.splitext(os.path.basename(p.file))[0]

    df, warns = load(p.file, p.sheet)
    if len(df) <= max(p.x, p.y) + 1: fail("Only %d bars - not enough for X=%d / Y=%d." % (len(df), p.x, p.y))
    eq, bh, pos, trades, upper, lower, skipped = run(df, p)
    if skipped:
        warns.append("%d entry signal(s) skipped: not enough cash to buy one whole share." % skipped)
    ppy = periods_per_year(df.date)
    s, dd = stats(eq, p.capital, ppy); b, bdd = stats(bh, p.capital, ppy); ts = trade_stats(trades, pos)
    mt = monthly_table(eq, p.capital)
    figs = charts(df, eq, bh, dd, bdd, trades, upper, lower, mt, p.name)
    rules, perf, trade = build_rows(p, df, s, b, ts); trows = trade_rows(trades)

    os.makedirs(p.outdir, exist_ok=True)
    stem = "%s_breakout_E%s%d_X%s%d" % (p.name.replace(" ", "_"), p.entry_basis[0].upper(), p.x, p.exit_basis[0].upper(), p.y)
    out = {"status": "ok", "html": os.path.join(p.outdir, stem + ".html")}
    with open(out["html"], "w", encoding="utf-8") as fh: fh.write(html(p, rules, perf, trade, trows, mt, figs, warns))
    out["trades_csv"] = os.path.join(p.outdir, stem + "_trades.csv")
    pd.DataFrame(trows, columns=TRADE_HDR).to_csv(out["trades_csv"], index=False)
    if p.pdf:
        out["pdf"] = os.path.join(p.outdir, stem + ".pdf"); pdf(out["pdf"], p, rules, perf, trade, trows, figs)
    out.update(period="%s to %s" % (dt(df.date.iloc[0]), dt(df.date.iloc[-1])), bars=len(df), warnings=warns,
               trades={k: v for k, v in trade})
    out["strategy"] = {r[0]: r[1] for r in perf}; out["buy_and_hold"] = {r[0]: r[2] for r in perf}
    print(json.dumps(out, indent=1))


if __name__ == "__main__":
    main()
```
