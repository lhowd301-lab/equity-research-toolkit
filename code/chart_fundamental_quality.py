# -*- coding: utf-8 -*-
"""Section 6: Fundamental Quality Analysis Chart for 创元科技 (000551)"""

from pathlib import Path as _PublicPath
_PUBLIC_ROOT = _PublicPath(__file__).resolve().parents[1]
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np
import pandas as pd

BASE = str(_PUBLIC_ROOT)
DATA = os.path.join(BASE, "data", "创元科技_000551_体检数据_20260806.xlsx")
IMG = os.path.join(BASE, "output", "图片", "06_基本面质量分析.png")

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "STSong"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["figure.dpi"] = 120

# ── Load & filter annual data ──────────────────────────────
xl = pd.ExcelFile(DATA)
pl = xl.parse(9)   # 利润表
bs = xl.parse(8)   # 资产负债表
cf = xl.parse(10)  # 现金流量表

# Find the annual report type (年报 = b'\xc4\xea\xb1\xa8' in GBK)
annual_type = None
for rt in pl["REPORT_TYPE"].unique():
    if "年" in str(rt) or rt.encode("gbk", errors="replace") == b'\xc4\xea\xb1\xa8':
        annual_type = rt
        break
print(f"Annual report type: [{annual_type}]")

def filter_annual(df):
    df = df.copy()
    df["REPORT_DATE"] = pd.to_datetime(df["REPORT_DATE"])
    mask = df["REPORT_TYPE"] == annual_type
    df = df[mask].copy()
    # Some years may have duplicate annual reports, keep last
    df["year"] = df["REPORT_DATE"].dt.year
    df = df.sort_values("REPORT_DATE").drop_duplicates("year", keep="last")
    return df.set_index("year").sort_index()

pl_a = filter_annual(pl)
bs_a = filter_annual(bs)
cf_a = filter_annual(cf)

# ── Extract key series ─────────────────────────────────────
# Profit statement
revenue     = pl_a.get("TOTAL_OPERATE_INCOME", pl_a.get("OPERATE_INCOME")) / 1e8  # 亿
operate_profit = pl_a["OPERATE_PROFIT"] / 1e8
net_profit  = pl_a["PARENT_NETPROFIT"] / 1e8  # 归母净利润
total_cost  = pl_a.get("TOTAL_OPERATE_COST", pl_a.get("OPERATE_COST")) / 1e8
basic_eps   = pl_a["BASIC_EPS"]

# Balance sheet
total_assets    = bs_a["TOTAL_ASSETS"] / 1e8
total_liab      = bs_a["TOTAL_LIABILITIES"] / 1e8
total_equity    = bs_a["TOTAL_EQUITY"] / 1e8
parent_equity   = bs_a["TOTAL_PARENT_EQUITY"] / 1e8

# Cash flow
ocf = cf_a["NETCASH_OPERATE"] / 1e8

# ── Derived metrics ────────────────────────────────────────
gross_margin   = (revenue - total_cost) / revenue * 100
operate_margin = operate_profit / revenue * 100
net_margin     = net_profit / revenue * 100
revenue_growth = revenue.pct_change() * 100
net_profit_growth = net_profit.pct_change() * 100

# ROE Dupont: ROE = Net Margin * Asset Turnover * Equity Multiplier
asset_turnover  = revenue / total_assets
equity_multiplier = total_assets / parent_equity
roe_dupont = net_margin / 100 * asset_turnover * equity_multiplier * 100  # ROE %

# Simple ROE: Net Profit / Parent Equity * 100
roe_simple = net_profit / parent_equity * 100

# OCF quality
ocf_to_ni = ocf / net_profit
ocf_to_ni_clean = ocf_to_ni.replace([np.inf, -np.inf], np.nan)

# Debt ratio
debt_ratio = total_liab / total_assets * 100

# ── Filter to 1999–2025 (skip early years with sparse data) ──
start_yr, end_yr = 1999, 2025
idx_slice = slice(start_yr, end_yr)
years = np.arange(start_yr, end_yr + 1)

rev     = revenue.loc[idx_slice]
op_prof = operate_profit.loc[idx_slice]
ni      = net_profit.loc[idx_slice]
eps     = basic_eps.loc[idx_slice]
gm      = gross_margin.loc[idx_slice]
om      = operate_margin.loc[idx_slice]
nm      = net_margin.loc[idx_slice]
roe     = roe_dupont.loc[idx_slice]
roe_s   = roe_simple.loc[idx_slice]
rg      = revenue_growth.loc[idx_slice]
ng      = net_profit_growth.loc[idx_slice]
ocf_ni  = ocf_to_ni_clean.loc[idx_slice]
at      = asset_turnover.loc[idx_slice]
em      = equity_multiplier.loc[idx_slice]
dr      = debt_ratio.loc[idx_slice]

y = years

print(f"Data range: {start_yr}-{end_yr}, Revenue 2025: {rev.get(2025, np.nan):.1f}亿")
print(f"Net Profit 2025: {ni.get(2025, np.nan):.1f}亿")

# ── Colors ─────────────────────────────────────────────────
C1 = "#2c7bb6"  # blue
C2 = "#d7191c"  # red
C3 = "#fdae61"  # orange
C4 = "#5e3c99"  # purple
C5 = "#1a9641"  # green
C6 = "#a6611a"  # brown
GRAY = "#999999"

# ── Plot ───────────────────────────────────────────────────
fig, axes = plt.subplots(3, 2, figsize=(18, 13.5))
fig.subplots_adjust(hspace=0.35, wspace=0.28)

# ---- 1. Revenue & Net Profit (top-left) ---------------------
ax = axes[0, 0]
bar_w = 0.35
x = np.arange(len(y))
b1 = ax.bar(x - bar_w/2, rev.values, bar_w, color=C1, alpha=0.85, label="Revenue (100M CNY)")
ax2_r = ax.twinx()
b2 = ax2_r.bar(x + bar_w/2, ni.values, bar_w, color=C2, alpha=0.85, label="Net Profit (100M CNY)")
ax.set_xticks(x)
ax.set_xticklabels(y, rotation=45, fontsize=7)
ax.set_ylabel("Revenue (100M CNY)", fontsize=9, color=C1)
ax2_r.set_ylabel("Net Profit (100M CNY)", fontsize=9, color=C2)
ax.set_title("Revenue & Net Profit (Attributable to Parent)", fontsize=11, fontweight="bold")
# Combine legends
lines1, labels1 = ax.get_legend_handles_labels()
lines2, labels2 = ax2_r.get_legend_handles_labels()
ax.legend(lines1 + lines2, labels1 + labels2, loc="upper left", fontsize=7, framealpha=0.8)

# ---- 2. Profit Margins (top-right) --------------------------
ax = axes[0, 1]
ax.plot(y, gm.values, "o-", color=C5, linewidth=1.8, markersize=4, label="Gross Margin")
ax.plot(y, om.values, "s-", color=C3, linewidth=1.8, markersize=4, label="Operating Margin")
ax.plot(y, nm.values, "D-", color=C2, linewidth=1.8, markersize=4, label="Net Margin")
ax.axhline(0, color="black", linewidth=0.5)
ax.set_ylabel("Margin (%)", fontsize=9)
ax.set_title("Profit Margins (Gross / Operating / Net)", fontsize=11, fontweight="bold")
ax.legend(loc="best", fontsize=7, framealpha=0.8)
ax.tick_params(axis="x", rotation=45, labelsize=7)
ax.set_xticks(y[::2])

# ---- 3. ROE Dupont Decomposition (middle-left) --------------
ax = axes[1, 0]
ax.plot(y, roe.values, "D-", color=C4, linewidth=2.2, markersize=5, label="ROE (Dupont)")
ax.axhline(0, color="black", linewidth=0.5)
ax.set_ylabel("ROE (%)", fontsize=9)
ax.set_title("ROE (Dupont Decomposition)", fontsize=11, fontweight="bold")
ax.tick_params(axis="x", rotation=45, labelsize=7)
ax.set_xticks(y[::2])

# Add dupont components as small text annotations for recent years
recent_y = [2021, 2022, 2023, 2024, 2025]
for yr in recent_y:
    if yr in roe.index and yr in nm.index and yr in at.index and yr in em.index:
        rv = roe.get(yr, np.nan)
        if not np.isnan(rv):
            ax.annotate(f"{rv:.1f}%", (yr, rv), textcoords="offset points",
                        xytext=(0, 10), fontsize=7, ha="center", color=C4)

# ---- 4. Revenue Growth YoY (middle-right) -------------------
ax = axes[1, 1]
colors_g = [C5 if v >= 0 else C2 for v in rg.values]
ax.bar(y, rg.values, color=colors_g, alpha=0.8, width=0.6)
ax.axhline(0, color="black", linewidth=0.5)
ax.set_ylabel("Revenue YoY Growth (%)", fontsize=9)
ax.set_title("Revenue Growth Rate (YoY)", fontsize=11, fontweight="bold")
ax.tick_params(axis="x", rotation=45, labelsize=7)
ax.set_xticks(y[::2])
# Add mean line
mean_g = rg.mean()
ax.axhline(mean_g, color=GRAY, linestyle="--", linewidth=1, alpha=0.7)
ax.text(0.98, 0.92, f"Mean: {mean_g:.1f}%", transform=ax.transAxes,
        fontsize=8, ha="right", va="top", color=GRAY)

# ---- 5. OCF / Net Income (bottom-left) ----------------------
ax = axes[2, 0]
colors_o = [C5 if v >= 0 else C2 for v in ocf_ni.values]
ax.bar(y, ocf_ni.values, color=colors_o, alpha=0.8, width=0.6)
ax.axhline(1.0, color=C1, linestyle="--", linewidth=1, alpha=0.6, label="Ideal = 1.0")
ax.axhline(0, color="black", linewidth=0.5)
ax.set_ylabel("OCF / Net Income", fontsize=9)
ax.set_title("Cash Flow Quality: Operating CF / Net Income", fontsize=11, fontweight="bold")
ax.legend(loc="best", fontsize=7)
ax.tick_params(axis="x", rotation=45, labelsize=7)
ax.set_xticks(y[::2])

# ---- 6. EPS (bottom-right) ----------------------------------
ax = axes[2, 1]
colors_e = [C5 if v >= 0 else C2 for v in eps.values]
ax.bar(y, eps.values, color=colors_e, alpha=0.8, width=0.6)
ax.axhline(0, color="black", linewidth=0.5)
ax.set_ylabel("EPS (CNY)", fontsize=9)
ax.set_title("Basic Earnings Per Share", fontsize=11, fontweight="bold")
ax.tick_params(axis="x", rotation=45, labelsize=7)
ax.set_xticks(y[::2])
# Annotate recent values
for yr in recent_y:
    if yr in eps.index:
        v = eps.get(yr, np.nan)
        if not np.isnan(v):
            ax.annotate(f"{v:.2f}", (yr, v), textcoords="offset points",
                        xytext=(0, 8), fontsize=7, ha="center", color=C4)

# ── Overall title ───────────────────────────────────────────
fig.suptitle("创元科技 (000551.SZ) — Fundamental Quality Analysis (1999–2025)",
             fontsize=14, fontweight="bold", y=1.01)

# ── Add summary text at the bottom ──────────────────────────
# Compute recent-5-year averages
mask_5y = years >= 2021
summary_lines = [
    f"5Y Avg Revenue Growth: {rg.loc[mask_5y].mean():.1f}% | "
    f"5Y Avg Net Margin: {nm.loc[mask_5y].mean():.1f}% | "
    f"5Y Avg ROE: {roe.loc[mask_5y].mean():.1f}%",
    f"5Y Avg OCF/NI: {ocf_ni.loc[mask_5y].mean():.2f} | "
    f"2025 Revenue: {rev.get(2025, np.nan):.1f}亿 | "
    f"2025 Net Profit: {ni.get(2025, np.nan):.1f}亿 | "
    f"2025 EPS: {eps.get(2025, np.nan):.2f}",
]
fig.text(0.5, -0.01, "\n".join(summary_lines), ha="center", fontsize=8.5,
         color=GRAY, transform=axes[2, 0].transAxes)

# ── Save ────────────────────────────────────────────────────
plt.savefig(IMG, dpi=150, bbox_inches="tight", facecolor="white")
plt.close()
print(f"Saved: {IMG}")
print("Done.")
