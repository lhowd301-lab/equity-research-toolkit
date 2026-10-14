# -*- coding: utf-8 -*-
"""
创元科技(000551.SZ) 历史收益表现分析

输入: data/创元科技_000551_体检数据_YYYYMMDD.xlsx
输出: output/创元科技_收益表现分析_YYYYMMDD.xlsx
      output/图片/*.png

用法: python analysis_return_performance.py
"""

from pathlib import Path as _PublicPath
_PUBLIC_ROOT = _PublicPath(__file__).resolve().parents[1]

import os
from datetime import datetime

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.ticker import FuncFormatter

BASE_DIR = str(_PUBLIC_ROOT)
DATA_DIR = os.path.join(BASE_DIR, "data")
OUT_DIR = os.path.join(BASE_DIR, "output")
IMG_DIR = os.path.join(OUT_DIR, "图片")

TODAY = datetime.now().strftime("%Y%m%d")
DATA_FILE = None
for f in os.listdir(DATA_DIR):
    if f.startswith("创元科技") and f.endswith(".xlsx") and "体检数据" in f:
        DATA_FILE = os.path.join(DATA_DIR, f)
        break
if DATA_FILE is None:
    raise FileNotFoundError("未找到体检数据 Excel 文件")

OUT_FILE = os.path.join(OUT_DIR, f"创元科技_收益表现分析_{TODAY}.xlsx")

ANALYSIS_START = "1999-01-01"
SEED_START = "1998-12-31"

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "STSong"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["figure.dpi"] = 110


def load_close_table():
    """读取个股(后复权)与基准指数收盘价, 返回长表与收盘价表。"""
    xl = pd.ExcelFile(DATA_FILE)

    hist = xl.parse("日线行情_后复权")
    hist["日期"] = pd.to_datetime(hist["日期"])
    hist = hist[["日期", "收盘"]].rename(columns={"收盘": "创元科技"}).set_index("日期").sort_index()

    idx = xl.parse("基准指数日线")
    idx["日期"] = pd.to_datetime(idx["日期"])
    idx_close = idx.pivot_table(index="日期", columns="指数名称", values="收盘")

    sw = xl.parse("申万行业指数日线")
    sw["日期"] = pd.to_datetime(sw["日期"])
    sw_close = sw.set_index("日期")[["收盘"]].rename(columns={"收盘": "申万环保"})

    close = hist.join(idx_close, how="outer").join(sw_close, how="outer").sort_index()
    close = close.dropna(how="all")
    return close


def pct(x):
    return f"{x * 100:.2f}%"


def build_tables(close):
    """生成各类收益统计表。"""
    stock = close["创元科技"].dropna()
    stock = stock.loc[ANALYSIS_START:]
    ret = stock.pct_change().dropna()
    years = (stock.index[-1] - stock.index[0]).days / 365.25
    total_ret = stock.iloc[-1] / stock.iloc[0] - 1
    cagr = (1 + total_ret) ** (1 / years) - 1

    # 年度收益
    yearly_raw = close.resample("Y").last()
    yearly = yearly_raw.pct_change().dropna(how="all")
    yearly = yearly[yearly.index.year >= 1999]
    yearly.index = yearly.index.year
    yearly.index.name = "年份"
    yearly = yearly.round(4)

    yearly_out = yearly.reset_index()
    if "沪深300" in yearly_out.columns:
        yearly_out["超额_沪深300"] = yearly_out["创元科技"] - yearly_out["沪深300"]
    if "中证500" in yearly_out.columns:
        yearly_out["超额_中证500"] = yearly_out["创元科技"] - yearly_out["中证500"]

    # 月度收益
    monthly = close.resample("M").last().pct_change().dropna(how="all")
    monthly = monthly[monthly.index >= ANALYSIS_START]
    stock_m = monthly["创元科技"].dropna()

    # 滚动收益(交易日口径, 1/3/5/10年)
    rolling_rows = []
    for years_back, label in [(1, "1年"), (3, "3年"), (5, "5年"), (10, "10年")]:
        h = round(years_back * 252)
        if len(stock) <= h:
            continue
        roll = stock / stock.shift(h) - 1
        roll = roll.dropna()
        ann = (1 + roll) ** (252 / h) - 1
        cur = ann.iloc[-1]
        rows = len(roll)
        rolling_rows.append(
            {
                "持有期限": label,
                "交易日窗口": h,
                "样本数": rows,
                "胜率": round((roll > 0).mean(), 4),
                "平均年化": round(ann.mean(), 4),
                "中位年化": round(ann.median(), 4),
                "最差年化": round(ann.min(), 4),
                "最好年化": round(ann.max(), 4),
                "当前年化": round(cur, 4),
                "当前分位": round((ann <= cur).mean(), 4),
            }
        )
    rolling_df = pd.DataFrame(rolling_rows)

    # 年度内月度统计
    m_stats = []
    for yr, grp in stock_m.groupby(stock_m.index.year):
        pos = (grp > 0).sum()
        neg = (grp <= 0).sum()
        m_stats.append(
            {
                "年份": yr,
                "月份数": len(grp),
                "正收益月数": pos,
                "负收益月数": neg,
                "月度胜率": round(pos / len(grp), 4),
                "最好月份": round(grp.max(), 4),
                "最差月份": round(grp.min(), 4),
                "年度累计": round((1 + grp).prod() - 1, 4),
            }
        )
    m_stats_df = pd.DataFrame(m_stats)

    # 持有体验模拟(月度口径)
    hold_rows = []
    for years_back in [1, 3, 5, 10]:
        horizon = years_back * 12
        vals = []
        mvals = stock_m.values
        for i in range(len(mvals) - horizon):
            r = np.prod(1 + mvals[i + 1 : i + 1 + horizon]) - 1
            vals.append(r)
        if not vals:
            continue
        arr = np.array(vals)
        ann = (1 + arr) ** (12 / horizon) - 1
        hold_rows.append(
            {
                "持有年限": years_back,
                "样本数": len(arr),
                "胜率": round((arr > 0).mean(), 4),
                "平均年化": round(ann.mean(), 4),
                "中位年化": round(float(np.median(ann)), 4),
                "最差年化": round(ann.min(), 4),
                "最好年化": round(ann.max(), 4),
            }
        )
    hold_df = pd.DataFrame(hold_rows)

    # 核心指标
    wealth = (1 + ret).cumprod()
    dd = wealth / wealth.cummax() - 1
    max_dd = dd.min()
    ann_vol = ret.std() * np.sqrt(252)

    core_rows = [
        ("数据区间", f"{stock.index[0].date()} 至 {stock.index[-1].date()}"),
        ("交易日数", len(stock)),
        ("累计收益率(后复权)", round(total_ret, 4)),
        ("年化收益率CAGR", round(cagr, 4)),
        ("年化波动率", round(ann_vol, 4)),
        ("最大回撤", round(max_dd, 4)),
        ("月度胜率", round((stock_m > 0).mean(), 4)),
        ("年度胜率(完整年度)", round((yearly["创元科技"].iloc[:-1] > 0).mean(), 4)),
        ("最好年度", f"{yearly['创元科技'].idxmax()} ({pct(yearly['创元科技'].max())})"),
        ("最差年度", f"{yearly['创元科技'].idxmin()} ({pct(yearly['创元科技'].min())})"),
        ("最好月份", pct(stock_m.max())),
        ("最差月份", pct(stock_m.min())),
    ]
    for years_back, label in [(1, "近1年"), (3, "近3年"), (5, "近5年"), (10, "近10年")]:
        target = stock.index[-1] - pd.DateOffset(years=years_back)
        anchor = stock[stock.index <= target]
        if len(anchor) > 0:
            core_rows.append((f"{label}收益", round(stock.iloc[-1] / anchor.iloc[-1] - 1, 4)))

    # 累计超额
    for bench in ["沪深300", "中证500", "申万环保"]:
        if bench not in close.columns:
            continue
        b = close[bench].dropna()
        common = pd.concat([stock, b], axis=1, join="inner").dropna()
        if len(common) < 60:
            continue
        s_cum = common.iloc[:, 0] / common.iloc[0, 0] - 1
        b_cum = common.iloc[:, 1] / common.iloc[0, 1] - 1
        core_rows.append((f"对{bench}累计超额", round(s_cum.iloc[-1] - b_cum.iloc[-1], 4)))

    core_df = pd.DataFrame(core_rows, columns=["指标", "数值"])
    return core_df, yearly_out, rolling_df, m_stats_df, hold_df, monthly, stock_m


def save_chart_1(close):
    """个股累计净值 + 回撤。"""
    stock = close["创元科技"].dropna()
    stock = stock.loc[ANALYSIS_START:]
    wealth = stock / stock.iloc[0]
    dd = wealth / wealth.cummax() - 1
    fig, axes = plt.subplots(2, 1, figsize=(13, 8), sharex=True,
                             gridspec_kw={"height_ratios": [3, 1]})
    axes[0].plot(wealth.index, wealth.values, color="#c0392b", lw=1.2)
    axes[0].set_yscale("log")
    axes[0].set_title(f"创元科技 累计净值（后复权, {ANALYSIS_START[:4]}-01 起）", fontsize=14)
    axes[0].set_ylabel("累计净值（对数轴）")
    axes[0].grid(alpha=0.3)
    axes[0].axhline(1, color="gray", lw=0.8, ls="--")

    axes[1].fill_between(dd.index, dd.values * 100, 0, color="#2980b9", alpha=0.55)
    axes[1].set_title("历史回撤（%）", fontsize=12)
    axes[1].set_ylabel("回撤 %")
    axes[1].grid(alpha=0.3)
    axes[1].set_xlim(wealth.index[0], wealth.index[-1])
    fig.tight_layout()
    fig.savefig(os.path.join(IMG_DIR, "01_累计净值与回撤.png"), dpi=150)
    plt.close(fig)


def save_chart_2(close):
    """基准对比累计净值。"""
    cols = ["创元科技", "沪深300", "中证500", "中证1000"]
    cols = [c for c in cols if c in close.columns]
    fig, ax = plt.subplots(figsize=(13, 7))
    colors = ["#c0392b", "#2471a3", "#e67e22", "#7d3c98"]
    for col, color in zip(cols, colors):
        s = close[col].dropna()
        s = s[s.index >= "2005-01-01"]
        if len(s) == 0:
            continue
        norm = s / s.iloc[0]
        ax.plot(norm.index, norm.values, label=col, color=color, lw=1.3)
    ax.set_yscale("log")
    ax.set_title("累计净值对比（2005 年起, 对数轴）", fontsize=14)
    ax.set_ylabel("净值（期初=1）")
    ax.legend()
    ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(IMG_DIR, "02_基准对比累计净值.png"), dpi=150)
    plt.close(fig)


def save_chart_3(yearly_out):
    """年度收益对比柱状图。"""
    df = yearly_out.copy()
    df = df[df["年份"] >= 2005]
    cols = ["创元科技", "沪深300", "中证500"]
    cols = [c for c in cols if c in df.columns]
    x = np.arange(len(df))
    width = 0.26
    fig, ax = plt.subplots(figsize=(14, 6.5))
    colors = ["#c0392b", "#2471a3", "#e67e22"]
    for i, col in enumerate(cols):
        ax.bar(x + (i - 1) * width, df[col].values * 100, width * 0.9,
               label=col, color=colors[i], alpha=0.92)
    ax.axhline(0, color="black", lw=0.8)
    ax.set_xticks(x)
    ax.set_xticklabels(df["年份"].astype(str), rotation=45)
    ax.set_title("年度收益率对比（%）", fontsize=14)
    ax.set_ylabel("收益率 %")
    ax.legend()
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(IMG_DIR, "03_年度收益对比.png"), dpi=150)
    plt.close(fig)


def save_chart_4(stock_m):
    """月度收益热力图。"""
    heat = stock_m * 100
    table = heat.groupby([heat.index.year, heat.index.month]).last().unstack()
    table = table.round(1)
    fig, ax = plt.subplots(figsize=(13, 10))
    im = ax.imshow(table.values, cmap="RdYlGn", aspect="auto", vmin=-30, vmax=30)
    ax.set_xticks(range(len(table.columns)))
    ax.set_xticklabels([f"{m}月" for m in table.columns])
    ax.set_yticks(range(len(table.index)))
    ax.set_yticklabels(table.index)
    ax.set_title("创元科技 月度收益率（%）", fontsize=14)
    for i in range(table.shape[0]):
        for j in range(table.shape[1]):
            v = table.values[i, j]
            if np.isnan(v):
                continue
            ax.text(j, i, f"{v:.0f}", ha="center", va="center",
                    fontsize=6.5, color="black")
    fig.colorbar(im, ax=ax, label="月度收益 %")
    fig.tight_layout()
    fig.savefig(os.path.join(IMG_DIR, "04_月度收益热力图.png"), dpi=150)
    plt.close(fig)


def save_chart_5(close):
    """滚动 1/3/5/10 年年化收益。"""
    stock = close["创元科技"].dropna()
    stock = stock.loc[ANALYSIS_START:]
    fig, axes = plt.subplots(2, 2, figsize=(14, 9), sharex=True)
    specs = [(1, "1年", axes[0, 0]), (3, "3年", axes[0, 1]),
             (5, "5年", axes[1, 0]), (10, "10年", axes[1, 1])]
    for years_back, label, ax in specs:
        h = round(years_back * 252)
        if len(stock) <= h:
            ax.text(0.5, 0.5, "样本不足", ha="center", transform=ax.transAxes)
            continue
        ann = (stock / stock.shift(h) - 1).dropna()
        ann = (1 + ann) ** (252 / h) - 1
        ax.plot(ann.index, ann.values * 100, color="#c0392b", lw=1.0)
        ax.axhline(0, color="gray", lw=0.8, ls="--")
        ax.set_title(f"滚动{label}年化收益（%）", fontsize=12)
        ax.grid(alpha=0.3)
        cur = ann.iloc[-1] * 100
        ax.annotate(f"当前 {cur:.1f}%", xy=(ann.index[-1], cur),
                    xytext=(-110, 12), textcoords="offset points", fontsize=10)
    fig.tight_layout()
    fig.savefig(os.path.join(IMG_DIR, "05_滚动收益.png"), dpi=150)
    plt.close(fig)


def save_chart_6(close):
    """相对强弱与累计超额。"""
    stock = close["创元科技"].dropna()
    fig, axes = plt.subplots(2, 1, figsize=(13, 9))

    # 相对强弱 vs 沪深300
    b = close["沪深300"].dropna()
    common = pd.concat([stock, b], axis=1, join="inner").dropna()
    common = common[common.index >= "2005-01-01"]
    rs = common.iloc[:, 0] / common.iloc[:, 1] * 100
    axes[0].plot(rs.index, rs.values, color="#2471a3", lw=1.1)
    axes[0].axhline(rs.iloc[0], color="gray", lw=0.8, ls="--")
    axes[0].set_title("相对强弱：创元科技 / 沪深300（基期=100）", fontsize=13)
    axes[0].grid(alpha=0.3)
    axes[0].set_ylabel("相对强弱")

    # 累计超额
    for bench, color, start in [("沪深300", "#c0392b", "2005-01-01"),
                                ("申万环保", "#16a085", "2022-01-01")]:
        if bench not in close.columns:
            continue
        b2 = close[bench].dropna()
        common2 = pd.concat([stock, b2], axis=1, join="inner").dropna()
        common2 = common2[common2.index >= start]
        if len(common2) < 60:
            continue
        s_cum = common2.iloc[:, 0] / common2.iloc[0, 0] - 1
        b_cum = common2.iloc[:, 1] / common2.iloc[0, 1] - 1
        excess = (s_cum - b_cum) * 100
        axes[1].plot(excess.index, excess.values, label=f"超额 vs {bench}", color=color, lw=1.1)
    axes[1].axhline(0, color="gray", lw=0.8, ls="--")
    axes[1].set_title("累计超额收益（%）", fontsize=13)
    axes[1].set_ylabel("累计超额 %")
    axes[1].legend()
    axes[1].grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(IMG_DIR, "06_相对强弱与超额收益.png"), dpi=150)
    plt.close(fig)


def save_chart_7(stock_m):
    """持有体验箱线图。"""
    data, labels = [], []
    for years_back in [1, 3, 5, 10]:
        horizon = years_back * 12
        mvals = stock_m.values
        vals = []
        for i in range(len(mvals) - horizon):
            r = np.prod(1 + mvals[i + 1 : i + 1 + horizon]) - 1
            vals.append((1 + r) ** (12 / horizon) - 1)
        if vals:
            data.append(np.array(vals) * 100)
            labels.append(f"{years_back}年")
    fig, ax = plt.subplots(figsize=(10, 6.5))
    bp = ax.boxplot(data, labels=labels, patch_artist=True,
                    medianprops={"color": "black"})
    colors = ["#c0392b", "#e67e22", "#2471a3", "#16a085"]
    for patch, color in zip(bp["boxes"], colors):
        patch.set_facecolor(color)
        patch.set_alpha(0.55)
    ax.axhline(0, color="gray", lw=0.8, ls="--")
    ax.set_title("任意时点买入持有 N 年的年化收益分布（%）", fontsize=14)
    ax.set_ylabel("年化收益 %")
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(IMG_DIR, "07_持有体验箱线图.png"), dpi=150)
    plt.close(fig)


def main():
    os.makedirs(IMG_DIR, exist_ok=True)
    close = load_close_table()
    close = close[close.index >= SEED_START]
    core_df, yearly_out, rolling_df, m_stats_df, hold_df, monthly, stock_m = build_tables(close)
    last_date = close["创元科技"].dropna().index[-1].date()

    save_chart_1(close)
    save_chart_2(close)
    save_chart_3(yearly_out)
    save_chart_4(stock_m)
    save_chart_5(close)
    save_chart_6(close)
    save_chart_7(stock_m)

    notes = pd.DataFrame(
        [
            ["股票", "000551 创元科技"],
            ["收益口径", "后复权收盘价计算, 已含分红再投资"],
            ["数据来源", "data/创元科技_000551_体检数据_20260806.xlsx"],
            ["样本起点", f"{ANALYSIS_START[:4]}-01 起（1996 年涨幅 716% 属极端异常值, 已剔除）"],
            ["说明1", f"{last_date.year} 年为截至 {last_date} 的当年收益, 非完整年度"],
            ["说明2", "滚动收益按交易日口径: 1/3/5/10 年分别约 252/756/1260/2520 个交易日"],
            ["说明3", "年度胜率仅统计完整年度"],
            ["说明4", "持有体验模拟按月度收益滚动计算"],
            ["图1", "累计净值与回撤"],
            ["图2", "基准对比累计净值"],
            ["图3", "年度收益对比"],
            ["图4", "月度收益热力图"],
            ["图5", "滚动收益"],
            ["图6", "相对强弱与超额收益"],
            ["图7", "持有体验箱线图"],
        ],
        columns=["字段", "值"],
    )

    with pd.ExcelWriter(OUT_FILE, engine="openpyxl") as writer:
        core_df.to_excel(writer, sheet_name="核心指标", index=False)
        yearly_out.to_excel(writer, sheet_name="年度收益", index=False)
        rolling_df.to_excel(writer, sheet_name="滚动收益", index=False)
        m_stats_df.to_excel(writer, sheet_name="月度统计", index=False)
        hold_df.to_excel(writer, sheet_name="持有体验", index=False)
        monthly.reset_index().to_excel(writer, sheet_name="月度收益明细", index=False)
        notes.to_excel(writer, sheet_name="说明", index=False)

    print(f"表格: {OUT_FILE}", flush=True)
    print(f"图片目录: {IMG_DIR}", flush=True)
    print(core_df.to_string(index=False), flush=True)


if __name__ == "__main__":
    main()
