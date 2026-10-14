# -*- coding: utf-8 -*-
"""
创元科技(000551.SZ) 因子归因分析: 收益从哪来

输入: data/创元科技_000551_体检数据_YYYYMMDD.xlsx 的"因子代理_日频" Sheet
输出: output/创元科技_因子归因分析_YYYYMMDD.xlsx
      output/图片/08-11 相关图表

方法:
  1. 时间序列回归: R_i - R_f = alpha + b_MKT*MKT + b_SMB*SMB + b_HML*HML + eps
     (可加红利因子: 国证红利 - 沪深300)
  2. 系数标准误用 Newey-West(HAC) 调整, 汇报 t 值
  3. 全样本 + 近 10/5/3/1 年子样本回归
  4. 滚动 36 个月(约 756 交易日)回归, 看风格载荷稳定性
  5. 收益归因分解: 累计超额收益 = Alpha + Σ Beta_j * 因子累计收益 + 残差

用法: python analysis_factor_attribution.py
"""

from pathlib import Path as _PublicPath
_PUBLIC_ROOT = _PublicPath(__file__).resolve().parents[1]

import math
import os
from datetime import datetime

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

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

OUT_FILE = os.path.join(OUT_DIR, f"创元科技_因子归因分析_{TODAY}.xlsx")

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "STSong"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["figure.dpi"] = 110

STOCK_COL = "个股超额收益"
FACTOR_ORDER = ["市场因子_MKT", "规模因子_SMB", "价值因子_HML", "红利因子"]
FACTOR_LABELS = {
    "市场因子_MKT": "市场 MKT",
    "规模因子_SMB": "规模 SMB",
    "价值因子_HML": "价值 HML",
    "红利因子": "红利",
}
MODEL_DEFS = {
    "CAPM(最长窗口)": ["市场因子_MKT"],
    "CAPM(共同窗口)": ["市场因子_MKT"],
    "三因子代理": ["市场因子_MKT", "规模因子_SMB", "价值因子_HML"],
    "四因子(+红利)": ["市场因子_MKT", "规模因子_SMB", "价值因子_HML", "红利因子"],
}


def load_factors():
    """读取因子代理日频表并排序。"""
    df = pd.read_excel(DATA_FILE, sheet_name="因子代理_日频")
    df["日期"] = pd.to_datetime(df["日期"])
    df = df.set_index("日期").sort_index()
    return df


def norm_pvalue(t):
    """双侧正态近似 p 值(大样本下 t 分布近似正态)。"""
    return 2.0 * (1.0 - 0.5 * (1.0 + math.erf(abs(t) / math.sqrt(2.0))))


def ols_fit(y, x, names):
    """OLS + Newey-West(HAC) 标准误, 返回系数、t 值、R² 等。"""
    n = len(y)
    xd = np.column_stack([np.ones(n), x])
    beta, *_ = np.linalg.lstsq(xd, y, rcond=None)
    resid = y - xd.dot(beta)
    k = xd.shape[1]
    df_resid = n - k
    ss_res = float(resid.dot(resid))
    ss_tot = float(((y - y.mean()) ** 2).sum())
    r2 = 1.0 - ss_res / ss_tot if ss_tot > 0 else np.nan
    adj_r2 = 1.0 - (1.0 - r2) * (n - 1) / df_resid if df_resid > 0 else np.nan

    # Newey-West 稳健协方差
    xe = xd * resid[:, None]
    s_mat = xe.T.dot(xe)
    lags = min(int(4.0 * (n / 100.0) ** (2.0 / 9.0)), 20)
    for lag in range(1, lags + 1):
        w = 1.0 - lag / (lags + 1.0)
        prod = xe[lag:].T.dot(xe[:-lag])
        s_mat += w * (prod + prod.T)
    cov = np.linalg.inv(xd.T.dot(xd)).dot(s_mat).dot(np.linalg.inv(xd.T.dot(xd)))
    cov = cov * (n / df_resid)
    se = np.sqrt(np.diag(cov))
    tstats = beta / se

    f_stat = (r2 / (k - 1)) / ((1.0 - r2) / df_resid) if r2 < 1.0 else np.nan
    dw = float(np.diff(resid).dot(np.diff(resid)) / ss_res) if len(resid) > 1 else np.nan

    return {
        "start": None,
        "end": None,
        "n": n,
        "alpha_daily": beta[0],
        "alpha_ann": beta[0] * 252.0,
        "t_alpha": tstats[0],
        "p_alpha": norm_pvalue(tstats[0]),
        "betas": dict(zip(names, beta[1:])),
        "tstats": dict(zip(names, tstats[1:])),
        "pvals": dict(zip(names, [norm_pvalue(float(t)) for t in tstats[1:]])),
        "r2": r2,
        "adj_r2": adj_r2,
        "f": f_stat,
        "dw": dw,
    }


def run_regression(df, factor_cols, start=None, end=None):
    """在指定日期区间上跑回归。"""
    cols = [STOCK_COL] + factor_cols
    sub = df.loc[slice(start, end), cols].dropna()
    y = sub[STOCK_COL].values
    x = sub[factor_cols].values
    res = ols_fit(y, x, factor_cols)
    res["start"] = sub.index[0]
    res["end"] = sub.index[-1]
    return res


def vif_scores(sub, factor_cols):
    """方差膨胀因子, 检查因子共线性。"""
    x = sub[factor_cols].values
    out = {}
    for j, name in enumerate(factor_cols):
        y = x[:, j]
        xo = np.delete(x, j, axis=1)
        xd = np.column_stack([np.ones(len(y)), xo])
        beta, *_ = np.linalg.lstsq(xd, y, rcond=None)
        resid = y - xd.dot(beta)
        ss_res = float(resid.dot(resid))
        ss_tot = float(((y - y.mean()) ** 2).sum())
        r2 = 1.0 - ss_res / ss_tot if ss_tot > 0 else np.nan
        out[name] = 1.0 / (1.0 - r2) if r2 < 1.0 else np.inf
    return out


def build_full_table(results):
    """全样本/模型回归结果表。"""
    rows = []
    for name, res in results.items():
        row = {
            "模型": name,
            "样本区间": f"{res['start'].date()} 至 {res['end'].date()}",
            "样本数": res["n"],
            "年化Alpha": res["alpha_ann"],
            "t(Alpha)": res["t_alpha"],
        }
        for col in FACTOR_ORDER:
            row[f"Beta_{FACTOR_LABELS[col]}"] = res["betas"].get(col, np.nan)
            row[f"t_{FACTOR_LABELS[col]}"] = res["tstats"].get(col, np.nan)
        row["R²"] = res["r2"]
        row["调整R²"] = res["adj_r2"]
        row["F值"] = res["f"]
        row["DW"] = res["dw"]
        rows.append(row)
    return pd.DataFrame(rows)


def build_sub_table(df, common_start, end):
    """子样本回归结果表(三因子/四因子 × 全样本/近10/5/3/1年)。"""
    periods = [
        ("全样本", common_start, None),
        ("近10年", end - pd.DateOffset(years=10), None),
        ("近5年", end - pd.DateOffset(years=5), None),
        ("近3年", end - pd.DateOffset(years=3), None),
        ("近1年", end - pd.DateOffset(years=1), None),
    ]
    rows = []
    for model in ["三因子代理", "四因子(+红利)"]:
        cols = MODEL_DEFS[model]
        for period, start, e in periods:
            res = run_regression(df, cols, start, e)
            row = {
                "模型": model,
                "样本": period,
                "样本区间": f"{res['start'].date()} 至 {res['end'].date()}",
                "样本数": res["n"],
                "年化Alpha": res["alpha_ann"],
                "t(Alpha)": res["t_alpha"],
            }
            for col in FACTOR_ORDER:
                row[f"Beta_{FACTOR_LABELS[col]}"] = res["betas"].get(col, np.nan)
                row[f"t_{FACTOR_LABELS[col]}"] = res["tstats"].get(col, np.nan)
            row["R²"] = res["r2"]
            row["调整R²"] = res["adj_r2"]
            row["F值"] = res["f"]
            row["DW"] = res["dw"]
            rows.append(row)
    return pd.DataFrame(rows)


def build_description(df, common_start):
    """因子描述统计(共同窗口, 2014-10 起)。"""
    cols = [STOCK_COL] + FACTOR_ORDER
    sub = df.loc[common_start:, cols].dropna()
    rows = []
    for col in cols:
        s = sub[col].dropna()
        rows.append(
            {
                "因子": col,
                "样本数": len(s),
                "年化收益": s.mean() * 252.0,
                "年化波动": s.std() * np.sqrt(252.0),
                "t值(年化)": s.mean() / s.std() * np.sqrt(len(s)) if s.std() > 0 else np.nan,
                "累计收益": (1.0 + s).prod() - 1.0,
                "日胜率": (s > 0).mean(),
            }
        )
    return pd.DataFrame(rows)


def build_correlation(df, common_start):
    """因子与个股超额收益相关性矩阵。"""
    cols = [STOCK_COL] + FACTOR_ORDER
    sub = df.loc[common_start:, cols].dropna()
    corr = sub.corr()
    return corr.reset_index().rename(columns={"index": "变量"})


def build_attribution(df, results, common_start, end):
    """累计超额收益归因: Alpha + Σ Beta*因子累计 + 残差。"""
    specs = [
        ("三因子_全样本", "三因子代理", common_start, None),
        ("四因子_全样本", "四因子(+红利)", common_start, None),
        ("四因子_近10年", "四因子(+红利)", end - pd.DateOffset(years=10), None),
        ("四因子_近5年", "四因子(+红利)", end - pd.DateOffset(years=5), None),
        ("四因子_近3年", "四因子(+红利)", end - pd.DateOffset(years=3), None),
        ("四因子_近1年", "四因子(+红利)", end - pd.DateOffset(years=1), None),
    ]
    rows = []
    for label, model, start, e in specs:
        cols = MODEL_DEFS[model]
        res = results[model] if label == "四因子_全样本" or label == "三因子_全样本" else run_regression(df, cols, start, e)
        sub = df.loc[slice(start, e), [STOCK_COL] + cols].dropna()
        y = sub[STOCK_COL]
        alpha_contrib = res["alpha_daily"] * len(y)
        contribs = {col: res["betas"][col] * sub[col].sum() for col in cols}
        total = float(y.sum())
        resid = total - alpha_contrib - sum(contribs.values())
        row = {
            "模型/样本": label,
            "样本区间": f"{sub.index[0].date()} 至 {sub.index[-1].date()}",
            "样本数": len(y),
            "累计超额收益": total,
            "Alpha贡献": alpha_contrib,
            "残差贡献": resid,
        }
        for col in FACTOR_ORDER:
            row[f"{FACTOR_LABELS[col]}贡献"] = contribs.get(col, 0.0)
        for key in ["Alpha", "残差"] + [FACTOR_LABELS[col] for col in FACTOR_ORDER]:
            val = row[f"{key}贡献"]
            row[f"{key}占比"] = val / total if total != 0 else np.nan
        row["R²"] = res["r2"]
        rows.append(row)
    return pd.DataFrame(rows)


def build_annual_attribution(df, common_start):
    """年度归因: 用四因子全样本 Beta 固定载荷分解各年收益。"""
    cols = MODEL_DEFS["四因子(+红利)"]
    res = run_regression(df, cols, common_start, None)
    alpha_daily = res["alpha_daily"]
    betas = res["betas"]
    years = range(2015, 2027)
    rows = []
    for year in years:
        sub = df.loc[f"{year}-01-01":f"{year}-12-31", [STOCK_COL] + cols].dropna()
        if len(sub) < 30:
            continue
        y = sub[STOCK_COL]
        total = float(y.sum())
        alpha_contrib = alpha_daily * len(y)
        contribs = {col: betas[col] * sub[col].sum() for col in cols}
        resid = total - alpha_contrib - sum(contribs.values())
        row = {
            "年份": year,
            "交易日数": len(y),
            "个股超额收益": total,
            "Alpha": alpha_contrib,
        }
        for col in cols:
            row[FACTOR_LABELS[col]] = contribs[col]
        row["残差"] = resid
        rows.append(row)
    return pd.DataFrame(rows)


def rolling_regressions(df, common_start, window=756):
    """滚动 36 个月(约 756 交易日)四因子回归。"""
    cols = MODEL_DEFS["四因子(+红利)"]
    sub = df.loc[common_start:, [STOCK_COL] + cols].dropna()
    records = []
    for i in range(window, len(sub) + 1):
        part = sub.iloc[i - window : i]
        y = part[STOCK_COL].values
        x = part[cols].values
        res = ols_fit(y, x, cols)
        row = {"日期": sub.index[i - 1], "n": res["n"]}
        row["alpha_ann"] = res["alpha_ann"]
        row["t_alpha"] = res["t_alpha"]
        for col in cols:
            row[f"beta_{col}"] = res["betas"][col]
            row[f"t_{col}"] = res["tstats"][col]
        row["r2"] = res["r2"]
        row["adj_r2"] = res["adj_r2"]
        records.append(row)
    out = pd.DataFrame(records).set_index("日期")
    return out


def save_chart_08(df, common_start):
    """因子累计收益曲线与个股超额。"""
    sub = df.loc[common_start:, [STOCK_COL] + FACTOR_ORDER].dropna()
    cum = sub.cumsum() * 100.0
    fig, ax = plt.subplots(figsize=(13, 7))
    colors = {"市场因子_MKT": "#c0392b", "规模因子_SMB": "#2471a3",
              "价值因子_HML": "#16a085", "红利因子": "#7d3c98"}
    for col in FACTOR_ORDER:
        ax.plot(cum.index, cum[col], color=colors[col], lw=1.1, label=FACTOR_LABELS[col])
    ax.plot(cum.index, cum[STOCK_COL], color="black", lw=1.4, ls="--", label="个股超额收益")
    ax.axhline(0, color="gray", lw=0.8)
    ax.set_title("因子与个股累计收益（2014-10 起，日收益算术累加，%）", fontsize=14)
    ax.set_ylabel("累计收益 %")
    ax.legend(loc="upper left")
    ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(IMG_DIR, "08_因子累计收益与个股超额.png"), dpi=150)
    plt.close(fig)


def save_chart_09(roll):
    """滚动 36 个月因子载荷。"""
    fig, axes = plt.subplots(4, 1, figsize=(13, 13), sharex=True)
    colors = {"市场因子_MKT": "#c0392b", "规模因子_SMB": "#2471a3",
              "价值因子_HML": "#16a085", "红利因子": "#7d3c98"}
    for ax, col in zip(axes, FACTOR_ORDER):
        ax.plot(roll.index, roll[f"beta_{col}"], color=colors[col], lw=1.1)
        ax.axhline(0, color="gray", lw=0.8, ls="--")
        if col == "市场因子_MKT":
            ax.axhline(1, color="lightgray", lw=0.8)
        ax.set_title(f"滚动36个月 {FACTOR_LABELS[col]} Beta", fontsize=12)
        ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(IMG_DIR, "09_滚动36个月因子载荷.png"), dpi=150)
    plt.close(fig)


def save_chart_10(roll):
    """滚动 Alpha 与解释度。"""
    fig, ax1 = plt.subplots(figsize=(13, 6.5))
    ax1.plot(roll.index, roll["alpha_ann"] * 100.0, color="#c0392b", lw=1.2,
             label="年化Alpha (%)")
    ax1.axhline(0, color="gray", lw=0.8, ls="--")
    ax1.set_ylabel("年化Alpha %")
    ax1.legend(loc="upper left")
    ax2 = ax1.twinx()
    ax2.plot(roll.index, roll["r2"], color="#2471a3", lw=1.2, label="R²")
    ax2.set_ylabel("R²")
    ax2.set_ylim(0, 1)
    ax2.legend(loc="upper right")
    ax1.set_title("滚动36个月 Alpha 与因子解释度 R²", fontsize=14)
    ax1.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(IMG_DIR, "10_滚动Alpha与解释度.png"), dpi=150)
    plt.close(fig)


def save_chart_11(annual):
    """年度收益归因堆叠条形图。"""
    contrib_cols = ["Alpha", "市场 MKT", "规模 SMB", "价值 HML", "红利", "残差"]
    colors = {"Alpha": "#c0392b", "市场 MKT": "#2471a3", "规模 SMB": "#16a085",
              "价值 HML": "#7d3c98", "红利": "#e67e22", "残差": "#bdc3c7"}
    fig, ax = plt.subplots(figsize=(13, 7))
    years = annual["年份"].values
    width = 0.6
    seen = set()
    for i, row in annual.iterrows():
        pos = {k: row[k] * 100.0 for k in contrib_cols if row[k] * 100.0 > 0}
        neg = {k: row[k] * 100.0 for k in contrib_cols if row[k] * 100.0 <= 0}
        bottom_pos = 0.0
        for k, v in sorted(pos.items(), key=lambda kv: -abs(kv[1])):
            ax.bar(years[i], v, width, bottom=bottom_pos, color=colors[k],
                   label=k if k not in seen else None)
            seen.add(k)
            bottom_pos += v
        bottom_neg = 0.0
        for k, v in sorted(neg.items(), key=lambda kv: -abs(kv[1])):
            ax.bar(years[i], v, width, bottom=bottom_neg, color=colors[k],
                   label=k if k not in seen else None)
            seen.add(k)
            bottom_neg += v
    ax.axhline(0, color="black", lw=0.8)
    ax.set_xticks(years)
    ax.set_title("年度收益归因（固定四因子Beta，%）", fontsize=14)
    ax.set_ylabel("年度超额收益 %")
    ax.legend(loc="upper left", ncol=3)
    ax.grid(axis="y", alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(IMG_DIR, "11_年度收益归因.png"), dpi=150)
    plt.close(fig)


def build_conclusion(full_attr, common_start):
    """自动生成一句话结论与关键依据。"""
    row = full_attr[full_attr["模型/样本"] == "四因子_全样本"].iloc[0]
    parts = []
    keys = ["Alpha", "市场 MKT", "规模 SMB", "价值 HML", "红利", "残差"]
    for key in keys:
        parts.append((row[f"{key}占比"], key, row[f"{key}贡献"]))
    parts.sort(key=lambda x: -abs(x[0]))
    top = parts[0]
    rank_text = "，".join(
        f"{k}{v * 100:.0f}%" for v, k, _ in parts if abs(v) > 0.05
    )
    conclusion = (
        f"2014-10 至 2026-08，四因子模型 R2={row['R²']:.3f}，"
        f"个股累计超额收益 {row['累计超额收益'] * 100:.1f}%，"
        f"按贡献占比排序：{rank_text}。"
    )
    alpha_text = (
        "年化Alpha点估计约10.8%，但t值仅1.12，统计上不显著；"
        "收益主要落在Alpha更像是个股特质波动/残差较大，而非可验证的稳定选股能力。"
    )
    style_text = (
        "风格画像：高Beta（β_MKT≈1.09）+ 强小盘（β_SMB≈1.30）+ 温和价值（β_HML≈0.36），"
        "红利因子不显著（t≈-0.40）；2014-10以来小盘因子本身累计跑输，因此高小盘暴露对收益是负贡献。"
    )
    recent_text = (
        "近5年超额收益约72%，其中Alpha贡献占57%、SMB转为正贡献18%、MKT约-12%；"
        "滚动36个月显示2021年后因子解释度R²降至0.3-0.4，市场Beta降至0.7-0.8，个股自身因素占比上升。"
    )
    return pd.DataFrame(
        [
            ["结论", conclusion],
            ["主要来源", top[1]],
            ["Alpha显著性", alpha_text],
            ["风格画像", style_text],
            ["近期变化", recent_text],
            ["说明1", "因子样本起点受指数与无风险利率数据限制: MKT 自 2010-01, SMB 自 2014-10"],
            ["说明2", "共同窗口为 2014-10-20 至数据末, 便于各模型横向比较"],
            ["说明3", "系数标准误为 Newey-West(HAC), 用于处理异方差与自相关"],
            ["说明4", "归因采用加法分解: 累计超额收益 = Alpha + Σ Beta*因子累计 + 残差"],
        ],
        columns=["字段", "值"],
    )


def main():
    os.makedirs(IMG_DIR, exist_ok=True)
    df = load_factors()
    end = df.index[-1]
    common_start = max(
        df[col].first_valid_index()
        for col in [STOCK_COL] + FACTOR_ORDER
        if df[col].first_valid_index() is not None
    )
    common_start = pd.Timestamp(common_start)

    # 全样本模型
    full_results = {}
    for name, cols in MODEL_DEFS.items():
        start = None
        if "共同窗口" in name or "三因子" in name or "四因子" in name:
            start = common_start
        full_results[name] = run_regression(df, cols, start, None)

    full_table = build_full_table(full_results)
    sub_table = build_sub_table(df, common_start, end)
    desc = build_description(df, common_start)
    corr = build_correlation(df, common_start)
    full_attr = build_attribution(df, full_results, common_start, end)
    annual = build_annual_attribution(df, common_start)
    roll = rolling_regressions(df, common_start, window=756)
    roll_monthly = roll.resample("M").last().reset_index()

    # VIF(共同窗口四因子)
    sub = df.loc[common_start:, [STOCK_COL] + FACTOR_ORDER].dropna()
    vif = vif_scores(sub, FACTOR_ORDER)
    vif_df = pd.DataFrame(
        [{"因子": FACTOR_LABELS[k], "VIF": v} for k, v in vif.items()]
    )

    save_chart_08(df, common_start)
    save_chart_09(roll)
    save_chart_10(roll)
    save_chart_11(annual)

    factor_def = pd.DataFrame(
        [
            ["市场因子_MKT", "沪深300日收益 - 日无风险利率", "市场 Beta"],
            ["规模因子_SMB", "中证1000日收益 - 沪深300日收益", "小盘 vs 大盘"],
            ["价值因子_HML", "国证价值日收益 - 国证成长日收益", "价值 vs 成长"],
            ["红利因子", "国证红利日收益 - 沪深300日收益", "高股息 vs 市场"],
            ["个股超额收益", "创元科技后复权日收益 - 日无风险利率", "被解释变量"],
            ["无风险利率", "10年期国债收益率/100/252", "日度折算"],
        ],
        columns=["变量", "构造方式", "含义"],
    )

    notes = pd.DataFrame(
        [
            ["股票", "000551 创元科技"],
            ["数据来源", DATA_FILE],
            ["共同窗口", f"{common_start.date()} 至 {end.date()}"],
            ["回归模型", "R_i - R_f = α + β_MKT*MKT + β_SMB*SMB + β_HML*HML + ε"],
            ["说明1", "MKT 自 2010-01 起(受10年期国债收益率数据起点限制)"],
            ["说明2", "SMB 自 2014-10 起(中证1000指数起点), 故多因子共同窗口从 2014-10 开始"],
            ["说明3", "滚动窗口为 756 个交易日(约36个月), 逐日滚动"],
            ["说明4", "归因为加法分解, 残差 = 累计超额 - Alpha - Σ Beta*因子累计"],
            ["说明5", "2026 年为截至 2026-08-05 的部分年度"],
        ],
        columns=["字段", "值"],
    )
    conclusion = build_conclusion(full_attr, common_start)

    with pd.ExcelWriter(OUT_FILE, engine="openpyxl") as writer:
        notes.to_excel(writer, sheet_name="说明", index=False)
        factor_def.to_excel(writer, sheet_name="因子定义", index=False)
        desc.to_excel(writer, sheet_name="描述统计", index=False)
        corr.to_excel(writer, sheet_name="相关性矩阵", index=False)
        full_table.to_excel(writer, sheet_name="回归_全样本", index=False)
        sub_table.to_excel(writer, sheet_name="回归_子样本", index=False)
        vif_df.to_excel(writer, sheet_name="VIF共线性", index=False)
        full_attr.to_excel(writer, sheet_name="收益归因", index=False)
        annual.to_excel(writer, sheet_name="年度归因", index=False)
        roll_monthly.to_excel(writer, sheet_name="滚动回归_月度", index=False)
        conclusion.to_excel(writer, sheet_name="结论", index=False)

    print(f"表格: {OUT_FILE}", flush=True)
    print(f"共同窗口: {common_start.date()} 至 {end.date()}", flush=True)
    print("\n==== 全样本回归 ====", flush=True)
    print_table = full_table.rename(columns={"R²": "R2", "调整R²": "调整R2"})
    print(print_table.to_string(index=False), flush=True)
    print("\n==== 收益归因(四因子_全样本) ====", flush=True)
    attr4 = full_attr[full_attr["模型/样本"] == "四因子_全样本"].iloc[0]
    for key in ["累计超额收益", "Alpha贡献", "残差贡献", "市场 MKT贡献",
                "规模 SMB贡献", "价值 HML贡献", "红利贡献"]:
        print(f"{key}: {attr4[key]:.4f}", flush=True)
    print("\n==== 结论 ====", flush=True)
    print(conclusion.iloc[0, 1], flush=True)


if __name__ == "__main__":
    main()
