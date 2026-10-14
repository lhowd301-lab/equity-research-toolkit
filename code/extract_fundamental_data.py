# -*- coding: utf-8 -*-
"""Extract all fundamental analysis data into organized Excel.
Uses column indices to avoid encoding issues with Chinese column names."""

from pathlib import Path as _PublicPath
_PUBLIC_ROOT = _PublicPath(__file__).resolve().parents[1]
import os
import numpy as np
import pandas as pd

DATA = str(_PUBLIC_ROOT / 'data/创元科技_000551_体检数据_20260806.xlsx')
OUT = str(_PUBLIC_ROOT / 'output/创元科技_基本面分析数据_20260806.xlsx')

xls = pd.ExcelFile(DATA)

pl = xls.parse("利润表")
bs = xls.parse("资产负债表")
cf = xls.parse("现金流量表")
val = xls.parse("估值历史")
_div = xls.parse("分红送转")
sh = xls.parse("股东户数")

# Find annual report type
annual_type = None
for rt in pl["REPORT_TYPE"].unique():
    if "年" in str(rt):
        annual_type = rt
        break
print(f"Annual type found: {repr(annual_type)}")


def filter_annual(df):
    df = df.copy()
    df["REPORT_DATE"] = pd.to_datetime(df["REPORT_DATE"])
    mask = df["REPORT_TYPE"] == annual_type
    df = df[mask].copy()
    df["year"] = df["REPORT_DATE"].dt.year
    df = df.sort_values("REPORT_DATE").drop_duplicates("year", keep="last")
    return df.set_index("year").sort_index()


pl_a = filter_annual(pl)
bs_a = filter_annual(bs)
cf_a = filter_annual(cf)

# --- 1. Profit Statement ---
profit_data = pd.DataFrame(
    {
        "营业总收入(亿)": pl_a["TOTAL_OPERATE_INCOME"] / 1e8,
        "营业总收入同比(%)": pl_a.get("TOTAL_OPERATE_INCOME_YOY", pl_a.get("OPERATE_INCOME_YOY")),
        "营业利润(亿)": pl_a["OPERATE_PROFIT"] / 1e8,
        "归母净利润(亿)": pl_a["PARENT_NETPROFIT"] / 1e8,
        "归母净利润同比(%)": pl_a.get("PARENT_NETPROFIT_YOY"),
        "基本每股收益": pl_a["BASIC_EPS"],
        "营业总成本(亿)": pl_a["TOTAL_OPERATE_COST"] / 1e8,
        "营业成本(亿)": pl_a["OPERATE_COST"] / 1e8,
    },
    index=pl_a.index,
)

# --- 2. Balance Sheet ---
balance_data = pd.DataFrame(
    {
        "总资产(亿)": bs_a["TOTAL_ASSETS"] / 1e8,
        "总负债(亿)": bs_a["TOTAL_LIABILITIES"] / 1e8,
        "股东权益合计(亿)": bs_a["TOTAL_EQUITY"] / 1e8,
        "归属母公司权益(亿)": bs_a["TOTAL_PARENT_EQUITY"] / 1e8,
    },
    index=bs_a.index,
)

# --- 3. Cash Flow ---
cashflow_data = pd.DataFrame(
    {
        "经营活动现金流(亿)": cf_a["NETCASH_OPERATE"] / 1e8,
        "投资活动现金流(亿)": cf_a["NETCASH_INVEST"] / 1e8,
        "筹资活动现金流(亿)": cf_a["NETCASH_FINANCE"] / 1e8,
    },
    index=cf_a.index,
)

# --- 4. Dupont Analysis ---
revenue = profit_data["营业总收入(亿)"]
net_profit = profit_data["归母净利润(亿)"]
total_assets = balance_data["总资产(亿)"]
parent_equity = balance_data["归属母公司权益(亿)"]
total_liab = balance_data["总负债(亿)"]

dupont_data = pd.DataFrame(
    {
        "ROE(%)": net_profit / parent_equity * 100,
        "销售净利率(%)": net_profit / revenue * 100,
        "资产周转率(次)": revenue / total_assets,
        "权益乘数": total_assets / parent_equity,
        "ROA(%)": net_profit / total_assets * 100,
        "资产负债率(%)": total_liab / total_assets * 100,
        "营收增速(%)": revenue.pct_change() * 100,
        "净利润增速(%)": net_profit.pct_change() * 100,
        "毛利率(%)": (revenue - profit_data["营业成本(亿)"]) / revenue * 100,
    },
    index=pl_a.index,
)

# --- 5. Cash Flow Quality ---
cf_quality = pd.DataFrame(
    {
        "经营现金流(亿)": cashflow_data["经营活动现金流(亿)"],
        "经营现金流/净利润": cashflow_data["经营活动现金流(亿)"] / net_profit,
        "自由现金流(亿)": cashflow_data["经营活动现金流(亿)"]
        + cashflow_data["投资活动现金流(亿)"],
    },
    index=cf_a.index,
)

# --- 6. Valuation (use column indices to avoid garbled names) ---
# Col 0=日期, Col 1=收盘价(后复权), Col 3=总市值, Col 7=PE(TTM), Col 9=PB
val_df = pd.DataFrame({
    "日期": pd.to_datetime(val.iloc[:, 0]),
    "收盘价(后复权)": val.iloc[:, 1],
    "PE_TTM": val.iloc[:, 7],
    "PB": val.iloc[:, 9],
})
val_df["year"] = val_df["日期"].dt.year
val_recent = val_df[val_df["year"] >= 2018].copy()

pe_annual = val_recent.groupby("year").agg(
    PE_TTM均值=("PE_TTM", "mean"),
    PE_TTM中位数=("PE_TTM", "median"),
    PE_TTM最低=("PE_TTM", "min"),
    PE_TTM最高=("PE_TTM", "max"),
    PB均值=("PB", "mean"),
    PB中位数=("PB", "median"),
    PB最低=("PB", "min"),
    PB最高=("PB", "max"),
).round(2)

# Percentile
all_pe = val_recent["PE_TTM"].dropna()
all_pb = val_recent["PB"].dropna()
latest_date = val_df["日期"].iloc[-1]
latest_price = val_df["收盘价(后复权)"].iloc[-1]
latest_pe = all_pe.iloc[-1]
latest_pb = all_pb.iloc[-1]
pe_pct = (all_pe < latest_pe).mean() * 100
pb_pct = (all_pb < latest_pb).mean() * 100

# 5Y stats
pe_5y = all_pe.tail(5 * 250)
pb_5y = all_pb.tail(5 * 250)

pct_data = {
    "指标": ["PE(TTM)", "PB"],
    "当前值": [round(latest_pe, 2), round(latest_pb, 3)],
    "历史分位(%)": [round(pe_pct, 1), round(pb_pct, 1)],
    "近5年最低": [round(pe_5y.min(), 2), round(pb_5y.min(), 3)],
    "近5年最高": [round(pe_5y.max(), 2), round(pb_5y.max(), 3)],
    "近5年中位数": [round(pe_5y.median(), 2), round(pb_5y.median(), 3)],
    "近5年20%分位": [round(pe_5y.quantile(0.2), 2), round(pb_5y.quantile(0.2), 3)],
    "近5年80%分位": [round(pe_5y.quantile(0.8), 2), round(pb_5y.quantile(0.8), 3)],
}
pct_df = pd.DataFrame(pct_data)

# --- 7. Dividend (use indices) ---
# Col 0=报告期, Col 5=现金分红金额(10派X元), Col 7=股息率
div = _div.copy()
div["报告期"] = pd.to_datetime(div.iloc[:, 0])
div["每股股利"] = div.iloc[:, 5] / 10  # 10派X元 -> 每股股利
div["股息率"] = div.iloc[:, 7]
div_recent = div[div["报告期"] >= "2015-01-01"].copy()
div_summary = div_recent[["报告期", "每股股利", "股息率"]].copy()

# --- 8. Shareholders (use indices) ---
# Col 0=截止日, Col 2=股东户数总量, Col 4=较上期变化
sh_data = pd.DataFrame({
    "截止日": pd.to_datetime(sh.iloc[:, 0]),
    "股东户数": sh.iloc[:, 2],
    "较上期变化": sh.iloc[:, 4],
})
sh_data = sh_data.sort_values("截止日")

# --- 9. Summary ---
y5 = dupont_data.loc[2021:2025]
roe_5y = y5["ROE(%)"].mean()
nm_5y = y5["销售净利率(%)"].mean()
rg_5y = y5["营收增速(%)"].mean()
nig_5y = y5["净利润增速(%)"].mean()
gm_5y = y5["毛利率(%)"].mean()
debt_2025 = dupont_data.loc[2025, "资产负债率(%)"]
ocf_ni_5y = cf_quality.loc[2021:2025, "经营现金流/净利润"].mean()
at_5y = y5["资产周转率(次)"].mean()
em_5y = y5["权益乘数"].mean()
roa_5y = y5["ROA(%)"].mean()
roe_2025 = dupont_data.loc[2025, "ROE(%)"]
nm_2025 = dupont_data.loc[2025, "销售净利率(%)"]
at_2025 = dupont_data.loc[2025, "资产周转率(次)"]
em_2025 = dupont_data.loc[2025, "权益乘数"]

summary = pd.DataFrame(
    {
        "指标": [
            "最新日期",
            "最新收盘价(后复权)",
            "PE(TTM)当前值",
            "PE(TTM)历史分位(%)-自2018",
            "PE(TTM)近5年20%分位",
            "PE(TTM)近5年中位数",
            "PE(TTM)近5年80%分位",
            "PB当前值",
            "PB历史分位(%)-自2018",
            "PB近5年20%分位",
            "PB近5年中位数",
            "PB近5年80%分位",
            "近5年ROE均值(%)",
            "近5年ROA均值(%)",
            "2025年ROE(%)-杜邦",
            "近5年销售净利率均值(%)",
            "近5年毛利率均值(%)",
            "近5年营收增速均值(%)",
            "近5年净利润增速均值(%)",
            "近5年资产周转率均值(次)",
            "近5年权益乘数均值",
            "2025年资产负债率(%)",
            "近5年经营现金流/净利润均值",
            "近5年自由现金流总和(亿)",
            "2025年营收(亿)",
            "2025年归母净利润(亿)",
            "2025年EPS(元)",
            "最新年度每股股利(元)",
            "最新年度股息率(%)",
            "最新股东户数",
            "一年前股东户数",
            "股东户数1年变化(%)",
            "股东户数趋势",
        ],
        "值": [
            str(latest_date.date()),
            round(latest_price, 2),
            round(latest_pe, 2),
            f"{pe_pct:.1f}%",
            round(pe_5y.quantile(0.2), 2),
            round(pe_5y.median(), 2),
            round(pe_5y.quantile(0.8), 2),
            round(latest_pb, 3),
            f"{pb_pct:.1f}%",
            round(pb_5y.quantile(0.2), 3),
            round(pb_5y.median(), 3),
            round(pb_5y.quantile(0.8), 3),
            round(roe_5y, 2),
            round(roa_5y, 2),
            round(roe_2025, 2),
            round(nm_5y, 2),
            round(gm_5y, 2),
            round(rg_5y, 2),
            round(nig_5y, 2),
            round(at_5y, 2),
            round(em_5y, 2),
            round(debt_2025, 2),
            round(ocf_ni_5y, 2),
            round(cf_quality.loc[2021:2025, "自由现金流(亿)"].sum(), 2),
            round(revenue.loc[2025], 2),
            round(net_profit.loc[2025], 2),
            round(profit_data.loc[2025, "基本每股收益"], 4),
            round(div_recent["每股股利"].iloc[-1], 4)
            if len(div_recent) > 0
            else "N/A",
            round(div_recent["股息率"].iloc[-1], 4)
            if len(div_recent) > 0
            else "N/A",
            int(sh_data["股东户数"].iloc[-1]) if len(sh_data) > 0 else "N/A",
            int(sh_data["股东户数"].iloc[-5]) if len(sh_data) > 5 else "N/A",
            f"{(sh_data['股东户数'].iloc[-1]/sh_data['股东户数'].iloc[-5]-1)*100:.1f}%"
            if len(sh_data) > 5
            else "N/A",
            "集中(户数下降)"
            if len(sh_data) >= 5
            and sh_data["股东户数"].iloc[-1] < sh_data["股东户数"].iloc[-5]
            else "分散",
        ],
    }
)

# --- 10. Year-by-year Dupont Decomposition table ---
dupont_detail = dupont_data.loc[2015:2025, [
    "ROE(%)", "销售净利率(%)", "资产周转率(次)", "权益乘数", "ROA(%)",
    "资产负债率(%)", "营收增速(%)", "净利润增速(%)", "毛利率(%)"
]].round(2)

# --- Write Excel ---
with pd.ExcelWriter(OUT, engine="openpyxl") as writer:
    profit_data.round(2).to_excel(writer, sheet_name="1_利润表")
    balance_data.round(2).to_excel(writer, sheet_name="2_资产负债表")
    cashflow_data.round(2).to_excel(writer, sheet_name="3_现金流量表")
    dupont_data.round(2).to_excel(writer, sheet_name="4_杜邦分析_全历史")
    dupont_detail.to_excel(writer, sheet_name="5_杜邦分析_2015起")
    cf_quality.round(2).to_excel(writer, sheet_name="6_现金流质量")
    pe_annual.to_excel(writer, sheet_name="7_估值年度汇总")
    pct_df.to_excel(writer, sheet_name="8_估值分位")
    div_summary.to_excel(writer, sheet_name="9_分红汇总")
    summary.to_excel(writer, sheet_name="10_核心指标汇总", index=False)

print(f"\nExcel saved to: {OUT}")
print(f"\n=== 核心指标 ===")
print(f"日期: {latest_date.date()}")
print(f"PE(TTM): {latest_pe:.2f}, 历史分位: {pe_pct:.1f}%")
print(f"PB: {latest_pb:.3f}, 历史分位: {pb_pct:.1f}%")
print(f"PE 5Y: 20%={pe_5y.quantile(0.2):.1f}, 50%={pe_5y.median():.1f}, 80%={pe_5y.quantile(0.8):.1f}")
print(f"PB 5Y: 20%={pb_5y.quantile(0.2):.3f}, 50%={pb_5y.median():.3f}, 80%={pb_5y.quantile(0.8):.3f}")
print(f"5Y ROE均值: {roe_5y:.2f}%")
print(f"2025 ROE: {roe_2025:.2f}% (杜邦: 净利率{nm_2025:.1f}% x 周转{at_2025:.2f} x 杠杆{em_2025:.2f})")
print(f"5Y 净利率均值: {nm_5y:.2f}%")
print(f"5Y 毛利率均值: {gm_5y:.2f}%")
print(f"5Y 营收增速均值: {rg_5y:.2f}%")
print(f"5Y 净利增速均值: {nig_5y:.2f}%")
print(f"5Y 资产周转率均值: {at_5y:.2f}次")
print(f"5Y OCF/NI均值: {ocf_ni_5y:.2f}")
print(f"资产负债率(2025): {debt_2025:.2f}%")
print(f"股东户数: {int(sh_data['股东户数'].iloc[-1])}")
print("Done.")
