# -*- coding: utf-8 -*-
"""
创元科技(000551.SZ) 长期持有体检数据抓取脚本

输出: data/创元科技_000551_体检数据_YYYYMMDD.xlsx
数据源: akshare (东方财富 / 新浪 / 申万宏源 / 同花顺 等公开接口)

用法: python fetch_stock_health_data.py
"""

from pathlib import Path as _PublicPath
_PUBLIC_ROOT = _PublicPath(__file__).resolve().parents[1]

import os
import time
from datetime import datetime

import akshare as ak
import numpy as np
import pandas as pd

STOCK = "000551"
SYMBOL_SZ = "sz000551"
START = "19900101"
END = datetime.now().strftime("%Y%m%d")
TODAY = datetime.now().strftime("%Y%m%d")

BASE_DIR = str(_PUBLIC_ROOT)
DATA_DIR = os.path.join(BASE_DIR, "data")
OUT_FILE = os.path.join(DATA_DIR, f"创元科技_{STOCK}_体检数据_{TODAY}.xlsx")

FAILURES = []


def fetch(label, fn, tries=3, sleep=3):
    """带重试的抓取封装, 失败记录到 FAILURES 并返回 None。"""
    for i in range(tries):
        try:
            df = fn()
            if df is None or len(df) == 0:
                raise ValueError("empty result")
            print(f"[OK] {label} rows={len(df)}", flush=True)
            return df
        except Exception as e:  # noqa: BLE001
            print(f"[RETRY {i + 1}/{tries}] {label}: {type(e).__name__} {str(e)[:140]}", flush=True)
            if i < tries - 1:
                time.sleep(sleep * (i + 1))
    FAILURES.append(label)
    print(f"[FAIL] {label}", flush=True)
    return None


def to_dt(series):
    return pd.to_datetime(series)


def get_hist(adjust):
    """东财日线为主, 失败时退回新浪日线。"""

    def em():
        return ak.stock_zh_a_hist(
            symbol=STOCK, period="daily", start_date=START, end_date=END, adjust=adjust
        )

    def sina():
        return ak.stock_zh_a_daily(symbol=SYMBOL_SZ, start_date=START, end_date=END, adjust=adjust)

    try:
        return em()
    except Exception:
        return sina()


def normalize_hist(df):
    """统一东财/新浪日线字段名。"""
    df = df.copy()
    rename_map = {
        "date": "日期",
        "open": "开盘",
        "high": "最高",
        "low": "最低",
        "close": "收盘",
        "volume": "成交量",
        "amount": "成交额",
        "outstanding_share": "流通股本",
        "turnover": "换手率",
    }
    df = df.rename(columns=rename_map)
    return df


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    t0 = time.time()
    sheets = {}

    # ---------- 行情 ----------
    hist_qfq = fetch("日线行情_前复权", lambda: get_hist("qfq"))
    hist_hfq = fetch("日线行情_后复权", lambda: get_hist("hfq"))
    if hist_qfq is not None:
        hist_qfq = normalize_hist(hist_qfq)
        sheets["日线行情_前复权"] = hist_qfq
    if hist_hfq is not None:
        hist_hfq = normalize_hist(hist_hfq)
        sheets["日线行情_后复权"] = hist_hfq

    # ---------- 指数 ----------
    index_map = {
        "sh000300": "沪深300",
        "sh000905": "中证500",
        "sh000852": "中证1000",
        "sz399370": "国证成长",
        "sz399371": "国证价值",
        "sz399321": "国证红利",
    }
    idx_frames = []
    for code, name in index_map.items():
        df = fetch(f"指数_{name}", lambda c=code: ak.stock_zh_index_daily(symbol=c), tries=2, sleep=2)
        if df is not None:
            df = df.copy()
            df["指数代码"] = code
            df["指数名称"] = name
            idx_frames.append(df)
    if idx_frames:
        idx_long = pd.concat(idx_frames, ignore_index=True)
        idx_long["日期"] = to_dt(idx_long["date"])
        idx_long = idx_long.rename(columns={"open": "开盘", "high": "最高", "low": "最低", "close": "收盘", "volume": "成交量"})
        sheets["基准指数日线"] = idx_long[["日期", "指数代码", "指数名称", "开盘", "最高", "最低", "收盘", "成交量"]]

    sw = fetch("申万行业指数_环保", lambda: ak.index_hist_sw(symbol="801970", period="day"), tries=2, sleep=2)
    if sw is not None:
        sheets["申万行业指数日线"] = sw

    # ---------- 财务与估值 ----------
    value = fetch("估值历史", lambda: ak.stock_value_em(symbol=STOCK))
    if value is not None:
        sheets["估值历史"] = value

    fin_ind = fetch("财务指标_新浪", lambda: ak.stock_financial_analysis_indicator(symbol=STOCK, start_year="2010"))
    if fin_ind is not None:
        sheets["财务指标_新浪"] = fin_ind

    fin_abs = fetch("财务摘要_东财", lambda: ak.stock_financial_abstract(symbol=STOCK))
    if fin_abs is not None:
        sheets["财务摘要_东财"] = fin_abs

    balance = fetch("资产负债表", lambda: ak.stock_balance_sheet_by_report_em(symbol="SZ" + STOCK))
    if balance is not None:
        sheets["资产负债表"] = balance

    profit = fetch("利润表", lambda: ak.stock_profit_sheet_by_report_em(symbol="SZ" + STOCK))
    if profit is not None:
        sheets["利润表"] = profit

    cashflow = fetch("现金流量表", lambda: ak.stock_cash_flow_sheet_by_report_em(symbol="SZ" + STOCK))
    if cashflow is not None:
        sheets["现金流量表"] = cashflow

    fhps = fetch("分红送转", lambda: ak.stock_fhps_detail_em(symbol=STOCK))
    if fhps is not None:
        sheets["分红送转"] = fhps

    # ---------- 股东 ----------
    gdhs = fetch("股东户数", lambda: ak.stock_zh_a_gdhs_detail_em(symbol=STOCK))
    if gdhs is not None:
        sheets["股东户数"] = gdhs

    report_dates = [
        "20240331", "20240630", "20240930", "20241231",
        "20250331", "20250630", "20250930", "20251231", "20260331",
    ]
    top10_frames, free_top10_frames = [], []
    for rd in report_dates:
        df1 = fetch(f"十大股东_{rd}", lambda d=rd: ak.stock_gdfx_top_10_em(symbol=SYMBOL_SZ, date=d), tries=1, sleep=1)
        if df1 is not None:
            df1 = df1.copy()
            df1["报告期"] = rd
            top10_frames.append(df1)
        time.sleep(0.3)
        df2 = fetch(f"十大流通股东_{rd}", lambda d=rd: ak.stock_gdfx_free_top_10_em(symbol=SYMBOL_SZ, date=d), tries=1, sleep=1)
        if df2 is not None:
            df2 = df2.copy()
            df2["报告期"] = rd
            free_top10_frames.append(df2)
        time.sleep(0.3)
    if top10_frames:
        sheets["十大股东_历史"] = pd.concat(top10_frames, ignore_index=True)
    if free_top10_frames:
        sheets["十大流通股东_历史"] = pd.concat(free_top10_frames, ignore_index=True)

    main_holder = fetch("主要股东_新浪", lambda: ak.stock_main_stock_holder(stock=STOCK), tries=2, sleep=2)
    if main_holder is not None:
        sheets["主要股东_新浪"] = main_holder

    circ_holder = fetch("流通股东_新浪", lambda: ak.stock_circulate_stock_holder(symbol=STOCK), tries=2, sleep=2)
    if circ_holder is not None:
        sheets["流通股东_新浪"] = circ_holder

    fund_holder = fetch("基金持股_新浪", lambda: ak.stock_fund_stock_holder(symbol=STOCK), tries=2, sleep=2)
    if fund_holder is not None:
        sheets["基金持股_新浪"] = fund_holder

    holder_change = fetch("股东持股变动_同花顺", lambda: ak.stock_shareholder_change_ths(symbol=STOCK), tries=2, sleep=2)
    if holder_change is not None:
        sheets["股东持股变动_同花顺"] = holder_change

    # ---------- 融资融券(季度末快照) ----------
    margin_rows = []
    if hist_hfq is not None:
        trade_dates = pd.to_datetime(hist_hfq["日期"]).sort_values().values
        quarter_ends = pd.date_range("2019-03-31", END, freq="QE")
        for qe in quarter_ends:
            candidates = trade_dates[trade_dates <= np.datetime64(qe)]
            if len(candidates) == 0:
                continue
            d = pd.Timestamp(candidates[-1]).strftime("%Y%m%d")
            df = fetch(f"融资融券_{d}", lambda dd=d: ak.stock_margin_detail_szse(date=dd), tries=2, sleep=1)
            if df is None:
                continue
            row = df[df["证券代码"].astype(str) == STOCK]
            if len(row) == 0:
                continue
            row = row.copy()
            row["报告期"] = qe.strftime("%Y-%m-%d")
            margin_rows.append(row)
            time.sleep(0.2)
    if margin_rows:
        sheets["融资融券_季度"] = pd.concat(margin_rows, ignore_index=True)

    # ---------- 北向持股 ----------
    hsgt = fetch("北向持股", lambda: ak.stock_hsgt_individual_em(stock=STOCK), tries=2, sleep=2)
    if hsgt is not None:
        hsgt = hsgt.copy()
        hsgt["股票代码"] = STOCK
        sheets["北向持股"] = hsgt

    hsgt_flow = fetch("北向资金总额", lambda: ak.stock_hsgt_hist_em(symbol="北向资金"), tries=2, sleep=2)
    if hsgt_flow is not None:
        sheets["北向资金总额"] = hsgt_flow

    # ---------- 利率与宏观 ----------
    bond = fetch("国债收益率", lambda: ak.bond_zh_us_rate(start_date="20100101"))
    if bond is not None:
        sheets["国债收益率"] = bond

    macro_map = {
        "宏观_CPI": lambda: ak.macro_china_cpi(),
        "宏观_PPI": lambda: ak.macro_china_ppi(),
        "宏观_M2": lambda: ak.macro_china_money_supply(),
        "宏观_GDP": lambda: ak.macro_china_gdp(),
        "宏观_PMI": lambda: ak.macro_china_pmi(),
        "宏观_社融": lambda: ak.macro_china_shrzgm(),
    }
    for name, fn in macro_map.items():
        df = fetch(name, fn, tries=2, sleep=2)
        if df is not None:
            sheets[name] = df

    # ---------- 衍生: 月度/年度收益 ----------
    close_map = {}
    if hist_hfq is not None:
        h = hist_hfq.copy()
        h["日期"] = to_dt(h["日期"])
        h = h.sort_values("日期").set_index("日期")
        close_map["创元科技"] = h["收盘"]
    if "基准指数日线" in sheets:
        idx_p = sheets["基准指数日线"].copy()
        idx_p["日期"] = to_dt(idx_p["日期"])
        for code, name in index_map.items():
            sub = idx_p[idx_p["指数代码"] == code].set_index("日期")["收盘"].sort_index()
            close_map[name] = sub
    if sw is not None:
        sw_p = sw.copy()
        sw_p["日期"] = to_dt(sw_p["日期"])
        close_map["申万环保"] = sw_p.set_index("日期")["收盘"].sort_index()

    if close_map:
        monthly = pd.DataFrame(close_map).resample("ME").last().pct_change().dropna(how="all")
        monthly = monthly.reset_index()
        monthly.columns = ["月份"] + list(monthly.columns[1:])
        sheets["月度收益"] = monthly

        yearly = pd.DataFrame(close_map).resample("YE").last().pct_change().dropna(how="all")
        yearly = yearly.reset_index()
        yearly.columns = ["年份"] + list(yearly.columns[1:])
        yearly["年份"] = pd.to_datetime(yearly["年份"]).dt.year
        sheets["年度收益"] = yearly

    # ---------- 衍生: 因子代理(日频) ----------
    if "基准指数日线" in sheets and bond is not None:
        idx_p = sheets["基准指数日线"].copy()
        idx_p["日期"] = to_dt(idx_p["日期"])
        pivot = idx_p.pivot_table(index="日期", columns="指数代码", values="收盘")
        rets = pivot.pct_change()

        bond_p = bond.copy()
        bond_p["日期"] = to_dt(bond_p["日期"])
        bond_p = bond_p.set_index("日期")["中国国债收益率10年"].sort_index()
        rf = bond_p.reindex(rets.index).ffill() / 100.0 / 252.0

        factors = pd.DataFrame(index=rets.index)
        factors["市场因子_MKT"] = rets.get("sh000300") - rf
        factors["规模因子_SMB"] = rets.get("sh000852") - rets.get("sh000300")
        factors["价值因子_HML"] = rets.get("sz399371") - rets.get("sz399370")
        factors["红利因子"] = rets.get("sz399321") - rets.get("sh000300")
        factors["无风险利率_日"] = rf
        if hist_hfq is not None:
            hh = hist_hfq.copy()
            hh["日期"] = to_dt(hh["日期"])
            stock_ret = hh.set_index("日期")["收盘"].pct_change()
            factors["个股日收益"] = stock_ret.reindex(factors.index)
            factors["个股超额收益"] = factors["个股日收益"] - rf
        factors = factors.reset_index().rename(columns={"index": "日期"})
        sheets["因子代理_日频"] = factors

    # ---------- 汇总统计与持有体验 ----------
    summary_rows = []
    hold_rows = []
    if hist_hfq is not None:
        hh = hist_hfq.copy()
        hh["日期"] = to_dt(hh["日期"])
        hh = hh.sort_values("日期").set_index("日期")
        close = hh["收盘"]
        ret = close.pct_change().dropna()
        log_ret = np.log(close / close.shift(1)).dropna()
        years = (close.index[-1] - close.index[0]).days / 365.25
        total_ret = close.iloc[-1] / close.iloc[0] - 1
        cagr = (1 + total_ret) ** (1 / years) - 1
        ann_vol = log_ret.std() * np.sqrt(252)
        wealth = (1 + ret).cumprod()
        drawdown = wealth / wealth.cummax() - 1
        max_dd = drawdown.min()

        bond10 = None
        if bond is not None:
            b = bond.copy()
            b["日期"] = to_dt(b["日期"])
            b = b.sort_values("日期").set_index("日期")
            bond10 = b["中国国债收益率10年"].reindex(close.index).ffill() / 100.0
            rf_daily = bond10 / 252.0
        else:
            rf_daily = pd.Series(0.0, index=ret.index)

        excess = ret - rf_daily
        sharpe = excess.mean() / excess.std() * np.sqrt(252) if excess.std() > 0 else np.nan
        downside = excess[excess < 0].std()
        sortino = excess.mean() / downside * np.sqrt(252) if downside and downside > 0 else np.nan
        calmar = cagr / abs(max_dd) if max_dd != 0 else np.nan

        m_ret = close.resample("ME").last().pct_change().dropna()
        monthly_win = (m_ret > 0).mean()

        summary_rows += [
            ("上市日期", close.index[0].strftime("%Y-%m-%d")),
            ("数据截止日", close.index[-1].strftime("%Y-%m-%d")),
            ("交易年数", round(years, 2)),
            ("累计收益率", round(total_ret, 4)),
            ("年化收益率_CAGR", round(cagr, 4)),
            ("年化波动率", round(ann_vol, 4)),
            ("最大回撤", round(max_dd, 4)),
            ("夏普比率(年化)", round(sharpe, 4) if pd.notna(sharpe) else None),
            ("索提诺比率(年化)", round(sortino, 4) if pd.notna(sortino) else None),
            ("卡玛比率", round(calmar, 4) if pd.notna(calmar) else None),
            ("月度胜率", round(monthly_win, 4)),
            ("Beta_对沪深300", None),
            ("近1年收益", None),
            ("近3年收益", None),
            ("近5年收益", None),
            ("近10年收益", None),
        ]

        if "基准指数日线" in sheets:
            idx_p = sheets["基准指数日线"].copy()
            idx_p["日期"] = to_dt(idx_p["日期"])
            idx300 = idx_p[idx_p["指数代码"] == "sh000300"].set_index("日期")["收盘"].sort_index()
            common = pd.concat([ret, idx300.pct_change()], axis=1, join="inner").dropna()
            common.columns = ["stock", "mkt"]
            if len(common) > 30 and common["mkt"].var() > 0:
                beta = common.cov().iloc[0, 1] / common["mkt"].var()
                summary_rows[summary_rows.index(("Beta_对沪深300", None))] = ("Beta_对沪深300", round(beta, 4))

        for label, years_back in [("近1年收益", 1), ("近3年收益", 3), ("近5年收益", 5), ("近10年收益", 10)]:
            target = close.index[-1] - pd.DateOffset(years=years_back)
            anchor = close[close.index <= target]
            if len(anchor) > 0:
                val = close.iloc[-1] / anchor.iloc[-1] - 1
                summary_rows[summary_rows.index((label, None))] = (label, round(val, 4))

        # 持有体验模拟: 任意月买入持有 N 年
        for years_back in [1, 3, 5, 10]:
            horizon = years_back * 12
            vals = []
            mvals = m_ret.values
            for i in range(len(mvals) - horizon):
                r = np.prod(1 + mvals[i + 1 : i + 1 + horizon]) - 1
                vals.append(r)
            if vals:
                arr = np.array(vals)
                ann = (1 + arr) ** (12 / horizon) - 1
                hold_rows.append(
                    {
                        "持有年限": years_back,
                        "样本数": len(arr),
                        "胜率": round((arr > 0).mean(), 4),
                        "平均年化收益": round(ann.mean(), 4),
                        "最差年化收益": round(ann.min(), 4),
                        "最好年化收益": round(ann.max(), 4),
                    }
                )

    if value is not None:
        v = value.copy()
        v["数据日期"] = to_dt(v["数据日期"])
        pe = v["PE(TTM)"].dropna()
        pb = v["市净率"].dropna()
        summary_rows.append(("PE_TTM当前值", round(float(pe.iloc[-1]), 4)))
        summary_rows.append(("PE_TTM历史分位(自2018)", round(float((pe <= pe.iloc[-1]).mean()), 4)))
        summary_rows.append(("PB当前值", round(float(pb.iloc[-1]), 4)))
        summary_rows.append(("PB历史分位(自2018)", round(float((pb <= pb.iloc[-1]).mean()), 4)))

    if fhps is not None and len(fhps) > 0:
        latest_div = fhps.iloc[-1]
        summary_rows.append(("最近报告期", str(latest_div.get("报告期", ""))))
        summary_rows.append(("最近现金分红比例", latest_div.get("现金分红-现金分红比例", None)))
        summary_rows.append(("最近股息率", latest_div.get("现金分红-股息率", None)))

    summary_df = pd.DataFrame(summary_rows, columns=["指标", "数值"])
    summary_df["说明"] = [
        "后复权收盘价计算" if k == "累计收益率" else ""
        for k in summary_df["指标"]
    ]
    sheets["汇总统计"] = summary_df
    if hold_rows:
        sheets["持有体验模拟"] = pd.DataFrame(hold_rows)

    # ---------- 说明 ----------
    notes = [
        ["股票代码", STOCK],
        ["股票名称", "创元科技(以行情接口返回为准)"],
        ["生成时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ["数据来源", "akshare 公开接口(东方财富/新浪/申万宏源/同花顺/交易所)"],
        ["说明1", "日线行情含前复权与后复权, 收益计算请优先使用后复权"],
        ["说明2", "因子为指数代理: MKT=沪深300超额, SMB=中证1000-沪深300, HML=国证价值-国证成长, 红利=国证红利-沪深300"],
        ["说明3", "北向个股持股自2024-08起交易所改为季度披露, 仅保留最新一期"],
        ["说明4", "融资融券为季度末交易日快照, 非连续日频"],
        ["说明5", "估值历史来自东财价值分析接口, 区间以接口返回为准(约2018年起)"],
        ["说明6", "社融接口若失败会记录在下方失败清单"],
        ["Sheet数", str(len(sheets) + 1)],
        ["失败清单", "; ".join(FAILURES) if FAILURES else "无"],
    ]
    sheets["说明"] = pd.DataFrame(notes, columns=["字段", "值"])

    # ---------- 写 Excel ----------
    order = ["说明"] + [k for k in sheets if k != "说明"]
    with pd.ExcelWriter(OUT_FILE, engine="openpyxl") as writer:
        for name in order:
            sheets[name].to_excel(writer, sheet_name=name[:31], index=False)

    print(f"\n输出文件: {OUT_FILE}", flush=True)
    print(f"Sheet 列表: {order}", flush=True)
    print(f"失败接口: {FAILURES if FAILURES else '无'}", flush=True)
    print(f"耗时: {time.time() - t0:.1f}s", flush=True)


if __name__ == "__main__":
    main()
