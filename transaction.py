# -*- coding: utf-8 -*-
import os
import sys
import subprocess
import webbrowser
import time
import tempfile
import shutil
import re
import importlib
from datetime import datetime, timedelta
from typing import Any, cast

import requests
import pandas as pd
import matplotlib.pyplot as plt
from dateutil.relativedelta import relativedelta


FUNDAMENTAL_STOCK_NAMES = {
    "3105": "穩懋",
    "3081": "聯亞",
    "2455": "全新",
    "4991": "環宇-KY",
    "4971": "IET-KY",
    "3163": "波若威",
    "4979": "華星光",
    "4977": "眾達-KY",
    "3234": "光環",
    "6442": "光聖",
    "3363": "上詮",
    "6451": "訊芯-KY",
    "3450": "聯鈞",
    "3380": "明泰",
    "3037": "欣興",
    "8046": "南電",
    "3189": "景碩",
    "2367": "燿華",
    "2313": "華通",
    "3305": "昇貿",
    "3491": "昇達科",
    "3138": "耀登",
    "2485": "兆赫",
}

FUNDAMENTAL_GROUPS = {
    "CPO": ["3105", "3081", "2455", "4991", "4971", "3163", "4979", "4977", "3234", "6442", "3363", "6451", "3450", "3380"],
    "ABF": ["3037", "8046", "3189"],
    "高頻通訊": ["3105", "2367", "2313", "3305", "3491", "3081", "3138", "2485"],
}
FUNDAMENTAL_GROUPS["穩懋綜合"] = list(dict.fromkeys(FUNDAMENTAL_GROUPS["CPO"] + FUNDAMENTAL_GROUPS["高頻通訊"]))

FUNDAMENTAL_PROFIT_METRICS = ["ROE(%)", "ROA(%)", "EPS", "毛利率(%)"]
FUNDAMENTAL_SOLVENCY_METRICS = ["流動比率", "速動比率"]
FUNDAMENTAL_METRICS = FUNDAMENTAL_PROFIT_METRICS + FUNDAMENTAL_SOLVENCY_METRICS
FINMIND_DATA_URL = "https://api.finmindtrade.com/api/v4/data"
STOCK_ID_COL = "股票代號"
STOCK_NAME_COL = "標的"
YEAR_COL = "年度"
GROUP_COL = "族群"
RANK_SUFFIX = "排名"
TREND_SUFFIX = "趨勢"


def _fmt_stock_name(stock_id):
    return f"{FUNDAMENTAL_STOCK_NAMES.get(str(stock_id), str(stock_id))} ({stock_id})"


def _safe_float(value):
    try:
        if value is None or pd.isna(value):
            return None
        return float(value)
    except Exception:
        return None


def _safe_div(numerator, denominator, multiplier=1.0):
    numerator = _safe_float(numerator)
    denominator = _safe_float(denominator)
    if numerator is None or denominator in (None, 0):
        return None
    return numerator / denominator * multiplier


def _round_or_none(value, digits=2):
    value = _safe_float(value)
    if value is None:
        return None
    return round(value, digits)


def _json_safe(value):
    if isinstance(value, dict):
        return {k: _json_safe(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_json_safe(v) for v in value]
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    return value


def _normalize_field_name(value):
    return re.sub(r"\s+", "", str(value or "")).lower()


def _fetch_finmind_dataset(dataset, stock_ids, start_date, end_date, token, timeout=30):
    frames = []
    headers = {"Authorization": f"Bearer {token}"} if token else {}
    for stock_id in stock_ids:
        try:
            resp = requests.get(
                FINMIND_DATA_URL,
                headers=headers,
                params={
                    "dataset": dataset,
                    "data_id": stock_id,
                    "start_date": start_date,
                    "end_date": end_date,
                },
                timeout=timeout,
            )
            resp.raise_for_status()
            payload = resp.json()
            rows = payload.get("data", [])
            if rows:
                frames.append(pd.DataFrame(rows))
        except Exception:
            continue
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def _field_value(rows, include_terms, exclude_terms=None):
    if rows is None or rows.empty:
        return None
    exclude_terms = exclude_terms or []
    best = None
    for _, row in rows.iterrows():
        name = _normalize_field_name(f"{row.get('type', '')} {row.get('origin_name', '')}")
        if any(term in name for term in exclude_terms):
            continue
        if all(term in name for term in include_terms):
            value = _safe_float(row.get("value"))
            if value is not None:
                best = value
                break
    return best


def _first_field_value(rows, candidates):
    for include_terms, exclude_terms in candidates:
        value = _field_value(rows, include_terms, exclude_terms)
        if value is not None:
            return value
    return None


def _statement_value(rows, field):
    candidates = {
        "revenue": [
            (["營業收入"], ["成本"]),
            (["營收"], ["成本"]),
            (["revenue"], ["cost"]),
        ],
        "gross_profit": [
            (["營業毛利"], []),
            (["毛利"], []),
            (["grossprofit"], []),
        ],
        "net_income": [
            (["本期淨利"], []),
            (["稅後淨利"], []),
            (["淨利"], ["毛利"]),
            (["profitloss"], []),
        ],
        "eps": [
            (["基本每股盈餘"], []),
            (["每股盈餘"], []),
            (["basiceps"], []),
            (["eps"], []),
        ],
    }
    return _first_field_value(rows, candidates.get(field, []))


def _balance_value(rows, field):
    candidates = {
        "assets": [
            (["資產總額"], []),
            (["資產合計"], []),
            (["assets"], ["current"]),
        ],
        "equity": [
            (["權益總額"], []),
            (["權益合計"], []),
            (["equity"], []),
        ],
        "current_assets": [
            (["流動資產合計"], []),
            (["流動資產"], ["非流動"]),
            (["currentassets"], []),
        ],
        "current_liabilities": [
            (["流動負債合計"], []),
            (["流動負債"], ["非流動"]),
            (["currentliabilities"], []),
        ],
        "inventory": [
            (["存貨"], []),
            (["inventor"], []),
        ],
        "prepayment": [
            (["預付款"], []),
            (["prepayment"], []),
            (["prepaid"], []),
        ],
    }
    return _first_field_value(rows, candidates.get(field, []))


def _prepare_financial_source(financial_df, balance_df):
    for df in (financial_df, balance_df):
        if df is not None and df.empty and "year" not in df.columns:
            df["year"] = pd.Series(dtype="int64")
        if df is not None and not df.empty and "date" in df.columns:
            df["date"] = pd.to_datetime(df["date"], errors="coerce")
            df["year"] = df["date"].dt.year
            df["value"] = pd.to_numeric(df["value"], errors="coerce")
    return financial_df, balance_df


def _build_yearly_fundamentals(financial_df, balance_df, stock_ids, today=None):
    today = today or datetime.today()
    current_year = today.year
    rows = []
    financial_df, balance_df = _prepare_financial_source(financial_df.copy(), balance_df.copy())

    for stock_id in stock_ids:
        fin_stock = financial_df[financial_df["stock_id"].astype(str) == str(stock_id)] if not financial_df.empty else pd.DataFrame()
        bal_stock = balance_df[balance_df["stock_id"].astype(str) == str(stock_id)] if not balance_df.empty else pd.DataFrame()
        fin_years = (
            set(fin_stock.loc[fin_stock["year"] < current_year, "year"].dropna().astype(int))
            if not fin_stock.empty and "year" in fin_stock.columns else set()
        )
        bal_years = (
            set(bal_stock.loc[bal_stock["year"] < current_year, "year"].dropna().astype(int))
            if not bal_stock.empty and "year" in bal_stock.columns else set()
        )
        available_years = sorted(fin_years | bal_years)
        target_years = available_years[-4:]
        prev_assets = None
        prev_equity = None

        for year in target_years:
            fin_year = fin_stock[fin_stock["year"] == year]
            bal_year = bal_stock[bal_stock["year"] == year]
            if not fin_year.empty:
                fin_date = fin_year["date"].max()
                fin_rows = fin_year[fin_year["date"] == fin_date]
            else:
                fin_rows = pd.DataFrame()
            if not bal_year.empty:
                bal_date = bal_year["date"].max()
                bal_rows = bal_year[bal_year["date"] == bal_date]
            else:
                bal_rows = pd.DataFrame()

            revenue = _statement_value(fin_rows, "revenue")
            gross_profit = _statement_value(fin_rows, "gross_profit")
            net_income = _statement_value(fin_rows, "net_income")
            eps = _statement_value(fin_rows, "eps")
            assets = _balance_value(bal_rows, "assets")
            equity = _balance_value(bal_rows, "equity")
            current_assets = _balance_value(bal_rows, "current_assets")
            current_liabilities = _balance_value(bal_rows, "current_liabilities")
            inventory = _balance_value(bal_rows, "inventory") or 0
            prepayment = _balance_value(bal_rows, "prepayment") or 0

            avg_assets = None
            if assets is not None:
                avg_assets = (assets + prev_assets) / 2 if prev_assets is not None else assets
            avg_equity = None
            if equity is not None:
                avg_equity = (equity + prev_equity) / 2 if prev_equity is not None else equity

            rows.append({
                STOCK_ID_COL: str(stock_id),
                STOCK_NAME_COL: _fmt_stock_name(stock_id),
                YEAR_COL: int(year),
                "ROE(%)": _round_or_none(_safe_div(net_income, avg_equity, 100)),
                "ROA(%)": _round_or_none(_safe_div(net_income, avg_assets, 100)),
                "EPS": _round_or_none(eps),
                "毛利率(%)": _round_or_none(_safe_div(gross_profit, revenue, 100)),
                "流動比率": _round_or_none(_safe_div(current_assets, current_liabilities)),
                "速動比率": _round_or_none(_safe_div((current_assets - inventory - prepayment) if current_assets is not None else None, current_liabilities)),
            })

            if assets is not None:
                prev_assets = assets
            if equity is not None:
                prev_equity = equity

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    return df.sort_values([STOCK_ID_COL, YEAR_COL]).groupby(STOCK_ID_COL, group_keys=False).tail(3).reset_index(drop=True)


def _rank_group_latest(yearly_df, group_name, stock_ids):
    if yearly_df is None or yearly_df.empty:
        return []
    latest_rows = []
    for stock_id in stock_ids:
        one = yearly_df[yearly_df[STOCK_ID_COL] == str(stock_id)].sort_values(YEAR_COL)
        if not one.empty:
            latest_rows.append(one.iloc[-1].to_dict())
        else:
            latest_rows.append({STOCK_ID_COL: str(stock_id), STOCK_NAME_COL: _fmt_stock_name(stock_id), YEAR_COL: None})
    df = pd.DataFrame(latest_rows)
    df.insert(0, GROUP_COL, group_name)
    for metric in FUNDAMENTAL_METRICS:
        if metric not in df.columns:
            df[metric] = None
        ranks = pd.to_numeric(df[metric], errors="coerce").rank(ascending=False, method="min")
        df[f"{metric}{RANK_SUFFIX}"] = [
            int(v) if pd.notna(v) else None
            for v in ranks
        ]
    return df.to_dict("records")


def _trend_label(values):
    clean = [_safe_float(v) for v in values]
    clean = [v for v in clean if v is not None]
    if len(clean) < 2:
        return "無資料"
    if clean[-1] > clean[0]:
        return "成長"
    if clean[-1] < clean[0]:
        return "衰退"
    return "持平"


def _build_trend_rows(yearly_df, stock_ids):
    rows = []
    for stock_id in stock_ids:
        one = yearly_df[yearly_df[STOCK_ID_COL] == str(stock_id)].sort_values(YEAR_COL) if yearly_df is not None and not yearly_df.empty else pd.DataFrame()
        row = {STOCK_NAME_COL: _fmt_stock_name(stock_id)}
        if one.empty:
            row.update({YEAR_COL: "無資料"})
            for metric in FUNDAMENTAL_PROFIT_METRICS:
                row[f"{metric}{TREND_SUFFIX}"] = "無資料"
            rows.append(row)
            continue
        row[YEAR_COL] = " / ".join(str(int(y)) for y in one[YEAR_COL].tolist())
        for metric in FUNDAMENTAL_PROFIT_METRICS:
            vals = [_round_or_none(v) for v in one.get(metric, pd.Series(dtype=float)).tolist()]
            row[metric] = " / ".join("無資料" if v is None else str(v) for v in vals)
            row[f"{metric}{TREND_SUFFIX}"] = _trend_label(vals)
        rows.append(row)
    return rows


def _fundamental_score(group_rows, target_id, trend_rows=None):
    df = pd.DataFrame(group_rows)
    if df.empty:
        return 0, "觀望", "無族群資料"
    target = df[df[STOCK_ID_COL].astype(str) == str(target_id)]
    if target.empty:
        return 0, "觀望", "無標的資料"
    target = target.iloc[0]
    valid_count = max(1, len(df))
    top_cut = max(1, int(valid_count / 3 + 0.999))
    bottom_cut = max(1, int(valid_count / 3 + 0.999))
    score = 0
    reasons = []

    for metric in FUNDAMENTAL_PROFIT_METRICS:
        rank = target.get(f"{metric}{RANK_SUFFIX}")
        if rank is None or pd.isna(rank):
            continue
        if rank <= top_cut:
            score += 1
            reasons.append(f"{metric}排名前段")
        elif rank > valid_count - bottom_cut:
            score -= 1
            reasons.append(f"{metric}排名後段")

    for metric in FUNDAMENTAL_SOLVENCY_METRICS:
        val = _safe_float(target.get(metric))
        vals = pd.to_numeric(df[metric], errors="coerce").dropna()
        if val is None or vals.empty:
            continue
        median = float(vals.median())
        rank = target.get(f"{metric}{RANK_SUFFIX}")
        if val >= median:
            score += 1
            reasons.append(f"{metric}高於族群中位數")
        elif rank is not None and not pd.isna(rank) and rank > valid_count - bottom_cut:
            score -= 1
            reasons.append(f"{metric}排名後段")

    trend_bonus = 0
    if trend_rows:
        target_name = _fmt_stock_name(target_id)
        trend = next((r for r in trend_rows if r.get(STOCK_NAME_COL) == target_name), None)
        if trend:
            trend_states = [trend.get(f"{m}{TREND_SUFFIX}") for m in FUNDAMENTAL_PROFIT_METRICS]
            growth = trend_states.count("成長")
            decline = trend_states.count("衰退")
            if growth >= 3:
                trend_bonus = 1
                reasons.append("近年獲利趨勢多數成長")
            elif decline >= 3:
                trend_bonus = -1
                reasons.append("近年獲利趨勢多數衰退")
    score += trend_bonus

    if score >= 3:
        suggestion = "買"
    elif score <= -2:
        suggestion = "不買"
    else:
        suggestion = "觀望"
    return score, suggestion, "；".join(reasons) if reasons else "分數中性，暫無明確基本面訊號"


def build_fundamental_payload(token, today=None):
    today = today or datetime.today()
    all_ids = list(dict.fromkeys(sum(FUNDAMENTAL_GROUPS.values(), [])))
    start_date = f"{today.year - 5}-01-01"
    end_date = today.strftime("%Y-%m-%d")
    financial_df = _fetch_finmind_dataset("TaiwanStockFinancialStatements", all_ids, start_date, end_date, token)
    balance_df = _fetch_finmind_dataset("TaiwanStockBalanceSheet", all_ids, start_date, end_date, token)
    yearly_df = _build_yearly_fundamentals(financial_df, balance_df, all_ids, today=today)

    group_tables = {
        group_name: _rank_group_latest(yearly_df, group_name, stock_ids)
        for group_name, stock_ids in FUNDAMENTAL_GROUPS.items()
    }
    trend_rows = _build_trend_rows(yearly_df, ["3105", "3037"])

    recommendation_rows = []
    for group_name, target_id in [("CPO", "3105"), ("高頻通訊", "3105"), ("穩懋綜合", "3105"), ("ABF", "3037")]:
        score, suggestion, reason = _fundamental_score(group_tables.get(group_name, []), target_id, trend_rows)
        recommendation_rows.append({
            "標的": _fmt_stock_name(target_id),
            "比較族群": group_name,
            "基本面分數": score,
            "基本面建議": suggestion,
            "說明": reason,
        })

    summary_rows = []
    for group_name, target_id in [("CPO", "3105"), ("高頻通訊", "3105"), ("穩懋綜合", "3105"), ("ABF", "3037")]:
        rows = group_tables.get(group_name, [])
        target = next((r for r in rows if str(r.get(STOCK_ID_COL)) == target_id), None)
        if not target:
            continue
        summary_rows.append({
            STOCK_NAME_COL: _fmt_stock_name(target_id),
            "比較族群": group_name,
            YEAR_COL: target.get(YEAR_COL),
            "ROE排名": target.get("ROE(%)排名"),
            "ROA排名": target.get("ROA(%)排名"),
            "EPS排名": target.get("EPS排名"),
            "毛利率排名": target.get("毛利率(%)排名"),
            "流動比率排名": target.get("流動比率排名"),
            "速動比率排名": target.get("速動比率排名"),
        })

    return _json_safe({
        "summary_rows": summary_rows,
        "recommendation_rows": recommendation_rows,
        "trend_rows": trend_rows,
        "group_tables": group_tables,
    })


# --- Streamlit 儀表板 ---
def run_streamlit_app():
    import streamlit as st
    try:
        DataLoader = importlib.import_module("FinMind.data").DataLoader
        webdriver = importlib.import_module("selenium.webdriver")
        By = importlib.import_module("selenium.webdriver.common.by").By
        Keys = importlib.import_module("selenium.webdriver.common.keys").Keys
        WebDriverWait = importlib.import_module("selenium.webdriver.support.ui").WebDriverWait
        EC = importlib.import_module("selenium.webdriver.support.expected_conditions")
    except ModuleNotFoundError as exc:
        missing = getattr(exc, "name", "unknown")
        msg = f"缺少套件：{missing}，請先安裝所需套件後再執行。"
        print(msg)
        st.error(msg)
        st.info("建議安裝：pip install FinMind selenium flask streamlit")
        return

    st.set_page_config(
        page_title="5475 / 3234 / 3105 / 3037 / 0050 / TAIEX 股票交易儀表板",
        layout="wide"
    )
    st.title("5475 / 3234 / 3105 / 3037 / 0050 與加權指數 (TAIEX) 股票交易儀表板")
    st.caption("包含股價走勢、累計報酬、相對大盤超額績效、投資組合、法人籌碼、基本面與融資融券借券分析。")

    # =============================
    # 基本參數
    # =============================
    FINMIND_TOKEN = os.getenv(
        "FINMIND_TOKEN",
        "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJkYXRlIjoiMjAyNi0wMy0xOCAyMjozNToxNyIsInVzZXJfaWQiOiI4OTEwMDciLCJlbWFpbCI6ImxpbGxhcmQ4MDZAZ21haWwuY29tIiwiaXAiOiIxMTQuMzYuMTgxLjcxIn0.WEgK_gYl-WQfRxxMR9bVrvkMsltT1pHfO_TEmvvIsMU"
    )

    WANTGOO_USERNAME = os.getenv("WANTGOO_USERNAME", "lillard1006@gmail.com")
    WANTGOO_PASSWORD = os.getenv("WANTGOO_PASSWORD", "lillard@80613")
    WANTGOO_HEADLESS = os.getenv("WANTGOO_HEADLESS", "false").lower() == "true"

    FINMIND_BASE_URL = "https://api.finmindtrade.com/api/v4/data"

    TARGET_STOCK_IDS = ["0050", "3234", "5475", "3105", "3037"]

    stock_name_map_local = {
        "0050": "0050",
        "3234": "光環 (3234)",
        "5475": "德宏 (5475)",
        "3105": "穩懋 (3105)",
        "3037": "欣興 (3037)"
    }

    # =============================
    # FinMind 登入
    # =============================
    @st.cache_resource
    def get_loader():
        api = DataLoader()
        try:
            login_by_token = getattr(api, "login_by_token")
            login_by_token(api_token=FINMIND_TOKEN)
        except AttributeError:
            login = getattr(api, "login")
            login(token=FINMIND_TOKEN)
        return api

    @st.cache_data(ttl=3600)
    def load_market_data(start_date_main, start_date_0050, end_date):
        try:
            api = get_loader()

            stock_config = {
                "5475": start_date_main,
                "3234": start_date_main,
                "3105": start_date_main,
                "3037": start_date_main,
                "0050": start_date_0050,
                "TAIEX": start_date_main,
            }

            raw_data = {}
            for sid, sdate in stock_config.items():
                df = api.taiwan_stock_daily(
                    stock_id=sid,
                    start_date=sdate,
                    end_date=end_date
                )
                raw_data[sid] = df

            if any(df is None or df.empty for df in raw_data.values()):
                return None, "部分標的無法取得市場資料"

            def process_df(df):
                df = df.rename(columns={
                    "date": "Date",
                    "open": "Open",
                    "max": "High",
                    "min": "Low",
                    "close": "Close",
                    "Trading_Volume": "Volume",
                    "volume": "Volume"
                })

                df["Date"] = pd.to_datetime(df["Date"])
                df = df.sort_values("Date")
                df.set_index("Date", inplace=True)

                for col in ["Open", "High", "Low", "Close", "Volume"]:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors="coerce")

                return df

            processed = {sid: process_df(df) for sid, df in raw_data.items()}

            common_index = None
            for sid, df in processed.items():
                if common_index is None:
                    common_index = df.index
                else:
                    common_index = common_index.intersection(df.index)

            assert common_index is not None
            aligned = {sid: df.loc[common_index].copy() for sid, df in processed.items()}
            return aligned, None

        except Exception as e:
            return None, str(e)

    # =============================
    # 0050 定期定額模擬
    # =============================
    @st.cache_data(ttl=3600)
    def simulate_0050_dca(df_0050, sim_start_date_str):
        if df_0050 is None or df_0050.empty:
            return pd.DataFrame(), 0.0, 0.0

        df = df_0050.copy().sort_index()
        df["MA60"] = df["Close"].rolling(window=60, min_periods=20).mean()

        trading_dates = df.index.sort_values()

        sim_start_date = pd.to_datetime(sim_start_date_str).normalize()
        last_available_date = trading_dates.max().normalize()

        if sim_start_date > last_available_date:
            return pd.DataFrame(), 0.0, 0.0

        month_cursor = pd.Timestamp(sim_start_date.year, sim_start_date.month, 1)
        month_end = pd.Timestamp(last_available_date.year, last_available_date.month, 1)

        scheduled_days = [1, 7, 13, 19, 25]
        records = []
        used_trade_dates = set()

        while month_cursor <= month_end:
            year = month_cursor.year
            month = month_cursor.month

            for day in scheduled_days:
                try:
                    planned_date = pd.Timestamp(year=year, month=month, day=day)
                except ValueError:
                    continue

                if planned_date < sim_start_date or planned_date > last_available_date:
                    continue

                candidate_dates = trading_dates[trading_dates >= planned_date]
                if len(candidate_dates) == 0:
                    continue

                trade_date = candidate_dates[0]
                if trade_date in used_trade_dates:
                    continue

                loc = int(df.index.get_indexer([trade_date])[0])
                if loc <= 0:
                    continue

                prev_trade_date = df.index[loc - 1]
                prev_close = float(df.loc[prev_trade_date, "Close"])
                quarter_avg = float(df.loc[prev_trade_date, "MA60"])

                if pd.isna(quarter_avg) or quarter_avg == 0:
                    invest_amount = 1000
                    diff_pct = None
                else:
                    diff_pct = (prev_close / quarter_avg - 1) * 100
                    if diff_pct <= -5:
                        invest_amount = 2000
                    elif diff_pct >= 5:
                        invest_amount = 800
                    else:
                        invest_amount = 1000

                buy_price = float(df.loc[trade_date, "Close"])
                buy_shares = invest_amount / buy_price if buy_price != 0 else 0

                records.append({
                    "預計買進日": planned_date,
                    "實際買進日": trade_date,
                    "前一交易日": prev_trade_date,
                    "前一日收盤價": prev_close,
                    "前一日MA60": float(quarter_avg) if pd.notna(quarter_avg) else None,
                    "乖離率(%)": diff_pct,
                    "投入金額": invest_amount,
                    "買進價格": buy_price,
                    "買進股數": buy_shares,
                })

                used_trade_dates.add(trade_date)

            month_cursor = month_cursor + relativedelta(months=1)

        dca_df = pd.DataFrame(records)
        if dca_df.empty:
            return dca_df, 0.0, 0.0

        total_dca_cost = float(dca_df["投入金額"].sum())
        total_dca_shares = float(dca_df["買進股數"].sum())
        return dca_df, total_dca_cost, total_dca_shares

    # =============================
    # 投資組合
    # =============================
    @st.cache_data(ttl=3600)
    def build_portfolio_df(df_5475, df_3234, df_0050, dca_df=None):
        if dca_df is None:
            dca_df = pd.DataFrame()

        base_portfolio_data = [
            {
                "標的": "德宏 (5475)",
                "代號": "5475",
                "平均成本": 108.5,
                "原始投入成本": 108543,
                "現價": float(df_5475["Close"].iloc[-1]),
            },
            {
                "標的": "光環 (3234)",
                "代號": "3234",
                "平均成本": 86.47,
                "原始投入成本": 129751,
                "現價": float(df_3234["Close"].iloc[-1]),
            },
            {
                "標的": "0050",
                "代號": "0050",
                "平均成本": 53.92,
                "原始投入成本": 48054,
                "現價": float(df_0050["Close"].iloc[-1]),
            },
        ]

        df_portfolio = pd.DataFrame(base_portfolio_data)
        df_portfolio["原始股數"] = df_portfolio["原始投入成本"] / df_portfolio["平均成本"]

        dca_cost_0050 = 0.0
        dca_shares_0050 = 0.0
        if not dca_df.empty:
            dca_cost_0050 = float(dca_df["投入金額"].sum())
            dca_shares_0050 = float(dca_df["買進股數"].sum())

        df_portfolio["定期定額成本"] = 0.0
        df_portfolio["定期定額股數"] = 0.0

        mask_0050 = df_portfolio["代號"] == "0050"
        df_portfolio.loc[mask_0050, "定期定額成本"] = dca_cost_0050
        df_portfolio.loc[mask_0050, "定期定額股數"] = dca_shares_0050

        df_portfolio["總成本"] = df_portfolio["原始投入成本"] + df_portfolio["定期定額成本"]
        df_portfolio["總股數"] = df_portfolio["原始股數"] + df_portfolio["定期定額股數"]

        df_portfolio["市值"] = df_portfolio["總股數"] * df_portfolio["現價"]
        df_portfolio["損益"] = df_portfolio["市值"] - df_portfolio["總成本"]
        df_portfolio["報酬率(%)"] = (df_portfolio["損益"] / df_portfolio["總成本"]) * 100

        total_cost = float(df_portfolio["總成本"].sum())
        total_value = float(df_portfolio["市值"].sum())
        total_profit = float(df_portfolio["損益"].sum())
        total_return = (total_profit / total_cost) * 100 if total_cost != 0 else 0.0

        return df_portfolio, total_cost, total_value, total_profit, total_return

    # =============================
    # 法人籌碼資料
    # =============================
    @st.cache_data(ttl=3600)
    def load_institutional_data(stock_id, start_date, end_date, is_market_total=False):
        api = get_loader()

        if is_market_total:
            df = api.taiwan_stock_institutional_investors_total(
                start_date=start_date,
                end_date=end_date
            )
        else:
            df = api.taiwan_stock_institutional_investors(
                stock_id=stock_id,
                start_date=start_date,
                end_date=end_date
            )

        if df is None or df.empty:
            return pd.DataFrame()

        df = df.copy()
        df["date"] = pd.to_datetime(df["date"])
        df["buy"] = pd.to_numeric(df["buy"], errors="coerce").fillna(0)
        df["sell"] = pd.to_numeric(df["sell"], errors="coerce").fillna(0)
        df["net"] = df["buy"] - df["sell"]

        return df.sort_values("date")

    def summarize_institutional_table(raw_df, target_name, last_n_days=10):
        if raw_df is None or raw_df.empty:
            return pd.DataFrame()

        df = raw_df.copy()
        grouped = df.groupby(["date", "name"], as_index=False).agg({"net": "sum"})
        pivot_df = grouped.pivot(index="date", columns="name", values="net").fillna(0)

        def get_series(col_name):
            if col_name in pivot_df.columns:
                return pivot_df[col_name]
            return pd.Series(0, index=pivot_df.index, dtype="float64")

        foreign = get_series("Foreign_Investor") + get_series("Foreign_Dealer_Self")
        investment_trust = get_series("Investment_Trust")
        dealer_total = get_series("Dealer_self") + get_series("Dealer_Hedging")
        total_3 = foreign + investment_trust + dealer_total

        result = pd.DataFrame({
            "日期": pivot_df.index,
            "外資買賣超": foreign.values,
            "投信買賣超": investment_trust.values,
            "自營商買賣超": dealer_total.values,
            "三大法人買賣超": total_3.values,
        }).sort_values("日期")

        result = result.tail(last_n_days).copy()
        result["標的"] = target_name

        cols = ["標的", "日期", "外資買賣超", "投信買賣超", "自營商買賣超", "三大法人買賣超"]
        return result[cols]

    def build_recommendation(one_target_df, target_name):
        if one_target_df is None or one_target_df.empty:
            return {
                "標的": target_name,
                "近三天外資買賣超": None,
                "近三天投信買賣超": None,
                "近三天自營商買賣超": None,
                "近三天三大法人買賣超": None,
                "法人建議": "無資料",
                "法人說明": "無法人資料",
            }

        recent3 = one_target_df.sort_values("日期").tail(3)

        foreign_3d = float(recent3["外資買賣超"].sum())
        trust_3d = float(recent3["投信買賣超"].sum())
        dealer_3d = float(recent3["自營商買賣超"].sum())
        total3_3d = float(recent3["三大法人買賣超"].sum())

        if trust_3d > 0:
            suggestion = "不買"
            reason = "投信近三天買超，依原規則判斷為不買"
        elif (foreign_3d < 0) and (total3_3d < 0):
            suggestion = "不買"
            reason = "外資與三大法人近三天同步賣超"
        elif (foreign_3d > 0) or (total3_3d > 0):
            suggestion = "買"
            reason = "外資或三大法人近三天買超"
        else:
            suggestion = "觀望"
            reason = "法人籌碼沒有明確買賣訊號"

        return {
            "標的": target_name,
            "近三天外資買賣超": foreign_3d,
            "近三天投信買賣超": trust_3d,
            "近三天自營商買賣超": dealer_3d,
            "近三天三大法人買賣超": total3_3d,
            "法人建議": suggestion,
            "法人說明": reason,
        }

    def style_net_table(df):
        def color_net(val):
            try:
                val = float(val)
                if val > 0:
                    return "color: red; font-weight: bold;"
                elif val < 0:
                    return "color: green; font-weight: bold;"
                return ""
            except Exception:
                return ""

        numeric_cols = [
            "外資買賣超", "投信買賣超", "自營商買賣超", "三大法人買賣超",
            "近三天外資買賣超", "近三天投信買賣超",
            "近三天自營商買賣超", "近三天三大法人買賣超",
        ]
        existing_numeric_cols = [c for c in numeric_cols if c in df.columns]

        styler = df.style
        if existing_numeric_cols:
            styler = styler.format({col: "{:,.0f}" for col in existing_numeric_cols})
            styler = styler.map(color_net, subset=existing_numeric_cols)

        return styler

    # =============================
    # 融資 / 融券 / 借券
    # =============================
    def fetch_finmind_data(dataset, data_id, start_date, end_date, token):
        headers = {
            "Authorization": f"Bearer {token}"
        }
        params = {
            "dataset": dataset,
            "data_id": data_id,
            "start_date": start_date,
            "end_date": end_date,
        }

        r = requests.get(FINMIND_BASE_URL, headers=headers, params=params, timeout=30)
        r.raise_for_status()
        data = r.json()

        if "data" not in data:
            raise ValueError(f"{dataset} 回傳資料缺少 data 欄位")

        return pd.DataFrame(data["data"])

    @st.cache_data(ttl=3600)
    def load_margin_short_lending_data(stock_id, start_date, end_date, token):
        try:
            df_margin = fetch_finmind_data(
                dataset="TaiwanStockMarginPurchaseShortSale",
                data_id=stock_id,
                start_date=start_date,
                end_date=end_date,
                token=token
            )

            if df_margin.empty:
                return pd.DataFrame(), f"{stock_id} 融資融券資料為空"

            df_margin["date"] = pd.to_datetime(df_margin["date"])

            margin_cols = [
                "date",
                "stock_id",
                "MarginPurchaseTodayBalance",
                "ShortSaleTodayBalance"
            ]

            for col in margin_cols:
                if col not in df_margin.columns:
                    raise KeyError(f"{stock_id} 融資融券資料缺少欄位：{col}")

            df_margin = df_margin[margin_cols].copy()

            df_lending = fetch_finmind_data(
                dataset="TaiwanStockSecuritiesLending",
                data_id=stock_id,
                start_date=start_date,
                end_date=end_date,
                token=token
            )

            if df_lending.empty:
                df_lending_daily = pd.DataFrame(columns=["date", "SecuritiesLendingVolume"])
            else:
                df_lending["date"] = pd.to_datetime(df_lending["date"])

                if "volume" not in df_lending.columns:
                    raise KeyError(f"{stock_id} 借券資料缺少欄位：volume")

                df_lending_daily = (
                    df_lending.groupby("date", as_index=False)
                    .agg({"volume": "sum"})
                    .rename(columns={"volume": "SecuritiesLendingVolume"})
                )

            df = pd.merge(
                df_margin,
                df_lending_daily,
                on="date",
                how="left"
            )

            df["SecuritiesLendingVolume"] = df["SecuritiesLendingVolume"].fillna(0)
            df = df.sort_values("date").tail(7).reset_index(drop=True)
            df["date_str"] = df["date"].dt.strftime("%Y-%m-%d")

            df["Prev_Margin"] = df["MarginPurchaseTodayBalance"].shift(1)
            df["Prev_Short"] = df["ShortSaleTodayBalance"].shift(1)
            df["Prev_Lending"] = df["SecuritiesLendingVolume"].shift(1)

            def get_signal(row):
                if pd.isna(row["Prev_Margin"]) or pd.isna(row["Prev_Short"]) or pd.isna(row["Prev_Lending"]):
                    return "無資料"

                if (
                    row["MarginPurchaseTodayBalance"] > row["Prev_Margin"]
                    and row["ShortSaleTodayBalance"] < row["Prev_Short"]
                ):
                    return "建議不買"

                elif row["SecuritiesLendingVolume"] < row["Prev_Lending"]:
                    return "建議買"

                else:
                    return "觀望"

            df["Signal"] = df.apply(get_signal, axis=1)

            plot_df = df.copy()

            if len(plot_df) > 0 and plot_df["MarginPurchaseTodayBalance"].iloc[0] != 0:
                plot_df["Margin_Index"] = (
                    plot_df["MarginPurchaseTodayBalance"] / plot_df["MarginPurchaseTodayBalance"].iloc[0] * 100
                )
            else:
                plot_df["Margin_Index"] = 100

            if len(plot_df) > 0 and plot_df["ShortSaleTodayBalance"].iloc[0] != 0:
                plot_df["Short_Index"] = (
                    plot_df["ShortSaleTodayBalance"] / plot_df["ShortSaleTodayBalance"].iloc[0] * 100
                )
            else:
                plot_df["Short_Index"] = 100

            if len(plot_df) > 0 and plot_df["SecuritiesLendingVolume"].iloc[0] != 0:
                plot_df["Lending_Index"] = (
                    plot_df["SecuritiesLendingVolume"] / plot_df["SecuritiesLendingVolume"].iloc[0] * 100
                )
            else:
                base = plot_df["SecuritiesLendingVolume"].replace(0, pd.NA).dropna()
                if len(base) > 0:
                    first_valid = base.iloc[0]
                    plot_df["Lending_Index"] = plot_df["SecuritiesLendingVolume"] / first_valid * 100
                else:
                    plot_df["Lending_Index"] = 100

            return plot_df, None

        except Exception as e:
            return pd.DataFrame(), str(e)

    @st.cache_data(ttl=3600)
    def load_all_margin_short_lending(stock_ids, start_date, end_date, token):
        result = {}
        for sid in stock_ids:
            df, err = load_margin_short_lending_data(sid, start_date, end_date, token)
            result[sid] = {
                "df": df,
                "error": err
            }
        return result

    def build_margin_signal_summary(margin_data_map):
        rows = []
        for sid in TARGET_STOCK_IDS:
            item = margin_data_map.get(sid, {})
            df = item.get("df", pd.DataFrame())
            err = item.get("error")

            if err:
                rows.append({
                    "標的": stock_name_map_local.get(sid, sid),
                    "最新日期": "",
                    "最新融資餘額": None,
                    "最新融券餘額": None,
                    "最新借券量": None,
                    "融資融券借券建議": "無資料",
                    "融資融券借券說明": err,
                })
                continue

            if df is None or df.empty:
                rows.append({
                    "標的": stock_name_map_local.get(sid, sid),
                    "最新日期": "",
                    "最新融資餘額": None,
                    "最新融券餘額": None,
                    "最新借券量": None,
                    "融資融券借券建議": "無資料",
                    "融資融券借券說明": "查無資料",
                })
                continue

            latest = df.iloc[-1]
            signal = latest.get("Signal", "無資料")

            if signal == "建議買":
                final_signal = "買"
            elif signal == "建議不買":
                final_signal = "不買"
            elif signal == "觀望":
                final_signal = "觀望"
            else:
                final_signal = "無資料"

            rows.append({
                "標的": stock_name_map_local.get(sid, sid),
                "最新日期": latest.get("date_str", ""),
                "最新融資餘額": latest.get("MarginPurchaseTodayBalance"),
                "最新融券餘額": latest.get("ShortSaleTodayBalance"),
                "最新借券量": latest.get("SecuritiesLendingVolume"),
                "融資融券借券建議": final_signal,
                "融資融券借券說明": signal,
            })

        return pd.DataFrame(rows)

    def prepare_margin_detail_table(df):
        if df is None or df.empty:
            return pd.DataFrame(columns=[
                "date_str",
                "MarginPurchaseTodayBalance",
                "ShortSaleTodayBalance",
                "SecuritiesLendingVolume",
                "Signal"
            ])

        show_df = df.copy()
        show_df = show_df[[
            "date_str",
            "MarginPurchaseTodayBalance",
            "ShortSaleTodayBalance",
            "SecuritiesLendingVolume",
            "Signal"
        ]].rename(columns={
            "date_str": "日期",
            "MarginPurchaseTodayBalance": "融資餘額",
            "ShortSaleTodayBalance": "融券餘額",
            "SecuritiesLendingVolume": "借券量",
            "Signal": "建議",
        })
        return show_df

    def plot_margin_short_lending_chart(plot_df, stock_label):
        if plot_df is None or plot_df.empty:
            return None

        fig, ax1 = plt.subplots(figsize=(14, 5))

        ax1.plot(
            plot_df["date_str"],
            plot_df["Margin_Index"],
            marker="o",
            linewidth=2,
            label="融資餘額指數化"
        )

        ax1.plot(
            plot_df["date_str"],
            plot_df["Short_Index"],
            marker="o",
            linewidth=2,
            label="融券餘額指數化"
        )

        ax2 = ax1.twinx()
        ax2.plot(
            plot_df["date_str"],
            plot_df["Lending_Index"],
            marker="o",
            linewidth=2.2,
            label="借券量指數化"
        )

        ax1.set_xlabel("Date")
        ax1.set_ylabel("")
        ax1.tick_params(axis="y", left=False, labelleft=False)
        ax1.set_yticks([])
        ax1.spines["left"].set_visible(False)
        ax1.spines["top"].set_visible(False)

        ax2.set_ylabel("")
        ax2.tick_params(axis="y", right=False, labelright=False)
        ax2.set_yticks([])
        ax2.spines["right"].set_visible(False)
        ax2.spines["top"].set_visible(False)

        try:
            ymin = min(
                plot_df["Margin_Index"].min(),
                plot_df["Short_Index"].min(),
                plot_df["Lending_Index"].min()
            )
            ymax = max(
                plot_df["Margin_Index"].max(),
                plot_df["Short_Index"].max(),
                plot_df["Lending_Index"].max()
            )
            padding = max((ymax - ymin) * 0.08, 3)
            ax1.set_ylim(ymin - padding, ymax + padding)
            ax2.set_ylim(ax1.get_ylim())
        except Exception:
            pass

        lines_1, labels_1 = ax1.get_legend_handles_labels()
        lines_2, labels_2 = ax2.get_legend_handles_labels()
        ax1.legend(lines_1 + lines_2, labels_1 + labels_2, loc="upper left", ncol=3)

        plt.title(f"{stock_label} 融資 / 融券 / 借券走勢", fontsize=14)
        fig.tight_layout()

        return fig

    # =============================
    # WantGoo / Selenium 工具
    # =============================
    def safe_click(driver, element):
        try:
            element.click()
        except Exception:
            driver.execute_script("arguments[0].click();", element)

    def _get_cell_text(driver, cell, retries=5, pause=0.25):
        for _ in range(retries):
            try:
                text = cell.get_attribute("innerText")
                if text and text.strip():
                    return text.strip()

                text = cell.get_attribute("textContent")
                if text and text.strip():
                    return text.strip()

                text = driver.execute_script(
                    "return arguments[0].innerText || arguments[0].textContent || '';",
                    cell
                )
                if text and str(text).strip():
                    return str(text).strip()
            except Exception:
                pass

            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", cell)
            except Exception:
                pass

            time.sleep(pause)

        return ""

    def _parse_float_from_text(text):
        if text is None:
            return None
        text = str(text).strip()
        if not text:
            return None

        text = text.replace("%", "").replace(",", "").strip()
        m = re.search(r"-?\d+(?:\.\d+)?", text)
        if not m:
            return None
        try:
            return float(m.group())
        except Exception:
            return None

    def build_chrome_driver(headless=False):
        profile_dir = tempfile.mkdtemp(prefix="wantgoo_chrome_")

        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-first-run")
        options.add_argument("--no-default-browser-check")
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--window-size=1920,1080")
        options.add_argument(f"--user-data-dir={profile_dir}")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)

        if headless:
            options.add_argument("--headless=new")

        driver = webdriver.Chrome(options=options)
        driver.set_page_load_timeout(40)

        try:
            driver.execute_cdp_cmd(
                "Page.addScriptToEvaluateOnNewDocument",
                {
                    "source": """
                        Object.defineProperty(navigator, 'webdriver', {
                            get: () => undefined
                        });
                    """
                }
            )
        except Exception:
            pass

        return driver, profile_dir

    def maybe_login_wantgoo(driver, wait, username, password):
        driver.get("https://www.wantgoo.com/")
        time.sleep(3)

        login_selectors = [
            (By.CSS_SELECTOR, "#unregistered-bar a.topbar-nav__a"),
            (By.XPATH, "//a[contains(., '登入')]"),
            (By.XPATH, "//button[contains(., '登入')]"),
        ]

        login_btn = None
        for by, sel in login_selectors:
            try:
                elems = driver.find_elements(by, sel)
                if elems:
                    login_btn = elems[0]
                    break
            except Exception:
                pass

        if login_btn is None:
            return

        try:
            safe_click(driver, login_btn)
            time.sleep(2)
        except Exception:
            return

        user_candidates = [
            (By.CSS_SELECTOR, 'input[c-model="userName"]'),
            (By.CSS_SELECTOR, 'input[type="email"]'),
            (By.XPATH, "//input[contains(@placeholder, 'Email')]"),
            (By.XPATH, "//input[contains(@placeholder, '帳號')]"),
        ]
        pass_candidates = [
            (By.CSS_SELECTOR, 'input[c-model="password"]'),
            (By.CSS_SELECTOR, 'input[type="password"]'),
        ]
        submit_candidates = [
            (By.CSS_SELECTOR, 'button[login=""]'),
            (By.XPATH, "//button[contains(., '登入')]"),
            (By.XPATH, "//button[contains(., 'Sign in')]"),
        ]

        user_input = None
        pass_input = None
        submit_btn = None

        for by, sel in user_candidates:
            try:
                user_input = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((by, sel))
                )
                if user_input:
                    break
            except Exception:
                pass

        for by, sel in pass_candidates:
            try:
                pass_input = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((by, sel))
                )
                if pass_input:
                    break
            except Exception:
                pass

        for by, sel in submit_candidates:
            try:
                submit_btn = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((by, sel))
                )
                if submit_btn:
                    break
            except Exception:
                pass

        if user_input and pass_input and submit_btn:
            user_input.clear()
            user_input.send_keys(username)
            pass_input.clear()
            pass_input.send_keys(password)
            safe_click(driver, submit_btn)
            time.sleep(4)

    # -----------------------------
    # 券商分點
    # -----------------------------
    def open_branch_buysell_page(driver, stock_no):
        target_urls = [
            f"https://www.wantgoo.com/stock/{stock_no}/major-investors/branch-buysell",
            f"https://www.wantgoo.com/stock/{stock_no}/major-investors",
            f"https://www.wantgoo.com/stock/{stock_no}",
        ]

        last_err = None
        for url in target_urls:
            try:
                driver.get(url)
                time.sleep(3)
                if stock_no in driver.current_url:
                    return
            except Exception as e:
                last_err = e

        if last_err:
            raise last_err

    def switch_to_1_day(driver):
        try:
            date_selector = WebDriverWait(driver, 6).until(
                EC.presence_of_element_located((By.ID, "dateSelector"))
            )
            driver.execute_script("""
                arguments[0].value = "1";
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """, date_selector)
            time.sleep(3)
            return
        except Exception:
            pass

        text_candidates = ["近1日", "近 1 日", "1日", "近一天"]
        for txt in text_candidates:
            xpath_list = [
                f"//*[normalize-space(text())='{txt}']",
                f"//button[contains(normalize-space(.), '{txt}')]",
                f"//a[contains(normalize-space(.), '{txt}')]",
                f"//li[contains(normalize-space(.), '{txt}')]",
                f"//span[contains(normalize-space(.), '{txt}')]",
                f"//option[contains(normalize-space(.), '{txt}')]",
            ]
            for xp in xpath_list:
                try:
                    elem = WebDriverWait(driver, 2).until(
                        EC.presence_of_element_located((By.XPATH, xp))
                    )
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
                    time.sleep(0.5)
                    safe_click(driver, elem)
                    time.sleep(3)
                    return
                except Exception:
                    pass

    def get_table_rows(driver):
        candidate_selectors = [
            "tbody.rt tr",
            "table tbody tr",
            ".rt-tbody .rt-tr-group",
            ".rt-table .rt-tr",
        ]

        last_err = None
        for selector in candidate_selectors:
            try:
                rows = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, selector))
                )
                rows = [r for r in rows if r.is_displayed()]
                if len(rows) >= 3:
                    return rows
            except Exception as e:
                last_err = e

        if last_err:
            raise last_err
        raise RuntimeError("找不到券商分點資料列")

    def extract_broker_name(driver, row):
        td_candidates = [
            "td:nth-child(2)",
            "div[role='cell']:nth-child(2)",
            ".rt-td:nth-child(2)"
        ]

        for selector in td_candidates:
            try:
                cell = row.find_element(By.CSS_SELECTOR, selector)
                name = _get_cell_text(driver, cell, retries=6, pause=0.2)
                if name:
                    return name.strip()
            except Exception:
                pass

        try:
            tds = row.find_elements(By.TAG_NAME, "td")
            if len(tds) >= 2:
                return _get_cell_text(driver, tds[1], retries=6, pause=0.2)
        except Exception:
            pass

        return ""

    def analyze_broker_branch(driver, stock_no):
        open_branch_buysell_page(driver, stock_no)
        switch_to_1_day(driver)
        rows = get_table_rows(driver)

        if len(rows) < 3:
            return {
                "stock_id": stock_no,
                "buy_top3": [],
                "sell_top3": [],
                "broker_signal": "無資料",
                "broker_reason": "分點資料不足",
            }

        top3_rows = rows[:3]
        bottom3_rows = rows[-3:]

        buy_top3 = [extract_broker_name(driver, r) for r in top3_rows]
        sell_top3 = [extract_broker_name(driver, r) for r in reversed(bottom3_rows)]

        buy_top3 = [x for x in buy_top3 if x]
        sell_top3 = [x for x in sell_top3 if x]

        buy_bad_list = [
            "凱基台北",
            "摩根大通",
            "美林",
            "港商野村",
            "新加坡商瑞銀",
            "富邦建國",
        ]

        sell_bad_list = [
            "摩根士丹利",
            "美商高盛",
        ]

        buy_good_list = [
            "摩根士丹利",
            "美商高盛",
        ]

        buy_bad_count = 0
        buy_good_found = False
        buy_good_matches = []
        sell_bad_found = False
        sell_bad_matches = []

        for b in buy_top3:
            for bad in buy_bad_list:
                if bad.lower() in b.lower():
                    buy_bad_count += 1
                    break

        for s in sell_top3:
            for bad in sell_bad_list:
                if bad.lower() in s.lower():
                    sell_bad_found = True
                    sell_bad_matches.append(s)
                    break

        for b in buy_top3:
            for good in buy_good_list:
                if good.lower() in b.lower():
                    buy_good_found = True
                    buy_good_matches.append(b)
                    break

        if buy_good_found and not (buy_bad_count >= 2 or sell_bad_found):
            broker_signal = "買"
            broker_reason = f"買方前三名出現偏多券商：{', '.join(buy_good_matches)}"
        elif (buy_bad_count >= 2 or sell_bad_found) and not buy_good_found:
            reasons = []
            if buy_bad_count >= 2:
                reasons.append("買方前三名出現多個偏空分點")
            if sell_bad_found:
                reasons.append(f"賣方前三名出現偏空券商：{', '.join(sell_bad_matches)}")
            broker_signal = "不買"
            broker_reason = "；".join(reasons)
        elif buy_good_found and (buy_bad_count >= 2 or sell_bad_found):
            reasons = [f"偏多訊號：買方前三名出現 {', '.join(buy_good_matches)}"]
            if buy_bad_count >= 2:
                reasons.append("偏空訊號：買方前三名出現多個偏空分點")
            if sell_bad_found:
                reasons.append(f"偏空訊號：賣方前三名出現 {', '.join(sell_bad_matches)}")
            broker_signal = "觀望"
            broker_reason = "；".join(reasons)
        else:
            broker_signal = "觀望"
            broker_reason = "無明確建議買或不買訊號，建議觀察"

        return {
            "stock_id": stock_no,
            "buy_top3": buy_top3,
            "sell_top3": sell_top3,
            "broker_signal": broker_signal,
            "broker_reason": broker_reason
        }

    # -----------------------------
    # 大戶籌碼
    # -----------------------------
    def go_to_stock_page(driver, wait, stock_no):
        target_urls = [
            f"https://www.wantgoo.com/stock/{stock_no}",
            f"https://www.wantgoo.com/stock/{stock_no}/major-investors",
        ]

        for url in target_urls:
            try:
                driver.get(url)
                time.sleep(3)
                if stock_no in driver.current_url:
                    return
            except Exception:
                pass

        search_selectors = [
            (By.CSS_SELECTOR, "input.frm-control.frm-control--sm.typeahead.tt-input"),
            (By.CSS_SELECTOR, "input[type='search']"),
            (By.XPATH, "//input[contains(@placeholder,'搜尋')]"),
        ]
        for by, sel in search_selectors:
            try:
                search_input = wait.until(EC.presence_of_element_located((by, sel)))
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", search_input)
                time.sleep(1)

                search_input.click()
                search_input.send_keys(Keys.CONTROL, "a")
                search_input.send_keys(Keys.DELETE)
                search_input.send_keys(stock_no)
                time.sleep(1)
                search_input.send_keys(Keys.ENTER)
                time.sleep(5)

                if stock_no in driver.current_url:
                    return
            except Exception:
                pass

        driver.get(f"https://www.wantgoo.com/stock/{stock_no}")
        time.sleep(3)

    def open_major_holders_tab(driver, wait):
        candidates = [
            (By.XPATH, "//a[contains(normalize-space(.), '大戶籌碼')]"),
            (By.XPATH, "//button[contains(normalize-space(.), '大戶籌碼')]"),
            (By.XPATH, "//*[contains(normalize-space(.), '大戶籌碼')]"),
            (By.XPATH, "//a[contains(normalize-space(.), '大戶持股')]"),
            (By.XPATH, "//button[contains(normalize-space(.), '大戶持股')]"),
            (By.XPATH, "//*[contains(normalize-space(.), '大戶持股')]"),
        ]

        for by, sel in candidates:
            try:
                elem = wait.until(EC.element_to_be_clickable((by, sel)))
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
                time.sleep(1)
                safe_click(driver, elem)
                time.sleep(3)
                return True
            except Exception:
                pass
        return False

    def get_major_holder_rows(driver, wait):
        selectors = [
            (By.CSS_SELECTOR, "tr[concentration-item]"),
            (By.CSS_SELECTOR, "tbody tr[concentration-item]"),
            (By.CSS_SELECTOR, "table tbody tr"),
        ]

        for by, sel in selectors:
            try:
                rows = wait.until(EC.presence_of_all_elements_located((by, sel)))
                rows = [r for r in rows if r.is_displayed()]
                if len(rows) >= 2:
                    return rows
            except Exception:
                pass
        return []

    def extract_major_holder_info_from_row(row):
        date_text = ""
        rate_text = ""

        try:
            date_text = row.find_element(By.CSS_SELECTOR, "td[c-model='date']").text.strip()
        except Exception:
            pass

        try:
            rate_text = row.find_element(By.CSS_SELECTOR, "td[c-model='rateOfDistribution']").text.strip()
        except Exception:
            pass

        if not date_text or not rate_text:
            try:
                tds = row.find_elements(By.TAG_NAME, "td")
                td_texts = [td.text.strip() for td in tds if td.text and td.text.strip()]

                if not date_text and len(td_texts) >= 1:
                    date_text = td_texts[0]

                if not rate_text:
                    for t in td_texts:
                        if "%" in t or re.search(r"\d+(?:\.\d+)?", t):
                            parsed = _parse_float_from_text(t)
                            if parsed is not None:
                                rate_text = t
                                break
            except Exception:
                pass

        rate_value = _parse_float_from_text(rate_text)
        return date_text, rate_value

    def analyze_major_investors(driver, wait, stock_no):
        try:
            go_to_stock_page(driver, wait, stock_no)

            opened = open_major_holders_tab(driver, wait)
            if not opened:
                return {
                    "stock_id": stock_no,
                    "this_week_date": None,
                    "this_week_rate": None,
                    "last_week_date": None,
                    "last_week_rate": None,
                    "major_signal": "無資料",
                    "major_reason": "無法開啟大戶持股頁籤",
                }

            rows = get_major_holder_rows(driver, wait)
            if len(rows) < 2:
                return {
                    "stock_id": stock_no,
                    "this_week_date": None,
                    "this_week_rate": None,
                    "last_week_date": None,
                    "last_week_rate": None,
                    "major_signal": "無資料",
                    "major_reason": "查無足夠的大戶持股資料",
                }

            this_week_date, this_week_rate = extract_major_holder_info_from_row(rows[0])
            last_week_date, last_week_rate = extract_major_holder_info_from_row(rows[1])

            if this_week_rate is None or last_week_rate is None:
                return {
                    "stock_id": stock_no,
                    "this_week_date": this_week_date,
                    "this_week_rate": this_week_rate,
                    "last_week_date": last_week_date,
                    "last_week_rate": last_week_rate,
                    "major_signal": "無資料",
                    "major_reason": "大戶持股比例無法解析",
                }

            if this_week_rate < last_week_rate:
                major_signal = "不買"
                major_reason = f"本週大戶持股 {this_week_rate:.2f}% 低於上週 {last_week_rate:.2f}%"
            elif this_week_rate > last_week_rate:
                major_signal = "買"
                major_reason = f"本週大戶持股 {this_week_rate:.2f}% 高於上週 {last_week_rate:.2f}%"
            else:
                major_signal = "觀望"
                major_reason = f"本週大戶持股 {this_week_rate:.2f}% 與上週 {last_week_rate:.2f}% 持平"

            return {
                "stock_id": stock_no,
                "this_week_date": this_week_date,
                "this_week_rate": this_week_rate,
                "last_week_date": last_week_date,
                "last_week_rate": last_week_rate,
                "major_signal": major_signal,
                "major_reason": major_reason
            }

        except Exception as e:
            return {
                "stock_id": stock_no,
                "this_week_date": None,
                "this_week_rate": None,
                "last_week_date": None,
                "last_week_rate": None,
                "major_signal": "無資料",
                "major_reason": f"{stock_no} 大戶籌碼讀取失敗：{e}",
            }

    def load_wantgoo_all_signals(username, password, stock_ids, headless=False):
        if not username or not password:
            return {
                sid: {
                    "stock_id": sid,
                    "buy_top3": [],
                    "sell_top3": [],
                    "broker_signal": "無資料",
                    "broker_reason": "未設定 WantGoo 帳號密碼",
                    "this_week_date": None,
                    "this_week_rate": None,
                    "last_week_date": None,
                    "last_week_rate": None,
                    "major_signal": "無資料",
                    "major_reason": "未設定 WantGoo 帳號密碼",
                }
                for sid in stock_ids
            }

        driver = None
        profile_dir = None

        try:
            driver, profile_dir = build_chrome_driver(headless=headless)
            wait = WebDriverWait(driver, 20)

            maybe_login_wantgoo(driver, wait, username, password)

            results = {}
            for sid in stock_ids:
                stock_result = {"stock_id": sid}

                try:
                    broker_result = analyze_broker_branch(driver, sid)
                    stock_result.update(broker_result)
                except Exception as e:
                    stock_result.update({
                        "buy_top3": [],
                        "sell_top3": [],
                        "broker_signal": "無資料",
                        "broker_reason": f"{sid} 券商分點讀取失敗：{e}",
                    })

                try:
                    major_result = analyze_major_investors(driver, wait, sid)
                    stock_result.update(major_result)
                except Exception as e:
                    stock_result.update({
                        "this_week_date": None,
                        "this_week_rate": None,
                        "last_week_date": None,
                        "last_week_rate": None,
                        "major_signal": "無資料",
                        "major_reason": f"{sid} 大戶籌碼讀取失敗：{e}",
                    })

                results[sid] = stock_result

            return results

        except Exception as e:
            return {
                sid: {
                    "stock_id": sid,
                    "buy_top3": [],
                    "sell_top3": [],
                    "broker_signal": "無資料",
                    "broker_reason": f"券商分點讀取失敗：{e}",
                    "this_week_date": None,
                    "this_week_rate": None,
                    "last_week_date": None,
                    "last_week_rate": None,
                    "major_signal": "無資料",
                    "major_reason": f"大戶籌碼讀取失敗：{e}",
                }
                for sid in stock_ids
            }
        finally:
            if driver is not None:
                try:
                    driver.quit()
                except Exception:
                    pass

            if profile_dir:
                try:
                    shutil.rmtree(profile_dir, ignore_errors=True)
                except Exception:
                    pass

    def combine_suggestion(*signals):
        normalized = []
        for x in signals:
            if x is None:
                continue
            x = str(x).strip()
            if x in ["買", "建議買"]:
                normalized.append("買")
            elif x in ["不買", "建議不買"]:
                normalized.append("不買")
            elif x in ["觀望", "無資料"]:
                normalized.append(x)

        if not normalized:
            return "無資料"

        if "不買" in normalized:
            return "不買"
        if "買" in normalized:
            return "買"
        return "觀望"

    def build_broker_top3_table(wantgoo_results):
        rows = []

        for sid in TARGET_STOCK_IDS:
            data = wantgoo_results.get(sid, {})
            buy_top3 = data.get("buy_top3", [])
            sell_top3 = data.get("sell_top3", [])

            row = {
                "標的": stock_name_map_local.get(sid, sid),
                "買方第1名": buy_top3[0] if len(buy_top3) > 0 else "",
                "買方第2名": buy_top3[1] if len(buy_top3) > 1 else "",
                "買方第3名": buy_top3[2] if len(buy_top3) > 2 else "",
                "賣方第1名": sell_top3[0] if len(sell_top3) > 0 else "",
                "賣方第2名": sell_top3[1] if len(sell_top3) > 1 else "",
                "賣方第3名": sell_top3[2] if len(sell_top3) > 2 else "",
                "券商分點建議": data.get("broker_signal", "無資料"),
                "券商分點說明": data.get("broker_reason", ""),
            }
            rows.append(row)

        return pd.DataFrame(rows)

    def build_major_holder_table(wantgoo_results):
        rows = []
        for sid in TARGET_STOCK_IDS:
            data = wantgoo_results.get(sid, {})
            rows.append({
                "標的": stock_name_map_local.get(sid, sid),
                "本週日期": data.get("this_week_date") or "",
                "本週大戶持股比例(%)": data.get("this_week_rate"),
                "上週日期": data.get("last_week_date") or "",
                "上週大戶持股比例(%)": data.get("last_week_rate"),
                "大戶籌碼建議": data.get("major_signal", "無資料"),
                "大戶籌碼說明": data.get("major_reason", ""),
            })
        return pd.DataFrame(rows)

    # =============================
    # 日期區間設定
    # =============================
    today = datetime.today()
    one_year_ago = today - relativedelta(years=1)

    start_date_main = one_year_ago.strftime("%Y-%m-%d")
    end_date = today.strftime("%Y-%m-%d")
    start_date_0050 = "2025-06-18"

    margin_start_date = (today.date() - timedelta(days=20)).strftime("%Y-%m-%d")
    margin_end_date = today.date().strftime("%Y-%m-%d")

    data_map, err = load_market_data(
        start_date_main=start_date_main,
        start_date_0050=start_date_0050,
        end_date=end_date
    )

    if err:
        st.error(f"資料讀取失敗：{err}")
        return

    if not data_map:
        st.error("沒有取得資料，請確認 API Token、網路連線與股票代號設定。")
        return

    df_5475 = data_map["5475"]
    df_3234 = data_map["3234"]
    df_3105 = data_map["3105"]
    df_3037 = data_map["3037"]
    df_0050 = data_map["0050"]
    df_index = data_map["TAIEX"]

    st.caption(
        f"5475 / 3234 / 3105 / 3037 / TAIEX 區間：{start_date_main} ~ {end_date}；"
        f"0050 區間：{start_date_0050} ~ {end_date}。"
    )

    color_5475 = "#d65f4a"
    color_3234 = "#9467bd"
    color_3105 = "#ff7f0e"
    color_3037 = "#17becf"
    color_0050 = "#2ca02c"
    color_index = "#5b9bd5"

    stock_name_map = {
        "5475": "德宏 (5475)",
        "3234": "光環 (3234)",
        "3105": "穩懋 (3105)",
        "3037": "欣興 (3037)",
        "0050": "0050",
        "TAIEX": "加權指數 (TAIEX)"
    }

    stock_color_map = {
        "5475": color_5475,
        "3234": color_3234,
        "3105": color_3105,
        "3037": color_3037,
        "0050": color_0050,
        "TAIEX": color_index
    }

    # =============================
    # 0050 定期定額模擬
    # =============================
    sim_start_date = start_date_0050
    dca_df, dca_total_cost, dca_total_shares = simulate_0050_dca(df_0050, sim_start_date)

    # =============================
    # 投資組合資料準備
    # =============================
    # Prepare fundamental data once; render recommendation and analysis tables in their requested sections.
    with st.spinner("載入基本面族群比較資料..."):
        fundamental_payload = cast(dict[str, Any], build_fundamental_payload(FINMIND_TOKEN, today=today))

    def color_fundamental_action(val: Any) -> str:
        if val == "買":
            return "color: red; font-weight: bold;"
        elif val == "不買":
            return "color: green; font-weight: bold;"
        elif val == "觀望":
            return "color: orange; font-weight: bold;"
        return ""

    fundamental_rec_df = pd.DataFrame(fundamental_payload.get("recommendation_rows", []))
    fundamental_summary_df = pd.DataFrame(fundamental_payload.get("summary_rows", []))
    trend_df = pd.DataFrame(fundamental_payload.get("trend_rows", []))
    group_tables = cast(dict[str, list[dict[str, Any]]], fundamental_payload.get("group_tables", {}))

    st.divider()

    # =============================
    # 圖表 1：股價走勢
    # =============================
    st.subheader("股價與大盤走勢")

    fig, ax1 = plt.subplots(figsize=(15, 6))

    for sid in ["5475", "3234", "3105", "3037", "0050"]:
        ax1.plot(
            data_map[sid].index,
            data_map[sid]["Close"],
            color=stock_color_map[sid],
            linewidth=2,
            label=stock_name_map[sid]
        )

    ax1.set_xlabel("Date")
    ax1.set_ylabel("")
    ax1.tick_params(axis="y", left=False, labelleft=False)
    ax1.set_yticks([])
    ax1.spines["left"].set_visible(False)
    ax1.spines["top"].set_visible(False)

    ax2 = ax1.twinx()
    ax2.plot(
        df_index.index,
        df_index["Close"],
        color=color_index,
        linewidth=2.4,
        label="加權指數 (TAIEX)"
    )
    ax2.set_ylabel("")
    ax2.tick_params(axis="y", right=False, labelright=False)
    ax2.set_yticks([])
    ax2.spines["right"].set_visible(False)
    ax2.spines["top"].set_visible(False)

    lines_1, labels_1 = ax1.get_legend_handles_labels()
    lines_2, labels_2 = ax2.get_legend_handles_labels()
    ax1.legend(lines_1 + lines_2, labels_1 + labels_2, loc="upper left", ncol=2)

    fig.tight_layout()
    st.pyplot(fig)

    st.divider()

    # =============================
    # 圖表 2：累計報酬率比較
    # =============================
    st.subheader("累計報酬率比較")
    st.markdown("以區間起始日為 **0%**，比較各標的與大盤的累計報酬率。")

    ret_map = {}
    for sid in ["5475", "3234", "3105", "3037", "0050", "TAIEX"]:
        ret_map[sid] = (data_map[sid]["Close"] / data_map[sid]["Close"].iloc[0] - 1) * 100

    df_compare_pct = pd.DataFrame({
        stock_name_map["5475"]: ret_map["5475"],
        stock_name_map["3234"]: ret_map["3234"],
        stock_name_map["3105"]: ret_map["3105"],
        stock_name_map["3037"]: ret_map["3037"],
        stock_name_map["0050"]: ret_map["0050"],
        stock_name_map["TAIEX"]: ret_map["TAIEX"]
    }).dropna()

    st.line_chart(
        df_compare_pct,
        color=[color_5475, color_3234, color_3105, color_3037, color_0050, color_index]
    )

    st.divider()

    # =============================
    # 圖表 3：相對大盤超額績效
    # =============================
    st.subheader("相對大盤超額績效")
    st.markdown("以 **加權指數 (TAIEX)** 為比較基準，計算各標的相對大盤的超額報酬。")

    excess_return_df = pd.DataFrame({
        "德宏 (5475) - 大盤": ret_map["5475"] - ret_map["TAIEX"],
        "光環 (3234) - 大盤": ret_map["3234"] - ret_map["TAIEX"],
        "穩懋 (3105) - 大盤": ret_map["3105"] - ret_map["TAIEX"],
        "欣興 (3037) - 大盤": ret_map["3037"] - ret_map["TAIEX"],
        "0050 - 大盤": ret_map["0050"] - ret_map["TAIEX"]
    }).dropna()

    st.line_chart(
        excess_return_df,
        color=[color_5475, color_3234, color_3105, color_3037, color_0050]
    )

    st.caption("0% 代表與大盤相同；正值代表優於大盤，負值代表落後大盤。")

    st.divider()

    # =============================
    # 相對大盤績效摘要
    # =============================
    st.subheader("相對大盤績效摘要")

    summary_rows = []
    for sid in ["5475", "3234", "3105", "3037", "0050"]:
        summary_rows.append({
            "標的": stock_name_map[sid],
            "累計報酬(%)": round(ret_map[sid].iloc[-1], 2),
            "大盤報酬(%)": round(ret_map["TAIEX"].iloc[-1], 2),
            "超額績效(%)": round((ret_map[sid] - ret_map["TAIEX"]).iloc[-1], 2)
        })

    summary_df = pd.DataFrame(summary_rows)

    def highlight_excess(val):
        try:
            val = float(val)
            if val > 0:
                return "color: red; font-weight: bold;"
            elif val < 0:
                return "color: green; font-weight: bold;"
            return ""
        except Exception:
            return ""

    st.dataframe(
        summary_df.style.format({
            "累計報酬(%)": "{:.2f}",
            "大盤報酬(%)": "{:.2f}",
            "超額績效(%)": "{:.2f}"
        }).map(highlight_excess, subset=["超額績效(%)"]),
        use_container_width=True
    )

    best_row = summary_df.loc[summary_df["超額績效(%)"].idxmax()]
    worst_row = summary_df.loc[summary_df["超額績效(%)"].idxmin()]

    st.markdown(
        f"""
**摘要**
- 超額績效最佳：**{best_row['標的']}**，超額績效 **{best_row['超額績效(%)']:.2f}%**
- 超額績效最弱：**{worst_row['標的']}**，超額績效 **{worst_row['超額績效(%)']:.2f}%**
"""
    )

    st.divider()

    # =============================
    # 法人近 10 日買賣超
    # =============================
    st.subheader("近 10 日法人買賣超")
    st.caption("5475、3234、3105、3037、0050 與大盤法人買賣超資料。")

    chip_start_date = (today - relativedelta(days=45)).strftime("%Y-%m-%d")
    chip_end_date = end_date

    raw_5475 = load_institutional_data("5475", chip_start_date, chip_end_date, is_market_total=False)
    raw_3234 = load_institutional_data("3234", chip_start_date, chip_end_date, is_market_total=False)
    raw_3105 = load_institutional_data("3105", chip_start_date, chip_end_date, is_market_total=False)
    raw_3037 = load_institutional_data("3037", chip_start_date, chip_end_date, is_market_total=False)
    raw_0050 = load_institutional_data("0050", chip_start_date, chip_end_date, is_market_total=False)
    raw_market = load_institutional_data(None, chip_start_date, chip_end_date, is_market_total=True)

    table_5475 = summarize_institutional_table(raw_5475, "德宏 (5475)", last_n_days=10)
    table_3234 = summarize_institutional_table(raw_3234, "光環 (3234)", last_n_days=10)
    table_3105 = summarize_institutional_table(raw_3105, "穩懋 (3105)", last_n_days=10)
    table_3037 = summarize_institutional_table(raw_3037, "欣興 (3037)", last_n_days=10)
    table_0050 = summarize_institutional_table(raw_0050, "0050", last_n_days=10)
    table_market = summarize_institutional_table(raw_market, "加權指數(大盤)", last_n_days=10)

    def prepare_chip_table(df):
        if df is None or df.empty:
            return pd.DataFrame(columns=["日期", "外資買賣超", "投信買賣超", "自營商買賣超", "三大法人買賣超"])

        show_df = df.copy().sort_values("日期").tail(10)
        show_df["日期"] = pd.to_datetime(show_df["日期"]).dt.strftime("%Y-%m-%d")
        show_df = show_df[["日期", "外資買賣超", "投信買賣超", "自營商買賣超", "三大法人買賣超"]]
        return show_df

    chip_tables = {
        "德宏 (5475)": prepare_chip_table(table_5475),
        "光環 (3234)": prepare_chip_table(table_3234),
        "穩懋 (3105)": prepare_chip_table(table_3105),
        "欣興 (3037)": prepare_chip_table(table_3037),
        "0050": prepare_chip_table(table_0050),
        "加權指數(大盤)": prepare_chip_table(table_market),
    }

    chip_names = list(chip_tables.keys())
    for i in range(0, len(chip_names), 2):
        col1, col2 = st.columns(2)
        name1 = chip_names[i]
        with col1:
            st.markdown(f"### {name1}")
            if chip_tables[name1].empty:
                st.warning("無資料")
            else:
                st.dataframe(
                    style_net_table(chip_tables[name1]),
                    use_container_width=True,
                    height=420
                )

        if i + 1 < len(chip_names):
            name2 = chip_names[i + 1]
            with col2:
                st.markdown(f"### {name2}")
                if chip_tables[name2].empty:
                    st.warning("無資料")
                else:
                    st.dataframe(
                        style_net_table(chip_tables[name2]),
                        use_container_width=True,
                        height=420
                    )

    st.divider()

    # =============================
    # WantGoo 資料擷取
    # =============================
    st.subheader("WantGoo 籌碼資料")
    st.caption("抓取券商分點與大戶持股資料，作為籌碼判斷參考。")

    if "wantgoo_results" not in st.session_state:
        st.session_state["wantgoo_results"] = None

    refresh_col1, refresh_col2 = st.columns([1, 6])
    with refresh_col1:
        refresh_clicked = st.button("重新抓取")
    with refresh_col2:
        st.caption("Streamlit 重新整理時會保留 WantGoo 抓取結果在 session_state。")

    if refresh_clicked or st.session_state["wantgoo_results"] is None:
        with st.spinner("正在抓取 WantGoo 券商分點與大戶籌碼資料..."):
            st.session_state["wantgoo_results"] = load_wantgoo_all_signals(
                WANTGOO_USERNAME,
                WANTGOO_PASSWORD,
                TARGET_STOCK_IDS,
                WANTGOO_HEADLESS
            )

    wantgoo_results = st.session_state["wantgoo_results"]

    if (not WANTGOO_USERNAME) or (not WANTGOO_PASSWORD):
        st.warning("尚未設定 WANTGOO_USERNAME / WANTGOO_PASSWORD，WantGoo 資料將顯示為無資料。")

    # =============================
    # 券商分點前三名
    # =============================
    st.subheader("券商分點前三名（近 1 日）")
    broker_top3_df = build_broker_top3_table(wantgoo_results)
    st.dataframe(broker_top3_df, use_container_width=True)

    st.divider()

    # =============================
    # 大戶籌碼比較表
    # =============================
    st.subheader("大戶籌碼持股變化")
    major_holder_df = build_major_holder_table(wantgoo_results)

    def color_action(val: Any) -> str:
        if val == "買":
            return "color: red; font-weight: bold;"
        elif val == "不買":
            return "color: green; font-weight: bold;"
        elif val == "觀望":
            return "color: orange; font-weight: bold;"
        return ""

    def color_major_delta(row):
        styles = [""] * len(row)
        try:
            now_v = row["本週大戶持股比例(%)"]
            prev_v = row["上週大戶持股比例(%)"]
            row_columns = list(row.index)
            now_idx = row_columns.index("本週大戶持股比例(%)")
            prev_idx = row_columns.index("上週大戶持股比例(%)")
            if pd.notna(now_v) and pd.notna(prev_v):
                if now_v > prev_v:
                    styles[now_idx] = "color: red; font-weight: bold;"
                    styles[prev_idx] = "color: red;"
                elif now_v < prev_v:
                    styles[now_idx] = "color: green; font-weight: bold;"
                    styles[prev_idx] = "color: green;"
        except Exception:
            pass
        return styles

    st.dataframe(
        major_holder_df.style.format({
            "本週大戶持股比例(%)": "{:.2f}",
            "上週大戶持股比例(%)": "{:.2f}",
        }).apply(color_major_delta, axis=1).map(
            color_action,
            subset=["大戶籌碼建議"]
        ),
        use_container_width=True
    )

    st.divider()

    # =============================
    # 融資 / 融券 / 借券
    # =============================
    st.subheader("融資 / 融券 / 借券分析")
    st.caption("抓取近 20 日資料，顯示近 7 日融資、融券、借券變化與買賣訊號。")

    margin_data_map = load_all_margin_short_lending(
        TARGET_STOCK_IDS,
        margin_start_date,
        margin_end_date,
        FINMIND_TOKEN
    )

    margin_signal_summary_df = build_margin_signal_summary(margin_data_map)

    st.markdown("#### 融資融券借券建議總表")
    st.dataframe(
        margin_signal_summary_df.style.format({
            "最新融資餘額": "{:,.0f}",
            "最新融券餘額": "{:,.0f}",
            "最新借券量": "{:,.0f}",
        }).map(
            color_action,
            subset=["融資融券借券建議"]
        ),
        use_container_width=True
    )

    st.markdown("#### 各標的近一週融資 / 融券 / 借券走勢")

    for sid in TARGET_STOCK_IDS:
        label = stock_name_map_local.get(sid, sid)
        item = margin_data_map.get(sid, {})
        plot_df = item.get("df", pd.DataFrame())
        err = item.get("error")

        st.markdown(f"### {label}")

        if err:
            st.warning(f"{label}：{err}")
            continue

        if plot_df is None or plot_df.empty:
            st.warning(f"{label}：查無資料")
            continue

        chart_col, table_col = st.columns([1.3, 1])

        with chart_col:
            fig_margin = plot_margin_short_lending_chart(plot_df, label)
            if fig_margin is not None:
                st.pyplot(fig_margin)

        with table_col:
            st.dataframe(
                prepare_margin_detail_table(plot_df).style.format({
                    "融資餘額": "{:,.0f}",
                    "融券餘額": "{:,.0f}",
                    "借券量": "{:,.0f}",
                }).map(
                    color_action,
                    subset=["建議"]
                ),
                use_container_width=True,
                height=320
            )

    st.markdown(
        """
**融資融券借券規則**
- **不買**：融資餘額增加且融券餘額下降，代表籌碼可能偏弱。
- **買**：借券量下降，代表放空壓力可能減輕。
- **觀望**：訊號不明顯時先觀察。
- **無資料**：資料不足，暫時不產生判斷。
"""
    )

    st.divider()

    # =============================
    # 基本面分析
    # =============================
    st.subheader("基本面分析")

    has_fundamental_analysis = (
        not fundamental_summary_df.empty
        or not trend_df.empty
        or bool(group_tables)
    )
    if not has_fundamental_analysis:
        st.warning("基本面分析資料不足。")

    if not fundamental_summary_df.empty:
        st.markdown("#### 穩懋 / 欣興族群排名摘要")
        st.dataframe(fundamental_summary_df, use_container_width=True)

    if not trend_df.empty:
        st.markdown("#### 穩懋 / 欣興近三年獲利趨勢")
        st.dataframe(trend_df, use_container_width=True)

    if group_tables:
        st.markdown("#### 族群財務比較明細")
        for group_name, rows in group_tables.items():
            st.markdown(f"##### {group_name}")
            group_df = pd.DataFrame(rows)
            if group_df.empty:
                st.warning(f"{group_name} 無資料")
            else:
                st.dataframe(
                    group_df.style.format({
                        "ROE(%)": "{:.2f}",
                        "ROA(%)": "{:.2f}",
                        "EPS": "{:.2f}",
                        "毛利率(%)": "{:.2f}",
                        "流動比率": "{:.2f}",
                        "速動比率": "{:.2f}",
                    }, na_rep="無資料"),
                    use_container_width=True
                )

    st.divider()

    # =============================
    # 鞎瑁都撱箄降銵?
    # =============================
    st.subheader("買賣建議表（法人 + 券商分點 + 大戶籌碼 + 融資融券借券整合）")
    st.caption("整合法人籌碼、券商分點、大戶持股與融資融券借券訊號。")

    st.markdown("#### 基本面買賣建議表")
    if fundamental_rec_df.empty:
        st.warning("基本面資料不足，暫時無法產生建議。")
    else:
        st.dataframe(
            fundamental_rec_df.style.map(color_fundamental_action, subset=["基本面建議"]),
            use_container_width=True
        )

    recommendation_rows = [
        build_recommendation(table_5475, "德宏 (5475)"),
        build_recommendation(table_3234, "光環 (3234)"),
        build_recommendation(table_3105, "穩懋 (3105)"),
        build_recommendation(table_3037, "欣興 (3037)"),
        build_recommendation(table_0050, "0050"),
        build_recommendation(table_market, "加權指數(大盤)")
    ]
    recommendation_df = pd.DataFrame(recommendation_rows)

    broker_map: dict[str, dict[str, Any]] = {
        "0050": wantgoo_results.get("0050", {}),
        "光環 (3234)": wantgoo_results.get("3234", {}),
        "德宏 (5475)": wantgoo_results.get("5475", {}),
        "穩懋 (3105)": wantgoo_results.get("3105", {}),
        "欣興 (3037)": wantgoo_results.get("3037", {})
    }

    def _broker_row(label):
        row = broker_map.get(str(label), {})
        return row if isinstance(row, dict) else {}

    def _broker_list(label, key):
        values = _broker_row(label).get(key, [])
        return values if isinstance(values, list) else []

    def _broker_text(label, key, default="-"):
        value = _broker_row(label).get(key, default)
        return value if value not in (None, "") else default

    def _margin_text(label, column, default="-"):
        if margin_signal_summary_df is None or margin_signal_summary_df.empty:
            return default
        if "標的" not in margin_signal_summary_df.columns or column not in margin_signal_summary_df.columns:
            return default
        matched = cast(pd.Series, margin_signal_summary_df.loc[margin_signal_summary_df["標的"] == label, column])
        return matched.iat[0] if not matched.empty else default

    recommendation_df["買方前三名"] = recommendation_df["標的"].map(
        lambda x: "、".join(_broker_list(x, "buy_top3")) if _broker_list(x, "buy_top3") else "-"
    )
    recommendation_df["賣方前三名"] = recommendation_df["標的"].map(
        lambda x: "、".join(_broker_list(x, "sell_top3")) if _broker_list(x, "sell_top3") else "-"
    )
    recommendation_df["券商分點建議"] = recommendation_df["標的"].map(
        lambda x: _broker_text(x, "broker_signal", "無資料")
    )
    recommendation_df["券商分點說明"] = recommendation_df["標的"].map(
        lambda x: _broker_text(x, "broker_reason", "-")
    )
    recommendation_df["大戶籌碼建議"] = recommendation_df["標的"].map(
        lambda x: _broker_text(x, "major_signal", "無資料")
    )
    recommendation_df["大戶籌碼說明"] = recommendation_df["標的"].map(
        lambda x: _broker_text(x, "major_reason", "-")
    )
    recommendation_df["融資融券借券建議"] = recommendation_df["標的"].map(
        lambda x: _margin_text(x, "融資融券借券建議", "無資料")
    )
    recommendation_df["融資融券借券說明"] = recommendation_df["標的"].map(
        lambda x: _margin_text(x, "融資融券借券說明", "-")
    )
    recommendation_df["綜合建議"] = recommendation_df.apply(
        lambda row: combine_suggestion(
            row.get("法人建議"),
            row.get("券商分點建議"),
            row.get("大戶籌碼建議"),
            row.get("融資融券借券建議")
        ),
        axis=1
    )

    def color_numeric(v):
        try:
            if pd.isna(v):
                return ""
            v = float(v)
            if v > 0:
                return "color: red; font-weight: bold;"
            elif v < 0:
                return "color: green; font-weight: bold;"
            return ""
        except Exception:
            return ""

    st.dataframe(
        recommendation_df.style.format({
            "近三天外資買賣超": "{:,.0f}",
            "近三天投信買賣超": "{:,.0f}",
            "近三天自營商買賣超": "{:,.0f}",
            "近三天三大法人買賣超": "{:,.0f}",
        }).map(
            color_action,
            subset=["法人建議", "券商分點建議", "大戶籌碼建議", "融資融券借券建議", "綜合建議"]
        ).map(
            color_numeric,
            subset=[
                "近三天外資買賣超",
                "近三天投信買賣超",
                "近三天自營商買賣超",
                "近三天三大法人買賣超",
            ]
        ),
        use_container_width=True
    )

    st.markdown(
        """
**規則說明**
- **法人建議**：投信近三天買超視為不買；外資與三大法人近三天同步賣超視為不買；外資或三大法人近三天買超視為買；其餘為觀望。
- **券商分點建議**：依買方與賣方前三名是否出現自訂偏多、偏空券商分點判斷。
- **大戶籌碼建議**：本週大戶持股比例高於上週視為買，低於上週視為不買，無明顯變化則觀望。
- **融資融券借券建議**：融資增加且融券下降視為不買；借券量下降視為買；其餘為觀望。
- **綜合建議**：任一來源為不買則不買；沒有不買且任一來源為買則買；其餘為觀望。
"""
    )

    st.divider()


# --- Flask 即時儀表板 ---
if __name__ == "__main__":
    if "--streamlit" in sys.argv:
        run_streamlit_app()
        raise SystemExit(0)

    import json as _json
    import threading as _threading
    try:
        _flask_mod = importlib.import_module("flask")
    except ModuleNotFoundError as exc:
        raise ModuleNotFoundError(
            "找不到 flask。請先安裝：pip install flask"
        ) from exc
    _Flask = _flask_mod.Flask
    _Response = _flask_mod.Response
    _jsonify = _flask_mod.jsonify
    _rts = _flask_mod.render_template_string
    _DataLoader = importlib.import_module("FinMind.data").DataLoader
    from dateutil.relativedelta import relativedelta as _relativedelta

    _FINMIND_TOKEN = os.getenv(
        "FINMIND_TOKEN",
        "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJkYXRlIjoiMjAyNi0wMy0xOCAyMjozNToxNyIsInVzZXJfaWQiOiI4OTEwMDciLCJlbWFpbCI6ImxpbGxhcmQ4MDZAZ21haWwuY29tIiwiaXAiOiIxMTQuMzYuMTgxLjcxIn0.WEgK_gYl-WQfRxxMR9bVrvkMsltT1pHfO_TEmvvIsMU"
    )

    _TARGET_IDS = ["0050", "3234", "5475", "3105", "3037"]
    _NAME_MAP = {
        "0050": "0050", "3234": "光環(3234)", "5475": "德宏(5475)",
        "3105": "穩懋(3105)", "3037": "欣興(3037)", "TAIEX": "加權指數(TAIEX)"
    }

    # 快取狀態
    _cache = {
        "status": "waiting",
        "last_updated": None,
        "error_msg": "",
        "is_running": False,
        "portfolio_total_cost": 0,
        "portfolio_total_value": 0,
        "portfolio_total_profit": 0,
        "portfolio_total_return": 0,
        "portfolio_rows": [],
        "perf_summary": [],
        "chart_dates": [],
        "chart_prices": {},
        "chart_returns": {},
        "chart_excess": {},
        "chip_tables": {},
        "margin_summary": [],
        "margin_tables": {},
        "recommendation_rows": [],
        "fundamental": {
            "summary_rows": [],
            "recommendation_rows": [],
            "trend_rows": [],
            "group_tables": {},
        },
        "dca_rows": [],
        "dca_total_cost": 0,
        "dca_total_shares": 0,
    }

    def _combine_sig(*signals):
        norms = []
        for x in signals:
            if not x:
                continue
            s = str(x).strip()
            if s in ("買", "建議買"):
                norms.append("買")
            elif s in ("不買", "建議不買"):
                norms.append("不買")
            elif s in ("觀望", "無資料"):
                norms.append(s)
        if not norms:
            return "無資料"
        if "不買" in norms:
            return "不買"
        if "買" in norms:
            return "買"
        return "觀望"

    # 資料擷取主流程
    def _run_scraper():
        if _cache["is_running"]:
            return
        _cache["is_running"] = True
        _cache["status"] = "scraping"
        _cache["error_msg"] = ""
        try:
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

            today = datetime.today()
            one_year_ago = today - _relativedelta(years=1)
            start_date_main = one_year_ago.strftime("%Y-%m-%d")
            end_date = today.strftime("%Y-%m-%d")
            start_date_0050 = "2025-06-18"

            # FinMind 登入
            api = _DataLoader()
            try:
                login_by_token = getattr(api, "login_by_token")
                login_by_token(api_token=_FINMIND_TOKEN)
            except AttributeError:
                login = getattr(api, "login")
                login(token=_FINMIND_TOKEN)

            # 股票資料設定
            stock_config = {
                "5475": start_date_main, "3234": start_date_main,
                "3105": start_date_main, "3037": start_date_main,
                "0050": start_date_0050, "TAIEX": start_date_main,
            }
            raw_data = {}
            for sid, sdate in stock_config.items():
                raw_data[sid] = api.taiwan_stock_daily(stock_id=sid, start_date=sdate, end_date=end_date)

            def _proc(df):
                df = df.rename(columns={
                    "date": "Date", "open": "Open", "max": "High",
                    "min": "Low", "close": "Close",
                    "Trading_Volume": "Volume", "volume": "Volume"
                })
                df["Date"] = pd.to_datetime(df["Date"])
                df = df.sort_values("Date").set_index("Date")
                for col in ["Open", "High", "Low", "Close", "Volume"]:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors="coerce")
                return df

            processed = {sid: _proc(df) for sid, df in raw_data.items()}
            common_index = None
            for df in processed.values():
                common_index = df.index if common_index is None else common_index.intersection(df.index)
            assert common_index is not None
            aligned = {sid: df.loc[common_index].copy() for sid, df in processed.items()}

            # 0050 定期定額模擬
            df_0050_full = _proc(api.taiwan_stock_daily(stock_id="0050", start_date=start_date_0050, end_date=end_date))
            df_0050_full["MA60"] = df_0050_full["Close"].rolling(60, min_periods=20).mean()
            tdates = pd.DatetimeIndex(df_0050_full.index).sort_values()
            sim_start = pd.to_datetime(start_date_0050).normalize()
            last_date = tdates.max().normalize()
            mc = pd.Timestamp(sim_start.year, sim_start.month, 1)
            me = pd.Timestamp(last_date.year, last_date.month, 1)
            dca_records, used_dates = [], set()
            while mc <= me:
                for day in [1, 7, 13, 19, 25]:
                    try:
                        pdate = pd.Timestamp(year=mc.year, month=mc.month, day=day)
                    except ValueError:
                        continue
                    if pdate < sim_start or pdate > last_date:
                        continue
                    cands = tdates[tdates >= pdate]
                    if not len(cands):
                        continue
                    tdate = pd.Timestamp(cands[0])
                    if tdate in used_dates:
                        continue
                    loc = int(tdates.get_indexer(pd.DatetimeIndex([tdate]))[0])
                    if loc <= 0:
                        continue
                    pc = float(df_0050_full["Close"].iloc[loc - 1])
                    qa = float(df_0050_full["MA60"].iloc[loc - 1])
                    if pd.isna(qa) or qa == 0:
                        amt = 1000
                    else:
                        diff = (pc / qa - 1) * 100
                        amt = 2000 if diff <= -5 else (800 if diff >= 5 else 1000)
                    bp = float(df_0050_full["Close"].iloc[loc])
                    bs = amt / bp if bp else 0
                    dca_records.append({
                        "預計買進日": pdate.strftime("%Y-%m-%d"),
                        "實際買進日": tdate.strftime("%Y-%m-%d"),
                        "投入金額": amt,
                        "買進價格": round(bp, 2),
                        "買進股數": round(bs, 4),
                    })
                    used_dates.add(tdate)
                mc += _relativedelta(months=1)
            dca_df = pd.DataFrame(dca_records)
            dca_total_cost = float(dca_df["投入金額"].sum()) if not dca_df.empty else 0.0
            dca_total_shares = float(dca_df["買進股數"].sum()) if not dca_df.empty else 0.0

            # 投資組合
            df_5475 = aligned["5475"]
            df_3234 = aligned["3234"]
            df_0050 = aligned["0050"]
            pdata = [
                {"標的": "德宏(5475)", "代號": "5475", "平均成本": 108.5, "原始投入成本": 108543, "現價": float(df_5475["Close"].iloc[-1])},
                {"標的": "光環(3234)", "代號": "3234", "平均成本": 86.47, "原始投入成本": 129751, "現價": float(df_3234["Close"].iloc[-1])},
                {"標的": "0050", "代號": "0050", "平均成本": 53.92, "原始投入成本": 48054, "現價": float(df_0050["Close"].iloc[-1])},
            ]
            dp = pd.DataFrame(pdata)
            dp["原始股數"] = dp["原始投入成本"] / dp["平均成本"]
            dp["定期定額成本"] = dp["代號"].apply(lambda x: dca_total_cost if x == "0050" else 0.0)
            dp["定期定額股數"] = dp["代號"].apply(lambda x: dca_total_shares if x == "0050" else 0.0)
            dp["總成本"] = dp["原始投入成本"] + dp["定期定額成本"]
            dp["總股數"] = dp["原始股數"] + dp["定期定額股數"]
            dp["市值"] = dp["總股數"] * dp["現價"]
            dp["損益"] = dp["市值"] - dp["總成本"]
            dp["報酬率(%)"] = dp["損益"] / dp["總成本"] * 100
            total_cost = float(dp["總成本"].sum())
            total_value = float(dp["市值"].sum())
            total_profit = float(dp["損益"].sum())
            total_return = (total_profit / total_cost * 100) if total_cost else 0.0
            portfolio_rows = [
                {"標的": r["標的"], "平均成本": round(float(r["平均成本"]), 2),
                 "總成本": round(float(r["總成本"]), 0),
                 "現價": round(float(r["現價"]), 2),
                 "總股數": round(float(r["總股數"]), 3),
                 "市值": round(float(r["市值"]), 0),
                 "損益": round(float(r["損益"]), 0),
                 "報酬率(%)": round(float(r["報酬率(%)"]), 2)}
                for _, r in dp.iterrows()
            ]

            # 圖表資料
            chart_dates = [d.strftime("%Y-%m-%d") for d in common_index]
            chart_prices: dict[str, list[float]] = {}
            chart_returns: dict[str, list[float]] = {}
            chart_excess: dict[str, list[float]] = {}
            for sid in ["5475", "3234", "3105", "3037", "0050", "TAIEX"]:
                closes = [round(float(v), 2) for v in aligned[sid]["Close"].values]
                chart_prices[sid] = closes
                c0 = closes[0] if closes[0] != 0 else 1
                chart_returns[sid] = [round((v / c0 - 1) * 100, 2) for v in closes]
            for sid in ["5475", "3234", "3105", "3037", "0050"]:
                chart_excess[sid] = [
                    round(chart_returns[sid][i] - chart_returns["TAIEX"][i], 2)
                    for i in range(len(chart_dates))
                ]
            perf_summary = [
                {"標的": _NAME_MAP.get(sid, sid),
                 "累計報酬(%)": chart_returns[sid][-1] if chart_returns[sid] else 0,
                 "大盤報酬(%)": chart_returns["TAIEX"][-1] if chart_returns["TAIEX"] else 0,
                 "超額績效(%)": chart_excess[sid][-1] if chart_excess[sid] else 0}
                for sid in ["5475", "3234", "3105", "3037", "0050"]
            ]

            # 法人籌碼
            chip_start = (today - _relativedelta(days=45)).strftime("%Y-%m-%d")

            def _load_chip(sid, is_total=False):
                try:
                    if is_total:
                        df = api.taiwan_stock_institutional_investors_total(
                            start_date=chip_start, end_date=end_date)
                    else:
                        df = api.taiwan_stock_institutional_investors(
                            stock_id=sid, start_date=chip_start, end_date=end_date)
                    if df is None or df.empty:
                        return []
                    df = df.copy()
                    df["date"] = pd.to_datetime(df["date"])
                    df["buy"]  = pd.to_numeric(df["buy"],  errors="coerce").fillna(0)
                    df["sell"] = pd.to_numeric(df["sell"], errors="coerce").fillna(0)
                    df["net"]  = df["buy"] - df["sell"]
                    grp   = df.groupby(["date", "name"], as_index=False).agg({"net": "sum"})
                    pivot = grp.pivot(index="date", columns="name", values="net").fillna(0)
                    def _gs(col):
                        return pivot[col] if col in pivot.columns else pd.Series(0, index=pivot.index)
                    foreign = _gs("Foreign_Investor") + _gs("Foreign_Dealer_Self")
                    trust   = _gs("Investment_Trust")
                    dealer  = _gs("Dealer_self") + _gs("Dealer_Hedging")
                    total3  = foreign + trust + dealer
                    result = pd.DataFrame({
                        "日期": pivot.index,
                        "外資買賣超": foreign.values,
                        "投信買賣超": trust.values,
                        "自營商買賣超": dealer.values,
                        "三大法人買賣超": total3.values,
                    }).sort_values("日期").tail(10)
                    return [{"日期": r["日期"].strftime("%Y-%m-%d"),
                             "外資買賣超": int(r["外資買賣超"]),
                             "投信買賣超": int(r["投信買賣超"]),
                             "自營商買賣超": int(r["自營商買賣超"]),
                             "三大法人買賣超": int(r["三大法人買賣超"])}
                            for _, r in result.iterrows()]
                except Exception:
                    return []

            chip_tables = {}
            for sid, name in [("5475", "德宏(5475)"), ("3234", "光環(3234)"),
                               ("3105", "穩懋(3105)"), ("3037", "欣興(3037)"), ("0050", "0050")]:
                chip_tables[name] = _load_chip(sid)
            chip_tables["加權指數(大盤)"] = _load_chip(None, is_total=True)

            def _build_rec(rows, name):
                if not rows:
                    return {"標的": name, "近三天外資買賣超": None, "近三天投信買賣超": None,
                            "近三天三大法人買賣超": None, "法人建議": "無資料", "法人說明": "查無資料"}
                r3   = rows[-3:]
                f3   = sum(r["外資買賣超"] for r in r3)
                t3   = sum(r["投信買賣超"] for r in r3)
                tot3 = sum(r["三大法人買賣超"] for r in r3)
                if t3 > 0:
                    sug, rea = "不買", "投信近三天買超，依原規則判斷為不買"
                elif f3 < 0 and tot3 < 0:
                    sug, rea = "不買", "外資與三大法人近三天同步賣超"
                elif f3 > 0 or tot3 > 0:
                    sug, rea = "買", "外資或三大法人近三天買超"
                else:
                    sug, rea = "觀望", "法人籌碼沒有明確買賣訊號"
                return {"標的": name, "近三天外資買賣超": f3, "近三天投信買賣超": t3,
                        "近三天三大法人買賣超": tot3, "法人建議": sug, "法人說明": rea}

            chip_rec = {
                name: _build_rec(chip_tables.get(name, []), name)
                for name in ["德宏(5475)", "光環(3234)", "穩懋(3105)", "欣興(3037)", "0050"]
            }

            # 融資融券借券
            _FINMIND_BASE = "https://api.finmindtrade.com/api/v4/data"
            m_start = (today.date() - timedelta(days=20)).strftime("%Y-%m-%d")
            m_end   = today.date().strftime("%Y-%m-%d")

            def _fetch_margin(sid):
                try:
                    hdrs = {"Authorization": f"Bearer {_FINMIND_TOKEN}"}
                    r = requests.get(_FINMIND_BASE, headers=hdrs, params={
                        "dataset": "TaiwanStockMarginPurchaseShortSale",
                        "data_id": sid, "start_date": m_start, "end_date": m_end
                    }, timeout=30)
                    r.raise_for_status()
                    data = r.json()
                    if "data" not in data or not data["data"]:
                        return [], "無資料"
                    dfm = pd.DataFrame(data["data"])
                    dfm["date"] = pd.to_datetime(dfm["date"])
                    r2 = requests.get(_FINMIND_BASE, headers=hdrs, params={
                        "dataset": "TaiwanStockSecuritiesLending",
                        "data_id": sid, "start_date": m_start, "end_date": m_end
                    }, timeout=30)
                    r2.raise_for_status()
                    data2 = r2.json()
                    if "data" in data2 and data2["data"]:
                        dfl = pd.DataFrame(data2["data"])
                        dfl["date"] = pd.to_datetime(dfl["date"])
                        dfl_d = (
                            dfl.groupby("date", as_index=False)
                            .agg({"volume": "sum"})
                            .rename(columns={"volume": "借券量"})
                        )
                    else:
                        dfl_d = pd.DataFrame(columns=["date", "借券量"])
                    df = pd.merge(
                        dfm[["date", "MarginPurchaseTodayBalance", "ShortSaleTodayBalance"]],
                        dfl_d, on="date", how="left"
                    )
                    df["借券量"] = df["借券量"].fillna(0)
                    df = df.sort_values("date").tail(7).reset_index(drop=True)
                    df["p_m"] = df["MarginPurchaseTodayBalance"].shift(1)
                    df["p_s"] = df["ShortSaleTodayBalance"].shift(1)
                    df["p_l"] = df["借券量"].shift(1)

                    def _sig(row):
                        if pd.isna(row["p_m"]):
                            return "無資料"
                        if (row["MarginPurchaseTodayBalance"] > row["p_m"]
                                and row["ShortSaleTodayBalance"] < row["p_s"]):
                            return "建議不買"
                        if row["借券量"] < row["p_l"]:
                            return "建議買"
                        return "觀望"

                    df["建議"] = df.apply(_sig, axis=1)
                    rows = [{"日期": row["date"].strftime("%Y-%m-%d"),
                             "融資餘額": int(row["MarginPurchaseTodayBalance"]),
                             "融券餘額": int(row["ShortSaleTodayBalance"]),
                             "借券量": int(row["借券量"]), "建議": row["建議"]}
                            for _, row in df.iterrows()]
                    return rows, None
                except Exception as e:
                    return [], str(e)

            margin_tables, margin_summary = {}, []
            for sid in _TARGET_IDS:
                name = _NAME_MAP.get(sid, sid)
                rows, err = _fetch_margin(sid)
                margin_tables[name] = rows
                if rows:
                    sig = rows[-1]["建議"]
                    fsig = {"建議買": "買", "建議不買": "不買", "觀望": "觀望"}.get(sig, "無資料")
                    margin_summary.append({"標的": name, "最新建議": fsig, "說明": sig})
                else:
                    margin_summary.append({"標的": name, "最新建議": "無資料", "說明": err or "無資料"})

            margin_sig_map = {r["標的"]: r["最新建議"] for r in margin_summary}

            # 撱箄降銵?
            recommendation_rows = []
            for sid in ["5475", "3234", "3105", "3037", "0050"]:
                name = _NAME_MAP.get(sid, sid)
                rec  = chip_rec.get(name, {})
                m_sig = margin_sig_map.get(name, "無資料")
                recommendation_rows.append({
                    "標的": name,
                    "近三天外資買賣超": rec.get("近三天外資買賣超"),
                    "近三天投信買賣超": rec.get("近三天投信買賣超"),
                    "近三天三大法人買賣超": rec.get("近三天三大法人買賣超"),
                    "法人建議": rec.get("法人建議", "無資料"),
                    "法人說明": rec.get("法人說明", ""),
                    "融資融券建議": m_sig,
                    "綜合建議": _combine_sig(rec.get("法人建議"), m_sig),
                })

            fundamental_payload = build_fundamental_payload(_FINMIND_TOKEN, today=today)

            _cache.update({
                "status": "ready",
                "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "portfolio_total_cost":   round(total_cost, 0),
                "portfolio_total_value":  round(total_value, 0),
                "portfolio_total_profit": round(total_profit, 0),
                "portfolio_total_return": round(total_return, 2),
                "portfolio_rows": portfolio_rows,
                "perf_summary":   perf_summary,
                "chart_dates":    chart_dates,
                "chart_prices":   chart_prices,
                "chart_returns":  chart_returns,
                "chart_excess":   chart_excess,
                "chip_tables":    chip_tables,
                "margin_summary": margin_summary,
                "margin_tables":  margin_tables,
                "recommendation_rows": recommendation_rows,
                "fundamental": fundamental_payload,
                "dca_rows":         dca_records,
                "dca_total_cost":   round(dca_total_cost, 0),
                "dca_total_shares": round(dca_total_shares, 3),
            })
        except Exception as e:
            _cache["status"]    = "error"
            _cache["error_msg"] = str(e)
        finally:
            _cache["is_running"] = False

    def _auto_scheduler():
        _threading.Thread(target=_run_scraper, daemon=True).start()
        while True:
            now = datetime.now()
            if now.hour == 8 and now.minute == 0:
                _threading.Thread(target=_run_scraper, daemon=True).start()
                time.sleep(61)
            time.sleep(30)

    _threading.Thread(target=_auto_scheduler, daemon=True).start()

    # Flask App
    app = _Flask(__name__)

    @app.route("/api/data")
    def api_data():
        return _jsonify(_cache)

    @app.route("/api/refresh", methods=["POST"])
    def api_refresh():
        if _cache["is_running"]:
            return _jsonify({"message": "資料正在更新中，請稍候。"})
        _threading.Thread(target=_run_scraper, daemon=True).start()
        return _jsonify({"message": "已開始重新抓取資料，完成後畫面會自動更新。"})

    @app.route("/api/stream")
    def api_stream():
        def generate():
            last_hash = None
            while True:
                snap = {k: _cache[k] for k in (
                    "status", "last_updated", "error_msg",
                    "portfolio_total_cost", "portfolio_total_value",
                    "portfolio_total_profit", "portfolio_total_return"
                )}
                h = hash(_json.dumps(snap, sort_keys=True))
                if h != last_hash:
                    yield f"data: {_json.dumps(snap)}\n\n"
                    last_hash = h
                time.sleep(2)
        return _Response(generate(), mimetype="text/event-stream",
                         headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})

    _DASHBOARD_HTML = r"""
<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>股票交易儀表板 5475 / 3234 / 3105 / 3037 / 0050</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Microsoft JhengHei",sans-serif;background:#f0f2f5;color:#333}
header{background:#1a1a2e;color:#fff;padding:16px 32px;display:flex;justify-content:space-between;align-items:center}
header h1{font-size:1.3rem}
header .sub{font-size:.8rem;opacity:.7;margin-top:3px}
.container{max-width:1500px;margin:20px auto;padding:0 20px}
.status-bar{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:18px;align-items:center}
.badge{padding:5px 14px;border-radius:20px;font-size:.82rem;font-weight:bold}
.badge.waiting{background:#e9ecef;color:#495057}
.badge.scraping{background:#fff3cd;color:#856404;animation:pulse 1s infinite}
.badge.ready{background:#d4edda;color:#155724}
.badge.error{background:#f8d7da;color:#721c24}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.5}}
.last-update{font-size:.82rem;color:#666;margin-left:auto}
.btn{padding:7px 18px;background:#0d6efd;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:.88rem}
.btn:hover{background:#0b5ed7}
.btn:disabled{background:#6c757d;cursor:not-allowed}
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:14px;margin-bottom:22px}
.card{background:#fff;border-radius:10px;padding:16px 18px;box-shadow:0 1px 4px rgba(0,0,0,.08);text-align:center}
.card .num{font-size:1.8rem;font-weight:bold;color:#0d6efd}
.card .num.pos{color:#dc3545}
.card .num.neg{color:#198754}
.card .lbl{font-size:.78rem;color:#888;margin-top:4px}
.tabs{display:flex;gap:0;margin-bottom:0;border-bottom:2px solid #dee2e6;overflow-x:auto}
.tab{padding:9px 18px;cursor:pointer;font-size:.88rem;border:1px solid transparent;border-bottom:none;margin-bottom:-2px;white-space:nowrap;color:#666;background:none}
.tab.active{background:#fff;border-color:#dee2e6;border-bottom-color:#fff;color:#0d6efd;font-weight:bold}
.tab-content{display:none;background:#fff;border:1px solid #dee2e6;border-top:none;padding:20px;border-radius:0 0 8px 8px}
.tab-content.active{display:block}
.chart-wrap{position:relative;height:350px;margin-bottom:12px}
table{width:100%;border-collapse:collapse;font-size:.85rem}
th{background:#1a1a2e;color:#fff;padding:9px 12px;text-align:left;position:sticky;top:0;white-space:nowrap}
td{padding:8px 12px;border-bottom:1px solid #f0f0f0;white-space:nowrap}
tr:hover td{background:#f8f9fa}
.tbl-wrap{overflow:auto;max-height:420px;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.08)}
.pos{color:#dc3545;font-weight:bold}
.neg{color:#198754;font-weight:bold}
.badge-buy{background:#dc3545;color:#fff;padding:2px 8px;border-radius:10px;font-size:.78rem}
.badge-sell{background:#198754;color:#fff;padding:2px 8px;border-radius:10px;font-size:.78rem}
.badge-watch{background:#fd7e14;color:#fff;padding:2px 8px;border-radius:10px;font-size:.78rem}
.badge-na{background:#6c757d;color:#fff;padding:2px 8px;border-radius:10px;font-size:.78rem}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:18px}
.error-msg{color:#721c24;background:#f8d7da;padding:10px;border-radius:6px;margin-bottom:14px}
h3{margin:0 0 12px;font-size:1rem;color:#1a1a2e}
@media(max-width:768px){.grid2{grid-template-columns:1fr}}
</style>
</head>
<body>
<header>
  <div>
    <h1>股票交易儀表板</h1>
    <div class="sub">5475 德宏 / 3234 光環 / 3105 穩懋 / 3037 欣興 / 0050 / 加權指數 (TAIEX)</div>
  </div>
  <div id="clock" style="font-size:.8rem;opacity:.7;text-align:right"></div>
</header>

<div class="container">
  <div class="status-bar">
    <span id="badge" class="badge waiting">等待中</span>
    <span id="last-update" class="last-update">尚未更新</span>
    <button class="btn" id="refreshBtn" onclick="triggerRefresh()">重新整理資料</button>
  </div>
  <div id="error-box" class="error-msg" style="display:none"></div>

  <!-- 投資組合摘要 -->
  <div class="cards">
    <div class="card"><div class="num" id="c-cost">-</div><div class="lbl">總成本</div></div>
    <div class="card"><div class="num" id="c-value">-</div><div class="lbl">目前市值</div></div>
    <div class="card"><div class="num" id="c-profit">-</div><div class="lbl">損益</div></div>
    <div class="card"><div class="num" id="c-return">-</div><div class="lbl">報酬率</div></div>
  </div>

  <!-- Tabs -->
  <div class="tabs">
    <div class="tab active" onclick="switchTab(0)">股價走勢</div>
    <div class="tab" onclick="switchTab(1)">累計報酬</div>
    <div class="tab" onclick="switchTab(2)">超額績效</div>
    <div class="tab" onclick="switchTab(3)">投資組合</div>
    <div class="tab" onclick="switchTab(4)">法人籌碼</div>
    <div class="tab" onclick="switchTab(5)">融資融券借券</div>
    <div class="tab" onclick="switchTab(6)">基本面分析</div>
    <div class="tab" onclick="switchTab(7)">綜合建議</div>
  </div>

  <!-- Tab 0: 股價走勢 -->
  <div class="tab-content active" id="tab0">
    <div class="chart-wrap"><canvas id="chartPrice"></canvas></div>
    <p style="font-size:.8rem;color:#888;margin-top:6px">顯示各標的收盤價走勢；0050 從 2025-06-18 開始納入。</p>
  </div>

  <!-- Tab 1: 累計報酬 -->
  <div class="tab-content" id="tab1">
    <div class="chart-wrap"><canvas id="chartReturn"></canvas></div>
    <div class="tbl-wrap" style="margin-top:16px" id="perfTable"></div>
  </div>

  <!-- Tab 2: 超額績效 -->
  <div class="tab-content" id="tab2">
    <div class="chart-wrap"><canvas id="chartExcess"></canvas></div>
    <p style="font-size:.8rem;color:#888;margin-top:6px">0% 代表與大盤相同；正值代表優於大盤，負值代表落後大盤。</p>
  </div>

  <!-- Tab 3: 投資組合 -->
  <div class="tab-content" id="tab3">
    <div class="tbl-wrap" id="portfolioTable"></div>
    <div style="margin-top:14px;font-size:.85rem;color:#555" id="dcaInfo"></div>
  </div>

  <!-- Tab 4: 法人籌碼 -->
  <div class="tab-content" id="tab4">
    <div id="chipTablesWrap"></div>
  </div>

  <!-- Tab 5: 融資融券借券 -->
  <div class="tab-content" id="tab5">
    <h3>融資融券借券建議摘要</h3>
    <div class="tbl-wrap" id="marginSummaryTable"></div>
    <div id="marginDetailWrap" style="margin-top:18px"></div>
  </div>

  <!-- Tab 6: 基本面分析 -->
  <div class="tab-content" id="tab6">
    <div id="fundamentalAnalysisWrap"></div>
  </div>

  <!-- Tab 7: 綜合建議 -->
  <div class="tab-content" id="tab7">
    <div class="tbl-wrap" id="recTable"></div>
    <div style="margin-top:12px;font-size:.82rem;color:#666">
      <b>法人規則：</b>投信近三天買超視為不買；外資與三大法人同步賣超視為不買；外資或三大法人買超視為買；其餘為觀望。<br>
      <b>融資融券規則：</b>融資增加且融券下降視為不買；借券量下降視為買；其餘為觀望。
    </div>
    <div id="fundamentalRecWrap" style="margin-top:22px"></div>
  </div>
</div>

<script>
let fullData = {};
const COLORS = {
  "5475":"#d65f4a","3234":"#9467bd","3105":"#ff7f0e",
  "3037":"#17becf","0050":"#2ca02c","TAIEX":"#5b9bd5"
};
const NAMES = {
  "5475":"德宏(5475)","3234":"光環(3234)","3105":"穩懋(3105)",
  "3037":"欣興(3037)","0050":"0050","TAIEX":"加權指數(TAIEX)"
};

// SSE 狀態更新
const es = new EventSource("/api/stream");
es.onmessage = (e) => {
  const d = JSON.parse(e.data);
  const badge = document.getElementById("badge");
  const lblMap = {waiting:"等待中",scraping:"資料更新中...",ready:"已更新",error:"更新失敗"};
  badge.textContent = lblMap[d.status] || d.status;
  badge.className = "badge " + d.status;
  document.getElementById("last-update").textContent =
    d.last_updated ? "最後更新：" + d.last_updated : "尚未更新";
  const errBox = document.getElementById("error-box");
  if (d.status === "error") {
    errBox.textContent = "更新失敗：" + d.error_msg;
    errBox.style.display = "block";
  } else {
    errBox.style.display = "none";
  }
  // 更新投資組合摘要卡片
  if (d.portfolio_total_cost) {
    const profit = d.portfolio_total_profit;
    const ret    = d.portfolio_total_return;
    document.getElementById("c-cost").textContent  = "$" + fmt(d.portfolio_total_cost);
    document.getElementById("c-value").textContent = "$" + fmt(d.portfolio_total_value);
    const pe = document.getElementById("c-profit");
    pe.textContent = "$" + fmt(profit);
    pe.className   = "num " + (profit >= 0 ? "pos" : "neg");
    const re = document.getElementById("c-return");
    re.textContent = ret.toFixed(2) + "%";
    re.className   = "num " + (ret >= 0 ? "pos" : "neg");
  }
  if (d.status === "ready") loadFullData();
  document.getElementById("refreshBtn").disabled = (d.status === "scraping");
};

function fmt(n) { return Number(n).toLocaleString("zh-TW"); }

// 載入完整資料
function loadFullData() {
  fetch("/api/data")
    .then(r => r.json())
    .then(data => {
      fullData = data;
      renderCharts(data);
      renderTables(data);
    });
}

// 圖表資料抽樣
let chartPrice = null, chartReturn = null, chartExcess = null;

function sampleData(dates, seriesMap, maxPoints = 180) {
  const n = dates.length;
  if (n <= maxPoints) return { dates, seriesMap };
  const step = Math.ceil(n / maxPoints);
  const idx = [];
  for (let i = 0; i < n; i += step) idx.push(i);
  if (idx[idx.length - 1] !== n - 1) idx.push(n - 1);
  const newDates = idx.map(i => dates[i]);
  const newSeries = {};
  for (const [k, v] of Object.entries(seriesMap)) {
    newSeries[k] = idx.map(i => v[i]);
  }
  return { dates: newDates, seriesMap: newSeries };
}

function renderCharts(data) {
  if (!data.chart_dates || !data.chart_dates.length) return;

  // 股價走勢
  const priceKeys = ["5475","3234","3105","3037","0050","TAIEX"];
  const { dates: pd, seriesMap: pm } = sampleData(data.chart_dates, data.chart_prices);
  if (chartPrice) chartPrice.destroy();
  chartPrice = new Chart(document.getElementById("chartPrice"), {
    type: "line",
    data: {
      labels: pd,
      datasets: priceKeys.map(k => ({
        label: NAMES[k], data: pm[k], borderColor: COLORS[k],
        pointRadius: 0, borderWidth: 2, tension: 0.1,
        yAxisID: k === "TAIEX" ? "y2" : "y1"
      }))
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: { legend: { position: "top" } },
      scales: {
        y1: { type: "linear", position: "left",  display: false },
        y2: { type: "linear", position: "right", display: false }
      }
    }
  });

  // 累計報酬
  const retKeys = ["5475","3234","3105","3037","0050","TAIEX"];
  const { dates: rd, seriesMap: rm } = sampleData(data.chart_dates, data.chart_returns);
  if (chartReturn) chartReturn.destroy();
  chartReturn = new Chart(document.getElementById("chartReturn"), {
    type: "line",
    data: {
      labels: rd,
      datasets: retKeys.map(k => ({
        label: NAMES[k], data: rm[k], borderColor: COLORS[k],
        pointRadius: 0, borderWidth: 2, tension: 0.1
      }))
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: { legend: { position: "top" } },
      scales: {
        y: {
          ticks: { callback: v => v.toFixed(1) + "%" }
        }
      }
    }
  });

  // 超額績效
  const exKeys = ["5475","3234","3105","3037","0050"];
  const { dates: ed, seriesMap: em } = sampleData(data.chart_dates, data.chart_excess);
  if (chartExcess) chartExcess.destroy();
  chartExcess = new Chart(document.getElementById("chartExcess"), {
    type: "line",
    data: {
      labels: ed,
      datasets: exKeys.map(k => ({
        label: NAMES[k], data: em[k], borderColor: COLORS[k],
        pointRadius: 0, borderWidth: 2, tension: 0.1,
        fill: false
      }))
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: { legend: { position: "top" } },
      scales: {
        y: {
          ticks: { callback: v => v.toFixed(1) + "%" },
          grid: { color: ctx => ctx.tick.value === 0 ? "#333" : "#e5e5e5" }
        }
      }
    }
  });
}

// 顯示輔助函式
function sigBadge(v) {
  if (v === "買")   return `<span class="badge-buy">買</span>`;
  if (v === "不買") return `<span class="badge-sell">不買</span>`;
  if (v === "觀望") return `<span class="badge-watch">觀望</span>`;
  return `<span class="badge-na">${v || "無資料"}</span>`;
}
function numColor(v) {
  if (v === null || v === undefined) return "-";
  const n = Number(v);
  if (isNaN(n)) return v;
  const s = n.toLocaleString("zh-TW");
  if (n > 0) return `<span class="pos">${s}</span>`;
  if (n < 0) return `<span class="neg">${s}</span>`;
  return s;
}

function renderTables(data) {
  if (data.perf_summary) {
    const rows = data.perf_summary.map(r =>
      `<tr>
        <td>${r["標的"]}</td>
        <td>${numColor(r["累計報酬(%)"])}</td>
        <td>${numColor(r["大盤報酬(%)"])}</td>
        <td>${numColor(r["超額績效(%)"])}</td>
      </tr>`
    ).join("");
    document.getElementById("perfTable").innerHTML =
      `<table><thead><tr><th>標的</th><th>累計報酬(%)</th><th>大盤報酬(%)</th><th>超額績效(%)</th></tr></thead><tbody>${rows}</tbody></table>`;
  }

  if (data.portfolio_rows) {
    const rows = data.portfolio_rows.map(r =>
      `<tr>
        <td>${r["標的"]}</td>
        <td>${Number(r["平均成本"]).toFixed(2)}</td>
        <td>$${fmt(r["總成本"])}</td>
        <td>${Number(r["現價"]).toFixed(2)}</td>
        <td>${Number(r["總股數"]).toFixed(3)}</td>
        <td>$${fmt(r["市值"])}</td>
        <td>${numColor(r["損益"])}</td>
        <td>${numColor(r["報酬率(%)"])}</td>
      </tr>`
    ).join("");
    document.getElementById("portfolioTable").innerHTML =
      `<table><thead><tr>
        <th>標的</th><th>平均成本</th><th>總成本</th><th>現價</th>
        <th>總股數</th><th>市值</th><th>損益</th><th>報酬率(%)</th>
      </tr></thead><tbody>${rows}</tbody></table>`;
  }

  if (data.dca_total_cost !== undefined) {
    document.getElementById("dcaInfo").innerHTML =
      `0050 定期定額投入成本：<b>$${fmt(data.dca_total_cost)}</b>；` +
      `累計買進股數：<b>${data.dca_total_shares} 股</b>`;
  }

  if (data.chip_tables) {
    const names = Object.keys(data.chip_tables);
    let html = '<div class="grid2">';
    for (const name of names) {
      const rows = data.chip_tables[name] || [];
      const trs = rows.map(r =>
        `<tr>
          <td>${r["日期"]}</td>
          <td>${numColor(r["外資買賣超"])}</td>
          <td>${numColor(r["投信買賣超"])}</td>
          <td>${numColor(r["自營商買賣超"])}</td>
          <td>${numColor(r["三大法人買賣超"])}</td>
        </tr>`
      ).join("");
      html += `<div><h3>${name}</h3>
        <div class="tbl-wrap"><table>
          <thead><tr><th>日期</th><th>外資</th><th>投信</th><th>自營商</th><th>三大法人</th></tr></thead>
          <tbody>${trs || "<tr><td colspan='5' style='text-align:center;color:#888'>無資料</td></tr>"}</tbody>
        </table></div></div>`;
    }
    html += "</div>";
    document.getElementById("chipTablesWrap").innerHTML = html;
  }

  if (data.margin_summary) {
    const rows = data.margin_summary.map(r =>
      `<tr><td>${r["標的"]}</td><td>${sigBadge(r["最新建議"])}</td><td>${r["說明"] || ""}</td></tr>`
    ).join("");
    document.getElementById("marginSummaryTable").innerHTML =
      `<table><thead><tr><th>標的</th><th>最新建議</th><th>說明</th></tr></thead><tbody>${rows}</tbody></table>`;
  }

  if (data.margin_tables) {
    let html = '<div class="grid2">';
    for (const [name, rows] of Object.entries(data.margin_tables)) {
      const trs = (rows || []).map(r =>
        `<tr>
          <td>${r["日期"]}</td>
          <td>${fmt(r["融資餘額"])}</td>
          <td>${fmt(r["融券餘額"])}</td>
          <td>${fmt(r["借券量"])}</td>
          <td>${sigBadge({"建議買":"買","建議不買":"不買","觀望":"觀望"}[r["建議"]] || r["建議"])}</td>
        </tr>`
      ).join("");
      html += `<div><h3>${name}</h3>
        <div class="tbl-wrap"><table>
          <thead><tr><th>日期</th><th>融資餘額</th><th>融券餘額</th><th>借券量</th><th>建議</th></tr></thead>
          <tbody>${trs || "<tr><td colspan='5' style='text-align:center;color:#888'>無資料</td></tr>"}</tbody>
        </table></div></div>`;
    }
    html += "</div>";
    document.getElementById("marginDetailWrap").innerHTML = html;
  }

  if (data.recommendation_rows) {
    const rows = data.recommendation_rows.map(r =>
      `<tr>
        <td>${r["標的"]}</td>
        <td>${numColor(r["近三天外資買賣超"])}</td>
        <td>${numColor(r["近三天投信買賣超"])}</td>
        <td>${numColor(r["近三天三大法人買賣超"])}</td>
        <td>${sigBadge(r["法人建議"])}</td>
        <td>${sigBadge(r["融資融券建議"])}</td>
        <td><b>${sigBadge(r["綜合建議"])}</b></td>
      </tr>`
    ).join("");
    document.getElementById("recTable").innerHTML =
      `<table><thead><tr>
        <th>標的</th><th>近三天外資</th><th>近三天投信</th><th>近三天三大法人</th>
        <th>法人建議</th><th>融資融券建議</th><th>綜合建議</th>
      </tr></thead><tbody>${rows}</tbody></table>`;
  }

  if (data.fundamental) {
    const f = data.fundamental;
    const cell = (v) => (v === null || v === undefined || v === "") ? "無資料" : v;
    const numCell = (v) => {
      if (v === null || v === undefined || v === "") return "無資料";
      const n = Number(v);
      return Number.isFinite(n) ? n.toFixed(2) : v;
    };
    let analysisHtml = `<h3>基本面分析</h3>`;
    let recHtml = `<h3>基本面買賣建議表</h3>`;

    if (f.recommendation_rows && f.recommendation_rows.length) {
      const rows = f.recommendation_rows.map(r =>
        `<tr>
          <td>${cell(r["標的"])}</td>
          <td>${cell(r["比較族群"])}</td>
          <td>${cell(r["基本面分數"])}</td>
          <td>${sigBadge(r["基本面建議"])}</td>
          <td>${cell(r["說明"])}</td>
        </tr>`
      ).join("");
      recHtml += `<div class="tbl-wrap"><table><thead><tr>
          <th>標的</th><th>比較族群</th><th>基本面分數</th><th>基本面建議</th><th>說明</th>
        </tr></thead><tbody>${rows}</tbody></table></div>`;
    } else {
      recHtml += `<div class="tbl-wrap"><table><tbody><tr><td style="text-align:center;color:#888">基本面資料不足，暫時無法產生建議。</td></tr></tbody></table></div>`;
    }

    if (f.summary_rows && f.summary_rows.length) {
      const rows = f.summary_rows.map(r =>
        `<tr>
          <td>${cell(r["標的"])}</td><td>${cell(r["比較族群"])}</td><td>${cell(r["年度"])}</td>
          <td>${cell(r["ROE排名"])}</td><td>${cell(r["ROA排名"])}</td><td>${cell(r["EPS排名"])}</td>
          <td>${cell(r["毛利率排名"])}</td><td>${cell(r["流動比率排名"])}</td><td>${cell(r["速動比率排名"])}</td>
        </tr>`
      ).join("");
      analysisHtml += `<h3 style="margin-top:18px">重點標的族群排名摘要</h3>
        <div class="tbl-wrap"><table><thead><tr>
          <th>標的</th><th>比較族群</th><th>年度</th><th>ROE排名</th><th>ROA排名</th>
          <th>EPS排名</th><th>毛利率排名</th><th>流動比率排名</th><th>速動比率排名</th>
        </tr></thead><tbody>${rows}</tbody></table></div>`;
    }

    if (f.trend_rows && f.trend_rows.length) {
      const rows = f.trend_rows.map(r =>
        `<tr>
          <td>${cell(r["標的"])}</td><td>${cell(r["年度"])}</td>
          <td>${cell(r["ROE(%)"])}</td><td>${cell(r["ROE(%)趨勢"])}</td>
          <td>${cell(r["ROA(%)"])}</td><td>${cell(r["ROA(%)趨勢"])}</td>
          <td>${cell(r["EPS"])}</td><td>${cell(r["EPS趨勢"])}</td>
          <td>${cell(r["毛利率(%)"])}</td><td>${cell(r["毛利率(%)趨勢"])}</td>
        </tr>`
      ).join("");
      analysisHtml += `<h3 style="margin-top:18px">近年獲利趨勢</h3>
        <div class="tbl-wrap"><table><thead><tr>
          <th>標的</th><th>年度</th><th>ROE(%)</th><th>ROE趨勢</th><th>ROA(%)</th><th>ROA趨勢</th>
          <th>EPS</th><th>EPS趨勢</th><th>毛利率(%)</th><th>毛利率趨勢</th>
        </tr></thead><tbody>${rows}</tbody></table></div>`;
    }

    if (f.group_tables) {
      for (const [groupName, rows] of Object.entries(f.group_tables)) {
        const trs = (rows || []).map(r =>
          `<tr>
            <td>${cell(r["標的"])}</td><td>${cell(r["年度"])}</td>
            <td>${numCell(r["ROE(%)"])}</td><td>${cell(r["ROE(%)排名"])}</td>
            <td>${numCell(r["ROA(%)"])}</td><td>${cell(r["ROA(%)排名"])}</td>
            <td>${numCell(r["EPS"])}</td><td>${cell(r["EPS排名"])}</td>
            <td>${numCell(r["毛利率(%)"])}</td><td>${cell(r["毛利率(%)排名"])}</td>
            <td>${numCell(r["流動比率"])}</td><td>${cell(r["流動比率排名"])}</td>
            <td>${numCell(r["速動比率"])}</td><td>${cell(r["速動比率排名"])}</td>
          </tr>`
        ).join("");
        analysisHtml += `<h3 style="margin-top:18px">${groupName} 族群財務比較</h3>
          <div class="tbl-wrap"><table><thead><tr>
            <th>標的</th><th>年度</th><th>ROE(%)</th><th>ROE排名</th><th>ROA(%)</th><th>ROA排名</th>
            <th>EPS</th><th>EPS排名</th><th>毛利率(%)</th><th>毛利率排名</th>
            <th>流動比率</th><th>流動比率排名</th><th>速動比率</th><th>速動比率排名</th>
          </tr></thead><tbody>${trs || "<tr><td colspan='14' style='text-align:center;color:#888'>無資料</td></tr>"}</tbody></table></div>`;
      }
    }

    document.getElementById("fundamentalAnalysisWrap").innerHTML = analysisHtml;
    document.getElementById("fundamentalRecWrap").innerHTML = recHtml;
  }
}

// Tab 切換
function switchTab(idx) {
  document.querySelectorAll(".tab").forEach((t, i) => t.classList.toggle("active", i === idx));
  document.querySelectorAll(".tab-content").forEach((c, i) => c.classList.toggle("active", i === idx));
  // 切換頁籤後重新調整圖表尺寸
  if (idx === 0 && chartPrice)  chartPrice.resize();
  if (idx === 1 && chartReturn) chartReturn.resize();
  if (idx === 2 && chartExcess) chartExcess.resize();
}

// 重新整理資料
function triggerRefresh() {
  fetch("/api/refresh", { method: "POST" })
    .then(r => r.json())
    .then(d => alert(d.message));
}

// 即時時鐘
setInterval(() => {
  document.getElementById("clock").textContent = new Date().toLocaleString("zh-TW");
}, 1000);

// 初次載入資料
loadFullData();
</script>
</body>
</html>
"""

    @app.route("/")
    def dashboard():
        return _rts(_DASHBOARD_HTML)

    print("Dashboard 啟動：http://127.0.0.1:5000")
    webbrowser.open("http://127.0.0.1:5000")
    app.run(debug=False, port=5000, threaded=True)