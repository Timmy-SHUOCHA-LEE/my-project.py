import requests
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.patches as mpatches
import matplotlib.gridspec as gridspec_mod
from matplotlib.gridspec import GridSpec
from matplotlib.widgets import Button
from datetime import datetime, timedelta
from typing import Optional, cast

# ---- 字型設定 ----
plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei", "SimHei", "Arial Unicode MS"]
plt.rcParams["axes.unicode_minus"] = False

# ---- FRED API Key ----
FRED_API_KEY = "74aa3a285c0f3cfc3e5b23800372ddf1"
FRED_URL     = "https://api.stlouisfed.org/fred/series/observations"

# ---- 資料起始日（回推三年）----
START_DATE = (datetime.today() - timedelta(days=365 * 3)).strftime("%Y-%m-%d")
# CPI 需 pct_change(12)，START_DATE 已涵蓋足夠資料
CPI_START_DATE = START_DATE

# ==========================================
# 抓取 FRED 資料
# ==========================================
def fetch_fred(series_id: str, start: Optional[str] = None) -> pd.Series:
    if start is None:
        start = START_DATE
    params = {
        "series_id": series_id,
        "api_key":   FRED_API_KEY,
        "file_type": "json",
        "observation_start": start,
    }
    resp = requests.get(FRED_URL, params=params, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    if "observations" not in data or not data["observations"]:
        raise ValueError(f"FRED 查無資料：{series_id}")
    df = pd.DataFrame(data["observations"])
    df["date"]  = pd.to_datetime(df["date"])
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    return df.dropna(subset=["value"]).set_index("date")["value"]


print("正在從 FRED 抓取資料...")
cpi_raw  = fetch_fred("CPIAUCSL", start=CPI_START_DATE); print("  CPIAUCSL ✓")
sticky   = fetch_fred("CORESTICKM159SFRBATL"); print("  CORESTICKM159SFRBATL ✓")
unrate   = fetch_fred("UNRATE");               print("  UNRATE ✓")
effr     = fetch_fred("DFF");                  print("  DFF (EFFR) ✓")
fedfunds = fetch_fred("FEDFUNDS");             print("  FEDFUNDS (Monthly) ✓")
dgs10    = fetch_fred("DGS10");                print("  DGS10 (10Y Treasury Yield) ✓")
vix      = fetch_fred("VIXCLS");               print("  VIXCLS (VIX) ✓")
oil      = fetch_fred("DCOILWTICO");           print("  DCOILWTICO (WTI Oil) ✓")
# 台幣對美金：DEXTAUS = TWD per 1 USD（數值高 = 台幣貶）
try:
    twd_raw = fetch_fred("DEXTAUS");           print("  DEXTAUS (TWD/USD) ✓")
except Exception as e:
    print(f"  DEXTAUS 失敗({e})，改用 yfinance USDTWD=X ...")
    _df_fx = cast(pd.DataFrame, yf.download("USDTWD=X", start=START_DATE, progress=False, auto_adjust=True))
    if isinstance(_df_fx.columns, pd.MultiIndex):
        _df_fx.columns = _df_fx.columns.get_level_values(0)
    twd_raw = _df_fx["Close"].dropna()
    print("  USDTWD=X (yfinance) ✓")

# 日/週資料轉月末值
effr_m  = effr.resample("ME").last().dropna()
dgs10_m = dgs10.resample("ME").last().dropna()
vix_m   = vix.resample("ME").last().dropna()
oil_m   = oil.resample("ME").last().dropna()
twd_m   = twd_raw.resample("ME").last().dropna()

# CPI 年增率 (YoY %)
cpi_yoy = cpi_raw.pct_change(12) * 100
cpi_yoy = cpi_yoy.dropna()

# ==========================================
# 趨勢判斷（近 3 個月 vs 前 3 個月均值）
# ==========================================
def trend(s: pd.Series, window: int = 3) -> str:
    s = s.dropna()
    if len(s) < window * 2:
        return "不明"
    newer = s.iloc[-window:].mean()
    older = s.iloc[-window * 2:-window].mean()
    return "上升" if newer > older else "下降"

def twd_direction(s: pd.Series, window: int = 3) -> str:
    """DEXTAUS 上升 = TWD 貶值；下降 = TWD 升值"""
    s = s.dropna()
    if len(s) < window * 2:
        return "不明"
    newer = s.iloc[-window:].mean()
    older = s.iloc[-window * 2:-window].mean()
    return "貶值" if newer > older else "升值"

cpi_trend   = trend(cpi_yoy)
ur_trend    = trend(unrate)
ff_trend    = trend(effr_m)
dgs10_trend = trend(dgs10_m)
vix_trend   = trend(vix_m)
oil_trend   = trend(oil_m)
twd_dir     = twd_direction(twd_m)

latest_cpi_yoy  = cpi_yoy.iloc[-1]
latest_ur       = unrate.iloc[-1]
latest_ff       = effr_m.iloc[-1]
latest_fedfunds = fedfunds.iloc[-1]
latest_sticky   = sticky.iloc[-1]
latest_dgs10    = dgs10_m.iloc[-1]
latest_vix      = vix_m.iloc[-1]
latest_oil      = oil_m.iloc[-1]
latest_twd      = twd_m.iloc[-1]
latest_date     = cpi_yoy.index[-1].strftime("%Y-%m")

# ==========================================
# 情境對照表 & 台股判斷
# Tuple: (id, 情況, UR, CPI, Fed利率方向, Fed動作, 美債, VIX, 油價, 台幣, 台股展望, 台股判斷, 顏色)
# ==========================================
SCENARIOS = [
    (1, "通膨↑ 景氣好", "下降↓", "上升↑", "上升↑", "升息",       "殖利率↑/債跌", "可能上升",  "偏強",    "貶值",  "估值壓力大，景氣股撐住", "中性偏利空",  "#FF9800"),
    (2, "通膨↑ 景氣差", "上升↑", "上升↑", "高檔",  "維持高位",   "殖利率難降",   "上升",      "可能偏高", "貶值",  "最不利，容易下跌",       "利空",        "#E53935"),
    (3, "通膨↓ 景氣好", "低/穩", "下降↓", "見頂↓", "準備降息",   "殖利率↓/債漲", "下降",      "偏弱/穩", "升值",  "股市偏多，最有利",       "利多",        "#43A047"),
    (4, "通膨↓ 景氣差", "上升↑", "下降↓", "下降↓", "降息救景氣", "殖利率↓/債漲", "先升後降",  "偏弱",    "不一定","初期跌，落底後反彈",     "短利空長利多","#1565C0"),
]

if   cpi_trend == "上升" and ur_trend == "下降": cur_id = 1
elif cpi_trend == "上升" and ur_trend == "上升":  cur_id = 2
elif cpi_trend == "下降" and ur_trend == "下降":  cur_id = 3
else:                                              cur_id = 4

cur = next(s for s in SCENARIOS if s[0] == cur_id)
(_, cur_situation, _, _, _, cur_fed_act, cur_bond_s, _, cur_oil_s,
 cur_twd_s, cur_detail, cur_verdict, cur_color) = cur

# ==========================================
# ==========================================
# 繪圖工具函式
# ==========================================
def plot_series(ax, s: pd.Series, title: str, color: str, tr: str,
                latest_str: str, tr_color: Optional[str] = None):
    s = s.dropna()
    ax.plot(s.index, s.values, color=color, linewidth=1.8)
    ax.fill_between(s.index, s.values, alpha=0.10, color=color)
    cutoff = s.index[-1] - pd.DateOffset(months=6)
    mask = s.index >= cutoff
    ax.fill_between(s[mask].index, s[mask].values, alpha=0.30, color=color)
    ax.set_title(title, fontsize=10, fontweight="bold")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
    ax.xaxis.set_major_locator(mdates.MonthLocator(bymonth=[1, 7]))
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha="right", fontsize=8)
    ax.grid(True, alpha=0.22, linestyle="--")
    ax.set_facecolor("white")
    if tr_color is None:
        tr_color = "#E53935" if tr == "上升" else "#43A047"
    ax.text(0.03, 0.95, f"趨勢：{tr}", transform=ax.transAxes,
            fontsize=9, color=tr_color, fontweight="bold", va="top",
            bbox=dict(boxstyle="round,pad=0.3", facecolor="white", alpha=0.85))
    ax.text(0.97, 0.05, f"最新：{latest_str}", transform=ax.transAxes,
            fontsize=9, color=color, fontweight="bold", ha="right", va="bottom")

# ==========================================
# 主圖：6 列 × 2 欄
# ==========================================
fig = plt.figure(figsize=(18, 42))
fig.patch.set_facecolor("#f0f2f5")
gs  = GridSpec(6, 2, figure=fig, hspace=0.80, wspace=0.40)

fig.suptitle(
    f"U.S. Economic Indicators ── 台股情勢分析 (截至 {latest_date})",
    fontsize=15, fontweight="bold"
)

# ---- Row 0 ----
plot_series(fig.add_subplot(gs[0, 0]), cpi_yoy,
            "CPI 年增率 YoY (%)", "#E05C5C", cpi_trend, f"{latest_cpi_yoy:.2f}%")
plot_series(fig.add_subplot(gs[0, 1]), sticky,
            "Sticky Price CPI (less Food & Energy)", "#9C27B0",
            trend(sticky), f"{latest_sticky:.2f}")

# ---- Row 1 ----
plot_series(fig.add_subplot(gs[1, 0]), unrate,
            "Unemployment Rate (%)", "#4A90D9", ur_trend, f"{latest_ur:.1f}%")
plot_series(fig.add_subplot(gs[1, 1]), effr_m,
            "Effective Federal Funds Rate (EFFR %)", "#E65100",
            ff_trend, f"{latest_ff:.2f}%")

# ---- Row 2: FEDFUNDS 全寬 ----
plot_series(fig.add_subplot(gs[2, :]), fedfunds,
            "Federal Funds Effective Rate — FEDFUNDS (Monthly Average, %)",
            "#5C6BC0", trend(fedfunds), f"{latest_fedfunds:.2f}%")

# ---- Row 3: 美債殖利率 | VIX ----
plot_series(fig.add_subplot(gs[3, 0]), dgs10_m,
            "10-Year Treasury Yield DGS10 (%)", "#00897B",
            dgs10_trend, f"{latest_dgs10:.2f}%")
plot_series(fig.add_subplot(gs[3, 1]), vix_m,
            "CBOE VIX 波動率指數 (VIXCLS)", "#D81B60",
            vix_trend, f"{latest_vix:.2f}")

# ---- Row 4: WTI 油價 | 台幣對美金 ----
plot_series(fig.add_subplot(gs[4, 0]), oil_m,
            "WTI 原油價格 ($/桶) DCOILWTICO", "#6D4C41",
            oil_trend, f"${latest_oil:.2f}")
# TWD: DEXTAUS 上升 = 台幣貶值（利空）; 下降 = 台幣升值（利多）
twd_tc = "#E53935" if twd_dir == "貶值" else "#43A047"
plot_series(fig.add_subplot(gs[4, 1]), twd_m,
            "台幣對美金匯率 TWD/USD（數值高=台幣貶）", "#00838F",
            twd_dir, f"{latest_twd:.3f}", tr_color=twd_tc)

# ==========================================
# Row 5: 情境分析面板（全寬）
# ==========================================
ax_b = fig.add_subplot(gs[5, :])
ax_b.set_facecolor("#fafafa")
ax_b.set_xlim(-1, 16)
ax_b.set_ylim(0, 10)
ax_b.axis("off")
ax_b.set_title("核心關係總表 與 台股情勢判斷", fontsize=12, fontweight="bold", pad=10)

# ---- 表格欄位（10 欄，x 軸 0~10） ----
col_headers = ["情況（景氣）", "失業率", "CPI", "Fed利率", "Fed動作",
               "美債", "VIX", "油價", "台幣", "台股展望"]
col_xs     = [0.02, 1.45, 2.52, 3.42, 4.32, 5.30, 6.07, 6.79, 7.54, 8.28]
col_widths = [1.40, 1.04, 0.87, 0.87, 0.95, 0.74, 0.69, 0.72, 0.71, 1.67]
row_top = 9.3
row_h   = 1.55

# 表頭
for hdr, x, w in zip(col_headers, col_xs, col_widths):
    rect = mpatches.FancyBboxPatch((x, row_top - 0.60), w, 0.65,
                                   boxstyle="square,pad=0.02",
                                   facecolor="#37474F", edgecolor="white", lw=0.5)
    ax_b.add_patch(rect)
    ax_b.text(x + w / 2, row_top - 0.28, hdr, ha="center", va="center",
              fontsize=7.5, color="white", fontweight="bold")

# 各情境列
row_colors_bg = ["#FFF3E0", "#FFEBEE", "#E8F5E9", "#E3F2FD"]
for i, sc in enumerate(SCENARIOS):
    sid, sit, ur_s, cpi_s, ff_s, fed_act, bond_s, vix_s, oil_s, twd_s, stock_s, verdict_s, sc_color = sc
    is_cur = (sid == cur_id)
    row_y  = row_top - 0.60 - (i + 1) * row_h
    bg     = sc_color + "55" if is_cur else row_colors_bg[i]

    full_rect = mpatches.FancyBboxPatch((0.02, row_y), 9.93, row_h - 0.08,
                                         boxstyle="square,pad=0.02",
                                         facecolor=bg,
                                         edgecolor="#CCCCCC",
                                         lw=0.4)
    ax_b.add_patch(full_rect)

    if is_cur:
        ax_b.text(-0.08, row_y + (row_h - 0.08) / 2, "現在 ►",
                  ha="right", va="center", fontsize=8.5, color=sc_color, fontweight="bold")

    row_data = [sit, ur_s, cpi_s, ff_s, fed_act, bond_s, vix_s, oil_s, twd_s, stock_s]
    for val, x, w in zip(row_data, col_xs, col_widths):
        ax_b.text(x + w / 2, row_y + (row_h - 0.08) / 2, val,
                  ha="center", va="center", fontsize=7.5,
                  color="#B71C1C" if is_cur else "#333333",
                  fontweight="bold" if is_cur else "normal")

# ---- 右側總評框 (x: 10.25 ~ 15.8) ----
vx, vy, vw, vh = 10.25, 1.2, 5.55, 8.6
summary_bg = mpatches.FancyBboxPatch((vx, vy), vw, vh,
                                      boxstyle="round,pad=0.15",
                                      facecolor=cur_color + "22",
                                      edgecolor=cur_color, lw=2.5)
ax_b.add_patch(summary_bg)

ax_b.text(vx + vw / 2, vy + vh - 0.5,
          "台股情勢判斷", ha="center", va="top",
          fontsize=11, fontweight="bold", color="#333333")

ax_b.text(vx + vw / 2, vy + vh - 1.3,
          f"目前情境：{cur_situation}",
          ha="center", va="top", fontsize=10, color=cur_color, fontweight="bold")

vbox = mpatches.FancyBboxPatch((vx + 0.3, vy + vh - 3.2), vw - 0.6, 1.3,
                                boxstyle="round,pad=0.1",
                                facecolor=cur_color, edgecolor=cur_color, lw=1)
ax_b.add_patch(vbox)
ax_b.text(vx + vw / 2, vy + vh - 2.55,
          f"台股：{cur_verdict}",
          ha="center", va="center", fontsize=13, fontweight="bold", color="white")

metrics = [
    ("CPI YoY 趨勢",   f"{cpi_trend}  ({latest_cpi_yoy:.2f}%)"),
    ("失業率趨勢",      f"{ur_trend}  ({latest_ur:.1f}%)"),
    ("Fed Funds Rate", f"{ff_trend}  ({latest_ff:.2f}%)"),
    ("10Y 殖利率",      f"{dgs10_trend}  ({latest_dgs10:.2f}%)"),
    ("VIX 波動率",      f"{vix_trend}  ({latest_vix:.2f})"),
    ("WTI 油價",        f"{oil_trend}  (${latest_oil:.2f})"),
    ("台幣方向",        f"{twd_dir}  ({latest_twd:.3f})"),
    ("預期 Fed 動作",   cur_fed_act),
    ("台股展望",        cur_detail),
]
for k, (lbl, val) in enumerate(metrics):
    y_pos = vy + vh - 3.8 - k * 0.52
    ax_b.text(vx + 0.3,  y_pos, f"• {lbl}：", ha="left", va="center",
              fontsize=8, color="#555555")
    ax_b.text(vx + 2.55, y_pos, val, ha="left", va="center",
              fontsize=8, color=cur_color, fontweight="bold")

fig.text(0.5, 0.005,
         "Source: Federal Reserve Bank of St. Louis (FRED) | CBOE | EIA",
         ha="center", fontsize=8, color="#888888")

# ==========================================
# 台灣股市指數 — 日K走勢 + RSI + MA（獨立圖表）
# ==========================================
_end   = datetime.today().strftime("%Y-%m-%d")
_start = START_DATE

df_twii  = cast(pd.DataFrame, yf.download("^TWII",  start=_start, end=_end, progress=False, auto_adjust=True))
df_twoii = cast(pd.DataFrame, yf.download("^TWOII", start=_start, end=_end, progress=False, auto_adjust=True))
if isinstance(df_twii.columns,  pd.MultiIndex):
    df_twii.columns  = df_twii.columns.get_level_values(0)
if isinstance(df_twoii.columns, pd.MultiIndex):
    df_twoii.columns = df_twoii.columns.get_level_values(0)


def _plot_kline(ax, df: pd.DataFrame, title: str,
                up_color="#e85454", down_color="#54a0e8"):
    ax.set_title(title, fontsize=13)
    ax.set_ylabel("指數", fontsize=11)
    ax.grid(True, alpha=0.25)
    ax.set_facecolor("#f8f9fa")
    for i, (idx, row) in enumerate(df.iterrows()):
        o, h, l, c = row["Open"], row["High"], row["Low"], row["Close"]
        color = up_color if c >= o else down_color
        ax.plot([i, i], [l, h], color=color, linewidth=0.8)
        body_bottom = min(o, c)
        body_height = abs(c - o) if abs(c - o) > 0 else (h - l) * 0.001
        rect = mpatches.FancyBboxPatch(
            (i - 0.35, body_bottom), 0.7, body_height,
            boxstyle="square,pad=0", linewidth=0,
            facecolor=color, edgecolor=color
        )
        ax.add_patch(rect)
    tick_positions, tick_labels, prev_month = [], [], None
    for i, idx in enumerate(df.index):
        if idx.month != prev_month:
            if idx.month in (1, 4, 7, 10):   # 每季顯示一次，避免擠在一起
                tick_positions.append(i)
                tick_labels.append(idx.strftime("%Y-%m"))
            prev_month = idx.month
    ax.set_xticks(tick_positions)
    ax.set_xticklabels(tick_labels, rotation=45, ha="right", fontsize=9)
    ax.set_xlim(-1, len(df))
    margin = (df["High"].max() - df["Low"].min()) * 0.03
    ax.set_ylim(df["Low"].min() - margin, df["High"].max() + margin)


def _calc_rsi(df, period):
    close    = df["Close"]
    delta    = close.diff()
    gain     = delta.clip(lower=0)
    loss     = (-delta).clip(lower=0)
    avg_gain = gain.rolling(window=period).mean()
    avg_loss = loss.rolling(window=period).mean()
    return 100 * avg_gain / (avg_gain + avg_loss)


def _calc_kdj(df, period=9):
    """RSV=(C-Ln)/(Hn-Ln)×100%, K=2/3*K_prev+1/3*RSV, D=2/3*D_prev+1/3*K, J=3K-2D"""
    low_n  = df["Low"].rolling(period).min()
    high_n = df["High"].rolling(period).max()
    denom  = (high_n - low_n).replace(0, 1)
    rsv    = ((df["Close"] - low_n) / denom * 100).fillna(50).clip(0, 100)
    k      = rsv.ewm(alpha=1/3, adjust=False).mean()
    d      = k.ewm(alpha=1/3, adjust=False).mean()
    j      = 3 * k - 2 * d
    return k, d, j


def _plot_rsi(ax, rsi7, rsi14, n):
    x = list(range(n))
    line7,  = ax.plot(x, rsi7.values,  color="#ffa500", linewidth=1.3, label="RSI 7")
    line14, = ax.plot(x, rsi14.values, color="#9370db", linewidth=1.3, label="RSI 14")
    ax.axhline(80, color="red",   linestyle="--", linewidth=0.8, alpha=0.7)
    ax.axhline(20, color="green", linestyle="--", linewidth=0.8, alpha=0.7)
    ax.fill_between(x, 80, 100, alpha=0.05, color="red")
    ax.fill_between(x, 0,  20,  alpha=0.05, color="green")
    ax.set_ylabel("RSI", fontsize=10)
    ax.set_ylim(0, 100)
    ax.set_yticks([0, 20, 50, 80, 100])
    ax.legend(loc="upper left", fontsize=9)
    ax.grid(True, alpha=0.25)
    ax.set_facecolor("#f8f9fa")
    ax.set_xlim(-1, n)
    ax.tick_params(axis="x", labelbottom=False)
    return line7, line14


def _plot_kdj(ax, k, d, j, n):
    """KDJ 子圖：K線(橘)、D線(紫)、J線(紅虛)；80/20 警戒線"""
    x = list(range(n))
    lk, = ax.plot(x, k.values, color="#ffa500", linewidth=1.3, label="K")
    ld, = ax.plot(x, d.values, color="#9370db", linewidth=1.3, label="D")
    lj, = ax.plot(x, j.values, color="#e85454", linewidth=1.0, label="J", linestyle="--")
    ax.axhline(80, color="red",   linestyle="--", linewidth=0.8, alpha=0.7)
    ax.axhline(20, color="green", linestyle="--", linewidth=0.8, alpha=0.7)
    ax.fill_between(x, 80, 100, alpha=0.05, color="red")
    ax.fill_between(x, 0,  20,  alpha=0.05, color="green")
    ax.set_title("KDJ (9, 3, 3)", fontsize=9, fontweight="bold")
    ax.set_ylabel("KDJ", fontsize=10)
    ax.set_ylim(-20, 120)
    ax.set_yticks([0, 20, 50, 80, 100])
    ax.legend(loc="upper left", fontsize=9)
    ax.grid(True, alpha=0.25)
    ax.set_facecolor("#f8f9fa")
    ax.set_xlim(-1, n)
    ax.tick_params(axis="x", labelbottom=False)
    return lk, ld, lj


# ---- RSI 計算 ----
rsi7_twii   = _calc_rsi(df_twii,   7)
rsi14_twii  = _calc_rsi(df_twii,  14)
rsi7_twoii  = _calc_rsi(df_twoii,  7)
rsi14_twoii = _calc_rsi(df_twoii, 14)

# ---- KDJ 計算 ----
kdj_k_twii,  kdj_d_twii,  kdj_j_twii  = _calc_kdj(df_twii)
kdj_k_twoii, kdj_d_twoii, kdj_j_twoii = _calc_kdj(df_twoii)

# ---- 預算 MA（條件判斷 & 乖離率）----
_ma5_t   = df_twii["Close"].rolling(5).mean()
_ma20_t  = df_twii["Close"].rolling(20).mean()
_ma60_t  = df_twii["Close"].rolling(60).mean()
_ma120_t = df_twii["Close"].rolling(120).mean()

_ma5_o   = df_twoii["Close"].rolling(5).mean()
_ma20_o  = df_twoii["Close"].rolling(20).mean()
_ma60_o  = df_twoii["Close"].rolling(60).mean()
_ma120_o = df_twoii["Close"].rolling(120).mean()

# 乖離率 = (收盤 - MA20) / MA20 × 100%
bias_twii  = (df_twii["Close"]  - _ma20_t) / _ma20_t  * 100
bias_twoii = (df_twoii["Close"] - _ma20_o) / _ma20_o  * 100


def _get_signals(close, ma5, ma20, ma60, ma120, rsi7, rsi14, verdict=""):
    """回傳 [(訊息文字, 顏色), ...] 信號清單（3 個條件）"""
    sig = []
    c = close.dropna().iloc[-1]
    # 條件 1 & 2：加權指數與櫃買指數，跌破月線或是月線與季線之間 → 提醒我，可考慮進場
    if pd.notna(ma20.iloc[-1]) and pd.notna(ma60.iloc[-1]) and ma20.iloc[-1] >= c >= ma60.iloc[-1]:
        sig.append(("▼ 介於月線(MA20)與季線(MA60)之間  → 可考慮進場", "#1B5E20"))
    if pd.notna(ma120.iloc[-1]) and c < ma120.iloc[-1]:
        sig.append(("▼ 跌破半年線(MA120)              → 可考慮進場", "#0D47A1"))
    if pd.notna(ma20.iloc[-1]) and c < ma20.iloc[-1]:
        sig.append(("⚠ 跌破月線(MA20)                → 注意",        "#BF360C"))
    # 條件 4（綜合）：介於月線與季線之間 + 台股情勢利空 → 警示可能出場
    is_bearish = "利空" in verdict
    if (pd.notna(ma20.iloc[-1]) and pd.notna(ma60.iloc[-1])
            and ma20.iloc[-1] >= c >= ma60.iloc[-1]
            and is_bearish):
        sig.append((f"🚨 月線~季線間 ✕ 情勢{verdict} → 注意，可能需出場", "#7B1FA2"))
    # 條件 3：RSI > 80 → 短線過熱、勿追高
    _r7  = rsi7.dropna();  _r14 = rsi14.dropna()
    if not _r7.empty  and _r7.iloc[-1]  > 80:
        sig.append(("● RSI7 > 80  短線過熱・勿追高",     "#C62828"))
    if not _r14.empty and _r14.iloc[-1] > 80:
        sig.append(("● RSI14 > 80 短線過熱・勿追高",     "#C62828"))
    return sig


sigs_twii  = _get_signals(df_twii["Close"],  _ma5_t, _ma20_t, _ma60_t, _ma120_t,
                           rsi7_twii,  rsi14_twii,  verdict=cur_verdict)
sigs_twoii = _get_signals(df_twoii["Close"], _ma5_o, _ma20_o, _ma60_o, _ma120_o,
                           rsi7_twoii, rsi14_twoii, verdict=cur_verdict)


def _plot_bias(ax, bias, n, tick_pos=None, tick_lbl=None):
    """乖離率長條圖（MA20）：紅=正乖離、藍=負乖離；±5% 虛線警戒"""
    x = list(range(n))
    vals = bias.values
    bar_colors = ["#e85454" if (not pd.isna(v) and v >= 0) else "#54a0e8" for v in vals]
    ax.bar(x, vals, color=bar_colors, alpha=0.75, width=0.8)
    ax.axhline(0,   color="#333333", linewidth=0.9)
    ax.axhline( 5,  color="#E53935", linestyle="--", linewidth=0.9, alpha=0.7)
    ax.axhline(-5,  color="#43A047", linestyle="--", linewidth=0.9, alpha=0.7)
    ax.set_title("乖離率 (MA20, %)", fontsize=10, fontweight="bold")
    ax.set_ylabel("乖離率(%)", fontsize=9)
    ax.grid(True, alpha=0.20, linestyle="--")
    ax.set_facecolor("#f8f9fa")
    ax.set_xlim(-1, n)
    latest = bias.dropna().iloc[-1] if not bias.dropna().empty else float("nan")
    c_lbl = "#E53935" if latest >= 0 else "#43A047"
    ax.text(0.97, 0.95, f"最新：{latest:.2f}%", transform=ax.transAxes,
            fontsize=9, color=c_lbl, fontweight="bold", ha="right", va="top")
    if tick_pos is not None:
        ax.set_xticks(tick_pos)
        ax.set_xticklabels(tick_lbl, rotation=45, ha="right", fontsize=9)
        ax.set_xlabel("日期", fontsize=11)
    else:
        ax.tick_params(axis="x", labelbottom=False)


# ---- 畫布與子圖配置（8 列：加權K、加權KDJ、加權RSI、加權乖離率、上櫃K、上櫃KDJ、上櫃RSI、上櫃乖離率）----
fig_tw = plt.figure(figsize=(16, 52))
fig_tw.suptitle("台灣股市指數 — 日K走勢 + KDJ + RSI + MA + 乖離率",
                fontsize=16, fontweight="bold", y=0.997)

gs_tw = gridspec_mod.GridSpec(
    8, 1,
    height_ratios=[2.5, 1, 1, 1, 2.5, 1, 1, 1],
    top=0.962, bottom=0.05,
    hspace=0.35
)

ax_tw1      = fig_tw.add_subplot(gs_tw[0])
ax_tw_kdj1  = fig_tw.add_subplot(gs_tw[1])
ax_tw_rsi1  = fig_tw.add_subplot(gs_tw[2])
ax_tw_bias1 = fig_tw.add_subplot(gs_tw[3])
ax_tw2      = fig_tw.add_subplot(gs_tw[4])
ax_tw_kdj2  = fig_tw.add_subplot(gs_tw[5])
ax_tw_rsi2  = fig_tw.add_subplot(gs_tw[6])
ax_tw_bias2 = fig_tw.add_subplot(gs_tw[7])

_plot_kline(ax_tw1, df_twii,  "加權指數（^TWII）")
# ---- 加權指數：5 大條件訊號標注 ----
for _j, (_msg, _col) in enumerate(sigs_twii):
    ax_tw1.text(0.01, 0.97 - _j * 0.075, _msg, transform=ax_tw1.transAxes,
                fontsize=8.5, color=_col, fontweight="bold", va="top",
                bbox=dict(boxstyle="round,pad=0.25", facecolor="white",
                          alpha=0.85, edgecolor=_col))
_plot_kdj(ax_tw_kdj1, kdj_k_twii, kdj_d_twii, kdj_j_twii, len(df_twii))
line7_1, line14_1 = _plot_rsi(ax_tw_rsi1, rsi7_twii,  rsi14_twii,  len(df_twii))
_plot_bias(ax_tw_bias1, bias_twii, len(df_twii))

_plot_kline(ax_tw2, df_twoii, "上櫃指數（^TWOII）")
# ---- 上櫃指數：5 大條件訊號標注 ----
for _j, (_msg, _col) in enumerate(sigs_twoii):
    ax_tw2.text(0.01, 0.97 - _j * 0.075, _msg, transform=ax_tw2.transAxes,
                fontsize=8.5, color=_col, fontweight="bold", va="top",
                bbox=dict(boxstyle="round,pad=0.25", facecolor="white",
                          alpha=0.85, edgecolor=_col))
_plot_kdj(ax_tw_kdj2, kdj_k_twoii, kdj_d_twoii, kdj_j_twoii, len(df_twoii))
line7_2, line14_2 = _plot_rsi(ax_tw_rsi2, rsi7_twoii, rsi14_twoii, len(df_twoii))

# 最底層乖離率套用日期刻度
_tick_pos, _tick_lbl, _prev_m = [], [], None
for _i, _idx in enumerate(df_twoii.index):
    if _idx.month != _prev_m:
        if _idx.month in (1, 4, 7, 10):   # 每季顯示一次
            _tick_pos.append(_i)
            _tick_lbl.append(_idx.strftime("%Y-%m"))
        _prev_m = _idx.month
_plot_bias(ax_tw_bias2, bias_twoii, len(df_twoii),
           tick_pos=_tick_pos, tick_lbl=_tick_lbl)

_rsi_lines = [line7_1, line14_1, line7_2, line14_2]

# ---- 移動平均線（MA5 / MA20 / MA60 / MA120 / MA240）----
MA_CONFIG = {
    "MA5":   {"period":   5, "color": "#f5c518"},
    "MA20":  {"period":  20, "color": "#26c6da"},
    "MA60":  {"period":  60, "color": "#ef5350"},
    "MA120": {"period": 120, "color": "#42a5f5"},
    "MA240": {"period": 240, "color": "#ab47bc"},
}

ma_lines = {}
for _key, _cfg in MA_CONFIG.items():
    _p, _c = int(_cfg["period"]), str(_cfg["color"])
    _ma_t = df_twii["Close"].rolling(_p).mean()
    _ma_o = df_twoii["Close"].rolling(_p).mean()
    _lt, = ax_tw1.plot(range(len(df_twii)),  _ma_t.to_numpy(), color=_c,
                       linewidth=1.0, linestyle="-", label=_key, alpha=0.85)
    _lo, = ax_tw2.plot(range(len(df_twoii)), _ma_o.to_numpy(), color=_c,
                       linewidth=1.0, linestyle="-", label=_key, alpha=0.85)
    ma_lines[_key] = [_lt, _lo]

ax_tw1.legend(loc="upper left", fontsize=8, ncol=2)
ax_tw2.legend(loc="upper left", fontsize=8, ncol=2)

# ==========================================
# 輸出靜態 HTML 網頁
# ==========================================
import io, base64, os

def _fig_to_base64(figure):
    buf = io.BytesIO()
    figure.savefig(buf, format="png", dpi=110, bbox_inches="tight",
                   facecolor=figure.get_facecolor())
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")

print("正在產生 HTML 網頁，請稍候...")
img1_b64 = _fig_to_base64(fig)
img2_b64 = _fig_to_base64(fig_tw)
plt.close("all")

_html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>經濟指標 & 台股分析</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    background: #0f172a;
    font-family: "Microsoft JhengHei", "Segoe UI", sans-serif;
    color: #e2e8f0;
    padding: 24px 32px 48px;
  }}
  h1 {{
    text-align: center;
    font-size: 24px;
    font-weight: 700;
    color: #93c5fd;
    letter-spacing: 3px;
    margin-bottom: 32px;
    padding-bottom: 16px;
    border-bottom: 2px solid #1e40af;
  }}
  .card {{
    background: #1e293b;
    border: 1px solid #334155;
    border-radius: 14px;
    padding: 20px 24px 24px;
    margin-bottom: 48px;
    box-shadow: 0 8px 32px rgba(0,0,0,0.5);
  }}
  .card h2 {{
    font-size: 15px;
    color: #60a5fa;
    margin-bottom: 18px;
    padding-bottom: 10px;
    border-bottom: 1px solid #334155;
    letter-spacing: 1px;
  }}
  .card img {{
    width: 100%;
    height: auto;
    display: block;
    border-radius: 8px;
  }}
  .footer {{
    text-align: center;
    font-size: 12px;
    color: #64748b;
    padding-top: 8px;
  }}
</style>
</head>
<body>
<h1>📊 台股情勢分析 &amp; 美國經濟指標</h1>
<div class="card">
  <h2>🇹🇼 台灣股市指數 ── 日K走勢 + KDJ + RSI + MA + 乖離率</h2>
  <img src="data:image/png;base64,{img2_b64}" alt="台灣股市指數">
</div>
<div class="card">
  <h2>🇺🇸 U.S. Economic Indicators ── 台股情勢判斷</h2>
  <img src="data:image/png;base64,{img1_b64}" alt="美國經濟指標">
</div>
<div class="footer">
  Source: Federal Reserve Bank of St. Louis (FRED) | CBOE | EIA | Yahoo Finance　｜　產生時間：{datetime.today().strftime("%Y-%m-%d %H:%M")}
</div>
</body>
</html>"""

_out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "economic_dashboard.html")
with open(_out, "w", encoding="utf-8") as _f:
    _f.write(_html)
print(f"✅ HTML 已產生：{_out}")

import webbrowser
webbrowser.open(_out)