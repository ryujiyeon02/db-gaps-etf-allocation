"""README용 차트: 계획 → 실제 운용 기록, ETF 구성종목 look-through, 계획 단계 백테스트.

실행: python src/build_portfolio_charts.py
입력: output/tables/realized_track.csv (월별 운용보고서 제출본의 비율 수치), output/tables/etf_*.csv,
      output/tables/backtest_daily_nav_net.csv, input/GAPS투자대회_포트폴리오_디비갭스투자대회.xlsx (6월 일별 평가액, 비공개)
"""
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
TABLES = ROOT / "output" / "tables"
CHARTS = ROOT / "output" / "charts"
CHARTS.mkdir(parents=True, exist_ok=True)

PORT, BENCH, ALT, ACCENT, AMBER = "#982A2F", "#8C8C8C", "#2F6DA3", "#1F9A6E", "#B7791F"
INK, MUTED = "#222222", "#666666"
for font_file in Path.home().glob("Library/Fonts/Pretendard-*.otf"):
    fm.fontManager.addfont(str(font_file))
plt.rcParams.update({
    "font.family": ["Pretendard", "AppleGothic"], "axes.unicode_minus": False,
    "axes.edgecolor": "#BBBBBB", "axes.titleweight": "bold", "axes.titlesize": 12,
    "xtick.color": MUTED, "ytick.color": MUTED, "axes.grid": True, "grid.color": "#E6E6E6",
    "axes.axisbelow": True, "axes.spines.top": False, "axes.spines.right": False,
    "savefig.dpi": 160, "savefig.bbox": "tight", "figure.facecolor": "white",
})

# 1. 실제 운용 경로 (원금 10억 = 100). 6월은 일별 평가액, 7~8월은 운용보고서 제출본의 월말·저점 수치.
track = pd.read_csv(TABLES / "realized_track.csv", parse_dates=["date"])
june_file = ROOT / "input" / "GAPS투자대회_포트폴리오_디비갭스투자대회.xlsx"
if june_file.exists():
    raw = pd.read_excel(june_file, header=None)
    june = pd.Series(raw.iloc[1, 1:].astype(float).values, index=pd.to_datetime(raw.iloc[0, 1:].values)).sort_index()
    june = (june / 1_000_000_000 * 100).rename("portfolio_index")
    june.to_csv(TABLES / "realized_june_daily_index.csv", encoding="utf-8-sig", index_label="date")
else:
    june = pd.read_csv(TABLES / "realized_june_daily_index.csv", index_col=0, parse_dates=True)["portfolio_index"]
june = june.loc[:"2026-06-30"]
port_path = pd.concat([pd.Series({pd.Timestamp("2026-05-29"): 100.0}), june,
                       track.dropna(subset=["portfolio_index"]).set_index("date")["portfolio_index"].loc["2026-07-01":]])
port_path = port_path[~port_path.index.duplicated(keep="last")].sort_index()
k200 = track.dropna(subset=["kospi200_index"]).set_index("date")["kospi200_index"]

fig = plt.figure(figsize=(15, 8.2))
gs = fig.add_gridspec(2, 2, height_ratios=[1.3, 1], width_ratios=[1.5, 1], hspace=0.5, wspace=0.2)
ax = fig.add_subplot(gs[0, :])
daily = port_path.loc[:"2026-06-30"]
sparse = port_path.loc["2026-06-30":]
ax.plot(daily.index, daily.values, color=PORT, linewidth=2.2, label="팀 포트폴리오 (원금 10억 = 100, 6월은 일별)")
ax.plot(sparse.index, sparse.values, color=PORT, linewidth=1.6, linestyle=":", marker="o", markersize=5,
        label="7~8월: 보고서에 기록된 저점·월말 값만 연결")
ax.plot(k200.index, k200.values, color=BENCH, linewidth=1.8, linestyle="--", marker="o", markersize=5,
        label="KOSPI200 (5/29 = 100, 월말 기준)")
for _, r in track.dropna(subset=["event"]).iterrows():
    y = port_path.asof(r["date"])
    ax.annotate(r["event"], (r["date"], y), xytext=(0, 28 if r.get("label_up", 1) else -38), textcoords="offset points",
                ha="center", fontsize=9, color=INK, arrowprops=dict(arrowstyle="-", color=MUTED, linewidth=0.8))
for d, v in port_path.loc[["2026-06-30", "2026-07-31", "2026-08-31"]].items():
    ax.text(d, v + 1.6, f"{v - 100:+.1f}%", ha="center", fontsize=9, color=PORT, fontweight="bold")
for d, v in k200.loc[["2026-07-31", "2026-08-31"]].items():
    ax.text(d, v - 3.4, f"{v - 100:+.1f}%", ha="center", fontsize=9, color=MUTED)
ax.set_ylim(68, 108)
ax.set_ylabel("지수 (시작 = 100)")
ax.set_title("A. 실제 운용 3개월 (2026-06-01 ~ 08-31): 7월 급락은 방어했고, 8월 반등은 놓쳤다", loc="left")
ax.legend(frameon=False, loc="lower left")
ax.xaxis.set_major_formatter(mdates.DateFormatter("%m/%d"))

ax2 = fig.add_subplot(gs[1, 0])
w = track.dropna(subset=["risk_weight_pct"]).set_index("date")["risk_weight_pct"]
ax2.axhspan(65, 70, color="#F3E3E4", linewidth=0)
ax2.axhline(70, color="#999999", linewidth=0.8, linestyle="--")
ax2.text(pd.Timestamp("2026-08-20"), 70.8, "규정 상한 70%", fontsize=8.5, color=MUTED)
ax2.text(pd.Timestamp("2026-06-02"), 66, "계획 초기 65% (버퍼 5%p)", fontsize=8.5, color=PORT)
ax2.step(w.index, w.values, where="post", color=PORT, linewidth=2)
ax2.scatter(w.index, w.values, color=PORT, s=30, zorder=3)
for d, v in w.items():
    ax2.text(d, v - 3.3, f"{v:.1f}", ha="center", fontsize=8.5, color=INK)
ax2.set_ylim(35, 74)
ax2.set_ylabel("위험자산 비중 (%)")
ax2.set_title("B. 위험자산 비중: 계획 65% → 실제 42~48%", loc="left")
ax2.xaxis.set_major_formatter(mdates.DateFormatter("%m/%d"))

ax3 = fig.add_subplot(gs[1, 1])
months = ["6월", "7월", "8월"]
pm = track.set_index("month_label").loc[months]
x = np.arange(3)
ax3.bar(x - 0.19, pm["portfolio_month_ret"], width=0.36, color=PORT, label="팀 포트폴리오")
ax3.bar(x + 0.19, pm["kospi200_month_ret"], width=0.36, color=BENCH, label="KOSPI200")
for xi, (a, b) in enumerate(zip(pm["portfolio_month_ret"], pm["kospi200_month_ret"])):
    ax3.text(xi - 0.19, a + (0.6 if a >= 0 else -0.6), f"{a:+.1f}", ha="center", va="bottom" if a >= 0 else "top", fontsize=9)
    ax3.text(xi + 0.19, b + (0.6 if b >= 0 else -0.6), f"{b:+.1f}", ha="center", va="bottom" if b >= 0 else "top", fontsize=9, color=MUTED)
ax3.axhline(0, color="#999999", linewidth=0.8)
ax3.set_xticks(x)
ax3.set_xticklabels(months)
ax3.set_ylim(-28, 8)
ax3.set_ylabel("월간 수익률 (%)")
ax3.set_title("C. 월간 수익률", loc="left")
ax3.legend(frameon=False, loc="lower right", fontsize=9)
ax3.grid(axis="x", visible=False)
fig.suptitle("DB GAPS 투자대회: 계획한 규칙이 실제 3개월 동안 어떻게 작동했나", x=0.01, ha="left", fontsize=15, fontweight="bold")
fig.text(0.01, 0.0, "출처: 팀 월간 운용보고서 제출본(6·7·8월), 6월 일별 평가액. KOSPI200 6월은 보유 ETF(TIGER 200) 수익률로 근사. "
         "평가 수익률은 대회 시스템 체결가·평가 기준과 다를 수 있음.", fontsize=9, color=MUTED)
fig.savefig(CHARTS / "hero_plan_vs_realized.png")
plt.close(fig)

# 2. ETF 구성종목 look-through: 초기 계획 비중 × 운용보고서 상위 10종목 비중
plan = pd.read_csv(TABLES / "plan_initial_weights.csv", dtype={"code": str})
hold = pd.read_csv(TABLES / "etf_holdings_top10.csv")
risky = plan[plan["asset_type"] == "위험"]
lt = hold[~hold["Holding"].astype(str).str.contains("금 현물")].merge(risky[["code", "etf", "weight_pct"]], left_on="Ticker", right_on="code")
lt["exposure_pct"] = lt["weight_pct"] * lt["Weight"] / 100
by_stock = lt.pivot_table(index="Holding", columns="etf", values="exposure_pct", aggfunc="sum", fill_value=0)
top = by_stock.loc[by_stock.sum(axis=1).sort_values(ascending=False).index[:12]]
top.to_csv(TABLES / "lookthrough_top_holdings.csv", encoding="utf-8-sig")
top_etfs = top.sum().sort_values(ascending=False)
main = list(top_etfs.index[:4])
plot = top[main].copy()
plot["기타 ETF"] = top.drop(columns=main).sum(axis=1)
colors = [PORT, ALT, ACCENT, AMBER, BENCH]
fig, ax = plt.subplots(figsize=(11, 6))
left = np.zeros(len(plot))
yy = np.arange(len(plot))[::-1]
for col, color in zip(plot.columns, colors):
    ax.barh(yy, plot[col], left=left, color=color, height=0.62, label=col, edgecolor="white", linewidth=1)
    left += plot[col].values
for yi, v in zip(yy, left):
    ax.text(v + 0.1, yi, f"{v:.2f}%", va="center", fontsize=9, color=INK)
ax.set_yticks(yy)
ax.set_yticklabels(plot.index, fontsize=10)
ax.set_xlabel("포트폴리오 전체 대비 간접 보유 비중 (%)")
ax.set_title("ETF를 열어보면: 초기 계획 포트폴리오의 개별 종목 look-through 노출 상위 12", loc="left")
ax.legend(frameon=False, fontsize=9, loc="lower right")
ax.grid(axis="y", visible=False)
fig.text(0.01, -0.03, "2026-04-30 ETF 월간 운용보고서의 상위 10종목 비중 × 초기 계획 비중. 상위 10종목 밖의 보유분은 빠져 있어 실제 노출보다 작게 잡힘.",
         fontsize=9, color=MUTED)
fig.savefig(CHARTS / "etf_lookthrough_exposure.png")
plt.close(fig)
print(top.sum(axis=1).round(2))

# 3. 계획 단계 백테스트 (2025-06-02 ~ 2026-05-22, 비용 반영)
nav = pd.read_csv(TABLES / "backtest_daily_nav_net.csv", index_col=0, parse_dates=True)
labels = {"Strategy_TargetVol_Rule": ("계획 규칙 (Target Vol Overlay)", PORT, 2.2),
          "Static_65_35_NoRebalance": ("초기 65:35, 리밸런싱 없음", ALT, 1.6),
          "Benchmark_65_35_TIGER200_CD": ("TIGER 200 65% + CD 35%", ACCENT, 1.6),
          "KOSPI200_100": ("KOSPI200 100%", BENCH, 1.6)}
fig, axes = plt.subplots(1, 2, figsize=(15, 4.8), gridspec_kw={"width_ratios": [1.6, 1]})
for col, (lab, color, lw) in labels.items():
    axes[0].plot(nav.index, nav[col], color=color, linewidth=lw, label=f"{lab}  {nav[col].iloc[-1]-1:+.0%}")
    dd = nav[col] / nav[col].cummax() - 1
    axes[1].plot(nav.index, dd * 100, color=color, linewidth=lw)
axes[0].set_title("계획 단계 백테스트 NAV — 강세장에서는 규칙이 수익을 덜 냈다", loc="left")
axes[0].legend(frameon=False, fontsize=9, loc="upper left")
axes[1].set_title("드로다운 (%) — 대신 낙폭은 가장 작았다", loc="left")
for a in axes:
    a.xaxis.set_major_formatter(mdates.DateFormatter("%y-%m"))
fig.savefig(CHARTS / "plan_backtest_nav_drawdown.png")
plt.close(fig)
