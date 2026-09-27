#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GAPS 투자계획서 백테스트 재현 스크립트.

노트북의 설명/시각화 셀을 제외하고 백테스트 산출에 필요한 코드만 정리했다.
실행하면 주 백테스트, 성과표/룩어헤드/비용 로그를 GAPS_백테스트_벤치마크_20260526.xlsx
파일로 저장한다.
"""

from pathlib import Path
import numpy as np
import pandas as pd

# 데이터 경로: 노트북을 현재 폴더에서 실행하는 경우를 우선하고, 기존 하드코딩 경로는 후순위 fallback으로 둔다.
DATA_DIR_CANDIDATES = [
    Path.cwd(),
    Path(__file__).resolve().parents[1] / 'input',
    Path.cwd() / 'input',
]
DATA_DIR = next((p for p in DATA_DIR_CANDIDATES if (p / 'universe.xlsx').exists()), DATA_DIR_CANDIDATES[0])
UNIVERSE_PATH = DATA_DIR / 'universe.xlsx'
FACTSHEET_DIR = DATA_DIR / 'factsheet'
PRICE_PATH = DATA_DIR / 'DBGAPS_2021_2026_수정주가_일간비영업일제외.xlsx'
VOLUME_PATH = DATA_DIR / 'DBGAPS_2021_2026_거래량_일간비영업일제외.xlsx'
VALUE_PATH = DATA_DIR / 'DBGAPS_2021_2026_거래대금_일간비영업일제외.xlsx'
VKOSPI_PATH = DATA_DIR / 'VKOSPI.xlsx'
MARKET_INDEX_PATH = DATA_DIR / '코스피, 코스피200, 코스닥150.xlsx'
REPORT_PATH = DATA_DIR / '자료' / '시장전망' / '삼성증권_DerivativesWeekly_SpotUpVolUp_TargetVol_20260508.pdf'

# ===== 백테스트 핵심 파라미터 =====
TRADING_DAYS = 252
BACKTEST_START_REQUESTED = pd.Timestamp('2025-06-02')
BACKTEST_END_REQUESTED = pd.Timestamp('2026-05-22')

# 거래비용: 대회 공식 기준은 국내 ETF 거래수수료 0.1%, 제세금 없음이다.
# 슬리피지는 공식 NAV 산정 비용이 아니므로 비용 반영 후 표에는 별도 차감하지 않는다.
COMMISSION_RATE = 0.00100          # 대회 기준 국내 ETF 거래수수료 0.1%
TAX_RATE = 0.0                     # 제세금 없음
APPLY_SLIPPAGE_IN_OFFICIAL_BACKTEST = False
SLIPPAGE_LARGE_ETF = 0.0           # 공식 NAV 비용에는 슬리피지 미반영
SLIPPAGE_THEME_SMALL_ETF = 0.0     # 공식 NAV 비용에는 슬리피지 미반영
INCLUDE_INITIAL_TRADING_COST = True # 최초 포트폴리오 설정 매매비용 포함
LARGE_LIQUID_TICKERS = {
    'A069500',  # KODEX 200
    'A102110',  # TIGER 200
    'A229200',  # KODEX 코스닥150
    'A360750',  # TIGER 미국S&P500
    'A360200',  # ACE 미국S&P500
    'A379800',  # KODEX 미국S&P500
    'A133690',  # TIGER 미국나스닥100
    'A411060',  # ACE KRX금현물
    'A459580',  # KODEX CD금리액티브(합성)
    'A423160',  # KODEX KOFR금리액티브(합성)
    'A488770',  # KODEX 머니마켓액티브
    'A157450',  # TIGER 단기통안채
    'A0061Z0',  # RISE 단기특수은행채액티브
    'A453850',  # ACE 미국30년국채액티브(H)
}

# 무위험수익률/Rf 및 Sortino MAR: None이면 백테스트 기간의 CD금리 ETF 평균 일수익률을 연환산해 사용한다.
RF_PROXY_TICKER = 'A459580'
RF_PROXY_DESCRIPTION = 'KODEX CD금리액티브(합성) 일별 수익률 평균의 252영업일 연환산값'
RISK_FREE_RATE_ANNUAL = None
MAR_ANNUAL = None

# 룩어헤드 방지 및 리밸런싱 체결 가정
SIGNAL_LOOKBACK_WINDOWS = (20, 60)
SIGNAL_EXECUTION_ASSUMPTION = 't일 종가 신호(VKOSPI/KOSPI200) 확인 후 t+1영업일 시가 체결'
REBALANCE_EXECUTION_ASSUMPTION = '월말 종가로 신호 확정, 다음 영업일(월초) 시가에 리밸런싱'
PRICE_EXECUTION_PROXY = '현재 파일에는 일별 수정주가 종가만 있어 t+1 수정종가를 월초 시가 체결의 보수적 프록시로 사용'

# 계획서 규칙형 백테스트 파라미터
# Target Vol 효과를 '방어 overlay'로 분리하려면 65% 초과 위험확대를 끄고,
# 70% 확대 옵션은 별도 민감도 분석으로 켠다.
PLAN_RULES_ALLOW_RISK_EXPANSION = False
PLAN_RULES_REQUIRE_FALLING_VKOSPI_FOR_EXPANSION = True
PLAN_RULES_BASE_RISK_BUDGET = 0.65
PLAN_RULES_EXPANSION_RISK_BUDGET = 0.70
PLAN_RULES_DEFENSE_RISK_BUDGET = 0.55
PLAN_RULES_FORCED_DEFENSE_RISK_BUDGET = 0.50
PLAN_RULES_REBALANCE_BAND_ABS = 0.01
PLAN_RULES_RISK_SOFT_BAND = 0.74
PLAN_RULES_RISK_HARD_BAND = 0.76
PLAN_RULES_DOMESTIC_SECTOR_SOFT_BAND = 0.17
PLAN_RULES_DOMESTIC_SECTOR_HARD_BAND = 0.18
PLAN_RULES_SINGLE_ETF_SOFT_BAND = 0.23
PLAN_RULES_SINGLE_ETF_HARD_BAND = 0.25
PLAN_RULES_SINGLE_ETF_TRIM_TARGET = 0.205

universe = pd.read_excel(UNIVERSE_PATH, sheet_name='기본정보')

backtest_parameters = pd.DataFrame([
    {'Parameter': 'Backtest request', 'Value': f'{BACKTEST_START_REQUESTED.date()} ~ {BACKTEST_END_REQUESTED.date()}'},
    {'Parameter': 'Contest ETF commission', 'Value': f'{COMMISSION_RATE:.2%}'},
    {'Parameter': 'Tax', 'Value': f'{TAX_RATE:.2%}'},
    {'Parameter': 'Slippage in official NAV', 'Value': 'Not deducted' if not APPLY_SLIPPAGE_IN_OFFICIAL_BACKTEST else 'Deducted'},
    {'Parameter': 'Initial trading cost', 'Value': 'Included' if INCLUDE_INITIAL_TRADING_COST else 'Excluded'},
    {'Parameter': 'Rf source', 'Value': RF_PROXY_DESCRIPTION if RISK_FREE_RATE_ANNUAL is None else f'User fixed {RISK_FREE_RATE_ANNUAL:.2%}'},
    {'Parameter': 'MAR for Sortino', 'Value': 'Same as Rf' if MAR_ANNUAL is None else f'{MAR_ANNUAL:.2%}'},
    {'Parameter': 'Signal execution', 'Value': SIGNAL_EXECUTION_ASSUMPTION},
    {'Parameter': 'Monthly rebalance', 'Value': REBALANCE_EXECUTION_ASSUMPTION},
    {'Parameter': 'Plan-rules risk expansion', 'Value': 'ON, only under strict falling-VKOSPI confirmation' if PLAN_RULES_ALLOW_RISK_EXPANSION else 'OFF for Target Vol effect table; 70% expansion is separate sensitivity'},
    {'Parameter': 'Plan-rules monthly trade band', 'Value': f'{PLAN_RULES_REBALANCE_BAND_ABS:.1%} absolute weight drift for scheduled rotation trades'},
    {'Parameter': 'Plan-rules risk band', 'Value': f'{PLAN_RULES_RISK_SOFT_BAND:.0%} monthly trim / {PLAN_RULES_RISK_HARD_BAND:.0%} mid-month trim'},
    {'Parameter': 'Plan-rules sector/single ETF band', 'Value': f'sector {PLAN_RULES_DOMESTIC_SECTOR_SOFT_BAND:.0%}/{PLAN_RULES_DOMESTIC_SECTOR_HARD_BAND:.0%}, single {PLAN_RULES_SINGLE_ETF_SOFT_BAND:.0%}/{PLAN_RULES_SINGLE_ETF_HARD_BAND:.0%}'},
])
print(backtest_parameters.to_string(index=False))

def load_dbgaps_xlsx(path):
    # DBGAPS 엑셀: 9행 코드, 10행 코드명, 15행부터 일별 데이터.
    raw = pd.read_excel(path, sheet_name='Sheet1', header=None)
    codes = raw.iloc[8, 1:].dropna().astype(str).str.strip().tolist()
    names = raw.iloc[9, 1:1 + len(codes)].tolist()
    data = raw.iloc[14:, :1 + len(codes)].copy()
    data.columns = ['Date'] + codes
    data['Date'] = pd.to_datetime(data['Date'])
    data = data.set_index('Date').apply(pd.to_numeric, errors='coerce')
    return data, dict(zip(codes, names))

price, code_names = load_dbgaps_xlsx(PRICE_PATH)


def load_vkospi_xlsx(path=VKOSPI_PATH):
    raw = pd.read_excel(path, sheet_name='Sheet1', header=None)
    df = raw.iloc[14:, :2].copy()
    df.columns = ['Date', 'VKOSPI']
    df['Date'] = pd.to_datetime(df['Date'])
    df['VKOSPI'] = pd.to_numeric(df['VKOSPI'], errors='coerce')
    return df.dropna().set_index('Date').sort_index()

vkospi = load_vkospi_xlsx()


def load_market_index_xlsx(path=MARKET_INDEX_PATH):
    raw = pd.read_excel(path, sheet_name='Sheet1', header=None)
    names = raw.iloc[9, 1:4].tolist()
    df = raw.iloc[14:, :4].copy()
    df.columns = ['Date'] + names
    df['Date'] = pd.to_datetime(df['Date'])
    for col in names:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    return df.dropna().set_index('Date').sort_index()

market_index = load_market_index_xlsx()


import math

BACKTEST_OUTPUT_PATH = DATA_DIR / 'GAPS_백테스트_벤치마크_20260526.xlsx'

BACKTEST_INITIAL_PORTFOLIO = pd.DataFrame([
    # 2025년 5월 자료만 보고 정한 백테스트 전용 초기 포트폴리오.
    # 2026년 제출용 포트폴리오와 분리해 미래정보 편향을 줄인다.
    {'Ticker': 'A102110', 'Weight': 0.17, 'Role': '2025 백테스트 국내 코어: TIGER 200'},
    {'Ticker': 'A069500', 'Weight': 0.03, 'Role': '2025 백테스트 국내 코어 보조: KODEX 200'},
    {'Ticker': 'A379800', 'Weight': 0.11, 'Role': '2025 백테스트 해외 코어: 미국 S&P500'},
    {'Ticker': 'A133690', 'Weight': 0.06, 'Role': '2025 백테스트 해외 성장: NASDAQ100'},
    {'Ticker': 'A395160', 'Weight': 0.02, 'Role': '2025 백테스트 국내 AI반도체'},
    {'Ticker': 'A487240', 'Weight': 0.03, 'Role': '2025 백테스트 국내 AI전력'},
    {'Ticker': 'A466920', 'Weight': 0.03, 'Role': '2025 백테스트 조선'},
    {'Ticker': 'A449450', 'Weight': 0.03, 'Role': '2025 백테스트 방산'},
    {'Ticker': 'A434730', 'Weight': 0.02, 'Role': '2025 백테스트 원자력'},
    {'Ticker': 'A466940', 'Weight': 0.02, 'Role': '2025 백테스트 은행/배당'},
    {'Ticker': 'A487230', 'Weight': 0.05, 'Role': '2025 백테스트 해외 AI전력'},
    {'Ticker': 'A390390', 'Weight': 0.03, 'Role': '2025 백테스트 해외 반도체'},
    {'Ticker': 'A411060', 'Weight': 0.05, 'Role': '2025 백테스트 금 현물'},
    {'Ticker': 'A459580', 'Weight': 0.15, 'Role': '2025 백테스트 CD 금리'},
    {'Ticker': 'A488770', 'Weight': 0.10, 'Role': '2025 백테스트 머니마켓'},
    {'Ticker': 'A157450', 'Weight': 0.07, 'Role': '2025 백테스트 단기통안채'},
    {'Ticker': 'A423160', 'Weight': 0.03, 'Role': '2025 백테스트 KOFR 초단기 금리'},
])

portfolio_meta = BACKTEST_INITIAL_PORTFOLIO.merge(
    universe[['Ticker', 'Name', 'Class1', 'Class2']],
    on='Ticker',
    how='left'
)
strategy_tickers = portfolio_meta['Ticker'].tolist()

first_dates = price[strategy_tickers].apply(lambda s: s.dropna().index.min())
last_dates = price[strategy_tickers].apply(lambda s: s.dropna().index.max())
BACKTEST_START = max(BACKTEST_START_REQUESTED, first_dates.max())
BACKTEST_END = min(BACKTEST_END_REQUESTED, last_dates.min(), market_index.index.max(), vkospi.index.max())

print('Backtest period:', BACKTEST_START.date(), '~', BACKTEST_END.date())
print('Common-period driver:', first_dates.idxmax(), code_names.get(first_dates.idxmax()), first_dates.max().date())

risk_base = portfolio_meta.loc[portfolio_meta['Class1'] == '위험', ['Ticker', 'Weight']].set_index('Ticker')['Weight']
safe_base = portfolio_meta.loc[portfolio_meta['Class1'] == '안전', ['Ticker', 'Weight']].set_index('Ticker')['Weight']
risk_base_norm = risk_base / risk_base.sum()
safe_base_norm = safe_base / safe_base.sum()


def scaled_weights_by_risk_budget(risk_weights_norm, safe_weights_norm, risk_budget):
    risk_scaled = risk_weights_norm * risk_budget
    safe_scaled = safe_weights_norm * (1 - risk_budget)
    w = pd.concat([risk_scaled, safe_scaled]).sort_index()
    return w / w.sum()


def target_weights_by_risk_budget(risk_budget):
    return scaled_weights_by_risk_budget(risk_base_norm, safe_base_norm, risk_budget)


def month_first_dates(index):
    idx = pd.Index(index).sort_values()
    return idx.to_series().groupby(idx.to_period('M')).first().tolist()


def trading_cost_rates(tickers):
    # 대회 공식 NAV 기준: 국내 ETF 거래수수료 0.1%, 제세금 없음.
    # 슬리피지는 성과표 비용으로 차감하지 않고 주문 체결 리스크로 별도 관리한다.
    return pd.Series({ticker: COMMISSION_RATE + TAX_RATE for ticker in tickers}, dtype=float)


def estimate_rebalance_cost(current_value, target_weights, nav_before_cost, tickers):
    target_value = target_weights.reindex(tickers).fillna(0.0) * nav_before_cost
    trades_value = target_value - current_value.reindex(tickers).fillna(0.0)
    turnover_by_ticker = trades_value.abs() / nav_before_cost
    rates = trading_cost_rates(tickers)
    cost = float((turnover_by_ticker * rates).sum() * nav_before_cost)
    turnover = float(turnover_by_ticker.sum())
    cost_rate = float(cost / nav_before_cost) if nav_before_cost else np.nan
    return cost, turnover, cost_rate


def fixed_weight_nav(price_df, target_weights, start=BACKTEST_START, end=BACKTEST_END,
                     rebalance_monthly=True, apply_costs=False, portfolio_name='Portfolio'):
    tickers = list(target_weights.index)
    px = price_df.loc[start:end, tickers].ffill().dropna()
    target_weights = target_weights.reindex(tickers).astype(float)
    target_weights = target_weights / target_weights.sum()
    rebal_dates = set(month_first_dates(px.index)) if rebalance_monthly else set()
    holdings = None
    nav_rows = []
    cost_rows = []

    for dt, row in px.iterrows():
        if holdings is None:
            nav = 1.0
            if apply_costs and INCLUDE_INITIAL_TRADING_COST:
                zero_value = pd.Series(0.0, index=tickers)
                cost, turnover, cost_rate = estimate_rebalance_cost(zero_value, target_weights, nav, tickers)
                nav_after_cost = nav - cost
                cost_rows.append({
                    'Date': dt,
                    'Portfolio': portfolio_name,
                    'Turnover': turnover,
                    'TradingCost': cost,
                    'TradingCostRate': cost_rate,
                })
            else:
                nav_after_cost = nav
            holdings = target_weights * nav_after_cost / row
            nav = float((holdings * row).sum())
        else:
            nav = float((holdings * row).sum())
            if dt in rebal_dates:
                current_value = holdings * row
                cost, turnover, cost_rate = estimate_rebalance_cost(current_value, target_weights, nav, tickers)
                nav_after_cost = nav - cost if apply_costs else nav
                holdings = target_weights * nav_after_cost / row
                nav = float((holdings * row).sum())
                cost_rows.append({
                    'Date': dt,
                    'Portfolio': portfolio_name,
                    'Turnover': turnover,
                    'TradingCost': cost if apply_costs else 0.0,
                    'TradingCostRate': cost_rate if apply_costs else 0.0,
                })
        nav_rows.append((dt, nav))

    nav = pd.Series(dict(nav_rows), name=portfolio_name).sort_index()
    costs = pd.DataFrame(cost_rows)
    return nav, costs


def latest_signal_snapshot_for_trade(trade_dt):
    # 룩어헤드 방지: trade_dt 당일 종가를 쓰지 않고, trade_dt 이전에 확정된 마지막 종가 신호만 사용한다.
    vk_hist = vkospi.loc[vkospi.index < trade_dt, 'VKOSPI'].dropna()
    kospi_hist = market_index['코스피 200'].loc[market_index.index < trade_dt].dropna()

    if vk_hist.empty:
        signal_date = pd.NaT
        v = np.nan
        v20 = np.nan
        v60 = np.nan
    else:
        signal_date = vk_hist.index[-1]
        v = float(vk_hist.iloc[-1])
        v20 = float(vk_hist.tail(20).mean()) if len(vk_hist) >= 20 else np.nan
        v60 = float(vk_hist.tail(60).mean()) if len(vk_hist) >= 60 else np.nan

    if kospi_hist.empty:
        kospi = np.nan
        ma20 = np.nan
        ma60 = np.nan
        above20 = np.nan
        above60 = np.nan
    else:
        kospi = float(kospi_hist.iloc[-1])
        ma20 = float(kospi_hist.tail(20).mean()) if len(kospi_hist) >= 20 else np.nan
        ma60 = float(kospi_hist.tail(60).mean()) if len(kospi_hist) >= 60 else np.nan
        above20 = bool(kospi > ma20) if not pd.isna(ma20) else np.nan
        above60 = bool(kospi > ma60) if not pd.isna(ma60) else np.nan

    return {
        'Signal_Date': signal_date,
        'VKOSPI': v,
        'VKOSPI_20D_MA': v20,
        'VKOSPI_60D_MA': v60,
        'KOSPI200_Close': kospi,
        'KOSPI200_MA20': ma20,
        'KOSPI200_MA60': ma60,
        'KOSPI200_Above_MA20': above20,
        'KOSPI200_Above_MA60': above60,
    }


def risk_budget_from_trade_signal(trade_dt):
    snap = latest_signal_snapshot_for_trade(trade_dt)
    v = snap['VKOSPI']
    above20 = snap['KOSPI200_Above_MA20']
    above60 = snap['KOSPI200_Above_MA60']
    trend_damaged = above20 is False

    if pd.isna(v):
        risk_budget = PLAN_RULES_BASE_RISK_BUDGET
        signal = 'VKOSPI 없음: 기본 65%'
    elif v > 80 and trend_damaged:
        risk_budget = PLAN_RULES_FORCED_DEFENSE_RISK_BUDGET
        signal = 'VKOSPI 80 초과 + KOSPI200 추세 훼손: 강한 방어'
    elif v > 70 and trend_damaged:
        risk_budget = PLAN_RULES_DEFENSE_RISK_BUDGET
        signal = 'VKOSPI 70 초과 + KOSPI200 20일선 이탈: 방어 전환'
    elif v > 70:
        risk_budget = PLAN_RULES_BASE_RISK_BUDGET
        signal = 'VKOSPI 70 초과 단독: 고변동성 상승장, 위험자산 유지'
    elif v <= 60 and above20 is True and above60 is True:
        risk_budget = PLAN_RULES_EXPANSION_RISK_BUDGET
        signal = 'VKOSPI 60 이하 + KOSPI200 20/60일선 상회: 위험 확대'
    else:
        risk_budget = PLAN_RULES_BASE_RISK_BUDGET
        signal = '고변동성 또는 추세 확인 부족: 기본 65%'

    info = {**snap, 'Risk_Budget': risk_budget, 'Signal': signal}
    return risk_budget, info

def vkospi_risk_budget(trade_dt):
    risk_budget, info = risk_budget_from_trade_signal(trade_dt)
    return target_weights_by_risk_budget(risk_budget), info


def monthly_dynamic_nav(price_df, tickers, weight_func, start=BACKTEST_START, end=BACKTEST_END,
                        apply_costs=False, portfolio_name='Strategy_TargetVol_Rule'):
    px = price_df.loc[start:end, tickers].ffill().dropna()
    rebal_dates = set(month_first_dates(px.index))
    holdings = None
    nav_rows = []
    schedule_rows = []

    for dt, row in px.iterrows():
        if holdings is None:
            nav = 1.0
            current_weights, budget_info = weight_func(dt)
            current_weights = current_weights.reindex(tickers).fillna(0.0)
            if apply_costs and INCLUDE_INITIAL_TRADING_COST:
                zero_value = pd.Series(0.0, index=tickers)
                cost, turnover, cost_rate = estimate_rebalance_cost(zero_value, current_weights, nav, tickers)
                nav_after_cost = nav - cost
            else:
                cost, turnover, cost_rate = 0.0, 0.0, 0.0
                nav_after_cost = nav
            holdings = current_weights * nav_after_cost / row
            nav = float((holdings * row).sum())
            schedule_rows.append({'Date': dt, 'Rebalance_Type': 'Initial', 'Turnover': turnover, 'TradingCost': cost if apply_costs else 0.0, 'TradingCostRate': cost_rate if apply_costs else 0.0, **budget_info})
        else:
            nav = float((holdings * row).sum())
            if dt in rebal_dates:
                new_weights, budget_info = weight_func(dt)
                new_weights = new_weights.reindex(tickers).fillna(0.0)
                current_value = holdings * row
                cost, turnover, cost_rate = estimate_rebalance_cost(current_value, new_weights, nav, tickers)
                nav_after_cost = nav - cost if apply_costs else nav
                holdings = new_weights * nav_after_cost / row
                nav = float((holdings * row).sum())
                schedule_rows.append({
                    'Date': dt,
                    'Rebalance_Type': 'Monthly/Signal',
                    'Turnover': turnover,
                    'TradingCost': cost if apply_costs else 0.0,
                    'TradingCostRate': cost_rate if apply_costs else 0.0,
                    **budget_info,
                })
        nav_rows.append((dt, nav))

    nav = pd.Series(dict(nav_rows), name=portfolio_name).sort_index()
    schedule = pd.DataFrame(schedule_rows).drop_duplicates('Date').set_index('Date')
    return nav, schedule


def normalized_series(series, start=BACKTEST_START, end=BACKTEST_END, name=None):
    s = series.loc[start:end].dropna().astype(float)
    s = s / s.iloc[0]
    s.name = name or series.name
    return s


def risk_free_inputs(start=BACKTEST_START, end=BACKTEST_END):
    if RISK_FREE_RATE_ANNUAL is None:
        rf_price = price[RF_PROXY_TICKER].loc[start:end].ffill().dropna()
        rf_returns = rf_price.pct_change().dropna()
        rf_daily = float(rf_returns.mean())
        rf_ann = float((1 + rf_daily) ** TRADING_DAYS - 1)
        rf_source = f'{RF_PROXY_TICKER} {code_names.get(RF_PROXY_TICKER, RF_PROXY_DESCRIPTION)}: 백테스트 기간 일별 수익률 평균의 252영업일 연환산'
    else:
        rf_ann = float(RISK_FREE_RATE_ANNUAL)
        rf_daily = float((1 + rf_ann) ** (1 / TRADING_DAYS) - 1)
        rf_source = '상단 파라미터 RISK_FREE_RATE_ANNUAL 고정값'

    if MAR_ANNUAL is None:
        mar_ann = rf_ann
        mar_daily = rf_daily
        mar_source = 'Sortino MAR = Rf'
    else:
        mar_ann = float(MAR_ANNUAL)
        mar_daily = float((1 + mar_ann) ** (1 / TRADING_DAYS) - 1)
        mar_source = '상단 파라미터 MAR_ANNUAL 고정값'

    return rf_daily, rf_ann, mar_daily, mar_ann, rf_source, mar_source


rf_daily, rf_ann, mar_daily, mar_ann, rf_source, mar_source = risk_free_inputs()
print('Risk-free rate used:', f'{rf_ann:.2%}', '|', rf_source)
print('Sortino MAR used:', f'{mar_ann:.2%}', '|', mar_source)


# ===== 계획서 규칙형 Target Vol / 로테이션 백테스트 =====
# 기존 monthly_dynamic_nav는 단순 위험예산 시뮬레이션이다.
# 아래 plan_rules_nav는 계획서의 월간 점검, NAV/MDD 방어, 섹터 로테이션 후보,
# 밴드 리밸런싱을 수치화한 별도 실행 엔진이다.
ROTATION_CANDIDATE_TICKERS = [
    'A305720',  # KODEX 2차전지산업
    'A434730',  # HANARO 원자력iSelect
    'A445290',  # KODEX 로봇액티브
    'A091170',  # KODEX 은행
    'A453640',  # KODEX 미국S&P500헬스케어
    'A453630',  # KODEX 미국S&P500필수소비재
]
PLAN_TICKERS = sorted(set(strategy_tickers + [t for t in ROTATION_CANDIDATE_TICKERS if t in price.columns]))

plan_meta = universe[['Ticker', 'Name', 'Class1', 'Class2']].drop_duplicates('Ticker').set_index('Ticker')
PLAN_RISK_TICKERS = [t for t in PLAN_TICKERS if plan_meta.loc[t, 'Class1'] == '위험']
PLAN_SAFE_TICKERS = [t for t in PLAN_TICKERS if plan_meta.loc[t, 'Class1'] == '안전']
PLAN_DOMESTIC_SECTOR_TICKERS = [t for t in PLAN_TICKERS if plan_meta.loc[t, 'Class2'] == '국내주식_섹터']

DEFENSE_REDUCTION_ORDER = [
    'A466920',  # SOL 조선TOP3플러스
    'A449450',  # PLUS K방산
    'A487240',  # KODEX AI전력핵심설비
    'A229200',  # KODEX 코스닥150
    'A395160',  # KODEX AI반도체
    'A133690',  # TIGER 미국나스닥100
    'A487230',  # KODEX 미국AI전력핵심인프라
    'A458730',  # TIGER 미국배당다우존스
    'A411060',  # ACE KRX금현물
    'A102110',  # TIGER 200
    'A379800',  # KODEX 미국S&P500
]
CASH_RECEIVER_WEIGHTS = pd.Series({
    'A459580': 0.40,  # KODEX CD금리액티브
    'A488770': 0.30,  # KODEX 머니마켓액티브
    'A157450': 0.15,  # TIGER 단기통안채
    'A0061Z0': 0.10,  # RISE 단기특수은행채액티브
    'A453850': 0.05,  # ACE 미국30년국채액티브(H)
})
RISK_EXPANSION_WEIGHTS = pd.Series({'A102110': 0.50, 'A379800': 0.30, 'A487230': 0.20})


def normalize_weights(weights):
    w = weights.reindex(PLAN_TICKERS).fillna(0.0).clip(lower=0.0)
    total = float(w.sum())
    return w / total if total else w


def risk_weight_of(weights):
    return float(weights.reindex(PLAN_RISK_TICKERS).fillna(0.0).sum())


def safe_weight_of(weights):
    return float(weights.reindex(PLAN_SAFE_TICKERS).fillna(0.0).sum())


def add_to_cash(weights, amount):
    w = weights.copy()
    receivers = CASH_RECEIVER_WEIGHTS.reindex(w.index).dropna()
    receivers = receivers / receivers.sum()
    for ticker, share in receivers.items():
        w.loc[ticker] = w.get(ticker, 0.0) + amount * share
    return normalize_weights(w)


def fund_from_safe(weights, amount):
    w = weights.copy()
    remaining = amount
    for ticker in ['A459580', 'A488770', 'A157450', 'A0061Z0', 'A453850']:
        if ticker not in w.index or remaining <= 0:
            continue
        sell = min(float(w.loc[ticker]), remaining)
        w.loc[ticker] -= sell
        remaining -= sell
    return w, amount - remaining


def reduce_risk_to_budget(weights, target_risk, reason):
    w = normalize_weights(weights)
    current_risk = risk_weight_of(w)
    cut_needed = max(0.0, current_risk - target_risk)
    if cut_needed <= 1e-9:
        return w, []

    remaining = cut_needed
    actions = []
    for ticker in DEFENSE_REDUCTION_ORDER:
        if ticker not in w.index or remaining <= 1e-9:
            continue
        sell = min(float(w.loc[ticker]), remaining)
        if sell > 1e-9:
            w.loc[ticker] -= sell
            remaining -= sell
            actions.append(f'{reason}: {ticker} {sell:.2%} 축소')
    executed = cut_needed - remaining
    w = add_to_cash(w, executed)
    return normalize_weights(w), actions


def increase_risk_to_budget(weights, target_risk, reason):
    w = normalize_weights(weights)
    current_risk = risk_weight_of(w)
    add_needed = max(0.0, target_risk - current_risk)
    if add_needed <= 1e-9:
        return w, []
    w, funded = fund_from_safe(w, add_needed)
    if funded <= 1e-9:
        return normalize_weights(w), []

    add_targets = RISK_EXPANSION_WEIGHTS.reindex(w.index).dropna()
    add_targets = add_targets / add_targets.sum()
    for ticker, share in add_targets.items():
        w.loc[ticker] = w.get(ticker, 0.0) + funded * share
    return normalize_weights(w), [f'{reason}: 위험자산 {funded:.2%} 확대']


def trailing_return_on(ticker, signal_date, days):
    if ticker not in price.columns or pd.isna(signal_date):
        return np.nan
    s = price[ticker].loc[:signal_date].dropna()
    if len(s) <= days:
        return np.nan
    return float(s.iloc[-1] / s.iloc[-days - 1] - 1)


def kospi200_trailing_return(signal_date, days):
    s = market_index['코스피 200'].loc[:signal_date].dropna()
    if len(s) <= days:
        return np.nan
    return float(s.iloc[-1] / s.iloc[-days - 1] - 1)


def exchange_weight(weights, from_ticker, to_ticker, amount, reason):
    w = weights.copy()
    if from_ticker not in w.index or to_ticker not in w.index:
        return w, []
    trade = min(float(w.loc[from_ticker]), amount)
    if trade <= 1e-9:
        return w, []
    w.loc[from_ticker] -= trade
    w.loc[to_ticker] = w.get(to_ticker, 0.0) + trade
    return normalize_weights(w), [f'{reason}: {from_ticker}->{to_ticker} {trade:.2%}']


def apply_sector_rotation(weights, trade_dt, signal_info, nav_state):
    # 월간 점검에서만 작동한다. 방어 상태/VKOSPI 70 초과에서는 신규 2차전지·로봇 편입을 막는다.
    w = normalize_weights(weights)
    actions = []
    signal_date = signal_info.get('Signal_Date')
    if pd.isna(signal_date) or trade_dt < pd.Timestamp('2025-07-01'):
        return w, actions
    if signal_info.get('VKOSPI', np.nan) > 70 or nav_state in ['방어 전환', '강제 방어']:
        return w, ['방어 상태: 신규 테마 로테이션 보류']

    kospi_1m = kospi200_trailing_return(signal_date, 21)
    kospi_3m = kospi200_trailing_return(signal_date, 63)

    # 조선·방산 1개월/3개월 상대열위 + 2차전지 개선 시 1~3%p 교체.
    for donor in ['A466920', 'A449450']:
        donor_1m = trailing_return_on(donor, signal_date, 21)
        donor_3m = trailing_return_on(donor, signal_date, 63)
        battery_1m = trailing_return_on('A305720', signal_date, 21)
        battery_3m = trailing_return_on('A305720', signal_date, 63)
        if all(pd.notna(x) for x in [donor_1m, donor_3m, battery_1m, battery_3m, kospi_1m, kospi_3m]):
            donor_weak = donor_1m < kospi_1m and donor_3m < kospi_3m
            battery_improving = battery_1m > 0 and battery_3m > 0 and battery_1m > donor_1m and battery_3m > donor_3m
            if donor_weak and battery_improving:
                w, a = exchange_weight(w, donor, 'A305720', 0.015, '국내 섹터 상대모멘텀 로테이션')
                actions.extend(a)

    # AI전력 과열 후 원자력 상대모멘텀 개선 시 1%p 제한 교체.
    ai_power_3m = trailing_return_on('A487240', signal_date, 63)
    nuclear_1m = trailing_return_on('A434730', signal_date, 21)
    ai_power_1m = trailing_return_on('A487240', signal_date, 21)
    if all(pd.notna(x) for x in [ai_power_3m, nuclear_1m, ai_power_1m]):
        if ai_power_3m > 0.30 and nuclear_1m > ai_power_1m:
            w, a = exchange_weight(w, 'A487240', 'A434730', 0.01, 'AI전력 과열 완화 로테이션')
            actions.extend(a)

    return normalize_weights(w), actions


def apply_limit_bands(weights, current_weights, signal_info):
    # 평가비중 초과는 즉시 전량 복원하지 않는다.
    # 월간 점검에서는 soft band를, 월중에는 hard band 또는 추세 훼손 동반 초과를 중심으로 일부 이익실현한다.
    w = normalize_weights(weights.copy())
    actions = []
    trend_damaged = signal_info.get('KOSPI200_Above_MA20') is False

    risk_w = risk_weight_of(current_weights)
    if risk_w > PLAN_RULES_RISK_SOFT_BAND:
        w, a = reduce_risk_to_budget(
            current_weights,
            PLAN_RULES_EXPANSION_RISK_BUDGET,
            f'위험자산 {PLAN_RULES_RISK_SOFT_BAND:.0%} 초과 완화 밴드'
        )
        actions.extend(a)

    domestic_sector_w = float(current_weights.reindex(PLAN_DOMESTIC_SECTOR_TICKERS).fillna(0.0).sum())
    if domestic_sector_w > PLAN_RULES_DOMESTIC_SECTOR_SOFT_BAND:
        sector_weights = current_weights.reindex(PLAN_DOMESTIC_SECTOR_TICKERS).fillna(0.0)
        top_ticker = sector_weights.idxmax()
        cut = min(float(w.get(top_ticker, 0.0)), domestic_sector_w - 0.16)
        if cut > 1e-9:
            w.loc[top_ticker] -= cut
            w = add_to_cash(w, cut)
            suffix = ' + 추세 훼손' if trend_damaged else ''
            actions.append(f'국내주식_섹터 {PLAN_RULES_DOMESTIC_SECTOR_SOFT_BAND:.0%} 초과 밴드{suffix}: {top_ticker} {cut:.2%} 이익실현')

    for ticker, weight in current_weights.items():
        if weight > PLAN_RULES_SINGLE_ETF_SOFT_BAND and ticker in w.index:
            cut = min(float(w.loc[ticker]), weight - PLAN_RULES_SINGLE_ETF_TRIM_TARGET)
            if cut > 1e-9:
                w.loc[ticker] -= cut
                w = add_to_cash(w, cut)
                suffix = ' + 추세 훼손' if trend_damaged else ''
                actions.append(f'개별 ETF {PLAN_RULES_SINGLE_ETF_SOFT_BAND:.0%} 초과 밴드{suffix}: {ticker} {cut:.2%} 이익실현')

    return normalize_weights(w), actions

def monthly_band_rebalance(current_weights, desired_weights, band=PLAN_RULES_REBALANCE_BAND_ABS):
    # 목표와 1%p 이상 벌어진 종목만 조정하고, 잔여 차이는 CD/MMF가 흡수한다.
    current = normalize_weights(current_weights)
    desired = normalize_weights(desired_weights)
    new_w = current.copy()
    for ticker in desired.index:
        if abs(float(current.get(ticker, 0.0) - desired.get(ticker, 0.0))) >= band:
            new_w.loc[ticker] = desired.loc[ticker]
    diff = 1.0 - float(new_w.sum())
    if diff > 0:
        new_w = add_to_cash(new_w, diff)
    elif diff < 0:
        new_w, _ = fund_from_safe(new_w, -diff)
    return normalize_weights(new_w)


def nav_defense_state(nav, peak_nav):
    nav_loss = nav - 1.0
    mdd = nav / peak_nav - 1.0 if peak_nav else 0.0
    if nav_loss <= -0.07 or mdd <= -0.08:
        return '강제 방어', PLAN_RULES_FORCED_DEFENSE_RISK_BUDGET, nav_loss, mdd
    if nav_loss <= -0.05 or mdd <= -0.06:
        return '방어 전환', PLAN_RULES_DEFENSE_RISK_BUDGET, nav_loss, mdd
    if nav_loss <= -0.03 or mdd <= -0.04:
        return '주의', PLAN_RULES_BASE_RISK_BUDGET, nav_loss, mdd
    return '정상', PLAN_RULES_BASE_RISK_BUDGET, nav_loss, mdd


def expansion_allowed(signal_info, nav_state):
    if not PLAN_RULES_ALLOW_RISK_EXPANSION or nav_state != '정상':
        return False
    if signal_info.get('VKOSPI', np.nan) > 60:
        return False
    if signal_info.get('KOSPI200_Above_MA20') is not True or signal_info.get('KOSPI200_Above_MA60') is not True:
        return False
    if PLAN_RULES_REQUIRE_FALLING_VKOSPI_FOR_EXPANSION:
        v = signal_info.get('VKOSPI', np.nan)
        v20 = signal_info.get('VKOSPI_20D_MA', np.nan)
        v60 = signal_info.get('VKOSPI_60D_MA', np.nan)
        return pd.notna(v) and pd.notna(v20) and pd.notna(v60) and v < v20 < v60
    return True


def desired_plan_weights(current_weights, trade_dt, signal_info, nav_state, nav_target_risk, monthly_check):
    w = normalize_weights(current_weights)
    actions = []

    v = signal_info.get('VKOSPI', np.nan)
    trend_damaged = signal_info.get('KOSPI200_Above_MA20') is False
    nav_pressure = nav_state in ['주의', '방어 전환', '강제 방어']
    signal_risk_target = PLAN_RULES_BASE_RISK_BUDGET

    if pd.notna(v) and v > 80 and (trend_damaged or nav_pressure):
        signal_risk_target = PLAN_RULES_FORCED_DEFENSE_RISK_BUDGET
        actions.append('VKOSPI 80 초과 + 추세 훼손/NAV 손실: 강제 방어')
    elif pd.notna(v) and v > 70 and (trend_damaged or nav_pressure):
        signal_risk_target = PLAN_RULES_DEFENSE_RISK_BUDGET
        actions.append('VKOSPI 70 초과 + 추세 훼손/NAV 손실: 방어 전환')
    elif trend_damaged and nav_pressure:
        signal_risk_target = PLAN_RULES_DEFENSE_RISK_BUDGET
        actions.append('NAV 주의 + KOSPI200 20일선 이탈: 방어 전환')
    elif pd.notna(v) and v > 70:
        actions.append('VKOSPI 70 초과 단독: 위험자산 유지, 신규 고베타 증액 보류')
    elif monthly_check and expansion_allowed(signal_info, nav_state):
        signal_risk_target = PLAN_RULES_EXPANSION_RISK_BUDGET
        actions.append('월간 점검: VKOSPI 하락 안정 + 추세 유지, 위험자산 단계적 확대')

    target_risk = signal_risk_target if nav_state == '정상' else min(signal_risk_target, nav_target_risk)
    if risk_weight_of(w) > target_risk:
        w, a = reduce_risk_to_budget(w, target_risk, f'{nav_state}/Target Vol 위험예산 {target_risk:.0%}')
        actions.extend(a)
    elif monthly_check and risk_weight_of(w) < target_risk:
        w, a = increase_risk_to_budget(w, target_risk, f'월간 점검 Target Vol 위험예산 {target_risk:.0%}')
        actions.extend(a)

    if monthly_check:
        w, a = apply_sector_rotation(w, trade_dt, signal_info, nav_state)
        actions.extend(a)

    return normalize_weights(w), actions, target_risk

def plan_rules_nav(price_df, start=BACKTEST_START, end=BACKTEST_END, apply_costs=False,
                   portfolio_name='Strategy_TargetVol_Rule'):
    px = price_df.loc[start:end, PLAN_TICKERS].ffill().dropna()
    rebal_dates = set(month_first_dates(px.index))
    weights0 = BACKTEST_INITIAL_PORTFOLIO.set_index('Ticker')['Weight'].reindex(PLAN_TICKERS).fillna(0.0)
    weights0 = normalize_weights(weights0)

    holdings = None
    decision_holdings = None
    nav_rows = []
    schedule_rows = []
    peak_nav = 1.0

    for dt, row in px.iterrows():
        if holdings is None:
            nav = 1.0
            decision_nav = 1.0
            if apply_costs and INCLUDE_INITIAL_TRADING_COST:
                zero_value = pd.Series(0.0, index=PLAN_TICKERS)
                cost, turnover, cost_rate = estimate_rebalance_cost(zero_value, weights0, nav, PLAN_TICKERS)
                nav_after_cost = nav - cost
            else:
                cost, turnover, cost_rate = 0.0, 0.0, 0.0
                nav_after_cost = nav
            holdings = weights0 * nav_after_cost / row
            decision_holdings = weights0 * decision_nav / row
            nav = float((holdings * row).sum())
            decision_nav = float((decision_holdings * row).sum())
            signal_info = latest_signal_snapshot_for_trade(dt)
            schedule_rows.append({
                'Date': dt,
                'Rebalance_Type': 'Initial',
                'NAV_State': '정상',
                'NAV_Loss': decision_nav - 1.0,
                'MDD_ToDate': min(decision_nav / peak_nav - 1.0, 0.0),
                'Risk_Budget': risk_weight_of(weights0),
                'Portfolio_Risk_Weight': risk_weight_of(weights0),
                'Signal': '초기 65:35 진입',
                'Actions': '초기 65:35 진입',
                'Turnover': turnover,
                'TradingCost': cost if apply_costs else 0.0,
                'TradingCostRate': cost_rate if apply_costs else 0.0,
                **signal_info,
            })
            nav_rows.append((dt, nav))
            continue

        nav = float((holdings * row).sum())
        decision_nav = float((decision_holdings * row).sum())
        peak_nav = max(peak_nav, decision_nav)
        current_value = holdings * row
        decision_value = decision_holdings * row
        current_weights = normalize_weights(decision_value / decision_nav)
        signal_info = latest_signal_snapshot_for_trade(dt)
        nav_state, nav_target_risk, nav_loss, mdd = nav_defense_state(decision_nav, peak_nav)

        monthly_check = dt in rebal_dates
        v = signal_info.get('VKOSPI', np.nan)
        trend_damaged = signal_info.get('KOSPI200_Above_MA20') is False
        vkospi_defense_event = pd.notna(v) and v > 70 and (trend_damaged or nav_state != '정상')
        mid_event = vkospi_defense_event or nav_state in ['방어 전환', '강제 방어']
        risk_w = risk_weight_of(current_weights)
        domestic_sector_w = float(current_weights.reindex(PLAN_DOMESTIC_SECTOR_TICKERS).fillna(0.0).sum())
        single_hard_event = any(
            (weight > PLAN_RULES_SINGLE_ETF_HARD_BAND) or
            (weight > PLAN_RULES_SINGLE_ETF_SOFT_BAND and trend_damaged)
            for weight in current_weights
        )
        band_event = (
            risk_w > PLAN_RULES_RISK_HARD_BAND or
            (risk_w > PLAN_RULES_RISK_SOFT_BAND and trend_damaged) or
            domestic_sector_w > PLAN_RULES_DOMESTIC_SECTOR_HARD_BAND or
            (domestic_sector_w > PLAN_RULES_DOMESTIC_SECTOR_SOFT_BAND and trend_damaged) or
            single_hard_event
        )

        if monthly_check or mid_event or band_event:
            desired, actions, target_risk = desired_plan_weights(
                current_weights, dt, signal_info, nav_state, nav_target_risk, monthly_check
            )
            desired, band_actions = apply_limit_bands(desired, current_weights, signal_info)
            actions.extend(band_actions)
            if monthly_check:
                desired = monthly_band_rebalance(current_weights, desired)

            cost, turnover, cost_rate = estimate_rebalance_cost(current_value, desired, nav, PLAN_TICKERS)
            nav_after_cost = nav - cost if apply_costs else nav
            holdings = desired * nav_after_cost / row
            decision_holdings = desired * decision_nav / row
            nav = float((holdings * row).sum())
            current_weights = desired
            schedule_rows.append({
                'Date': dt,
                'Rebalance_Type': 'Monthly' if monthly_check else ('MidRisk' if mid_event else 'Band'),
                'NAV_State': nav_state,
                'NAV_Loss': nav_loss,
                'MDD_ToDate': mdd,
                'Risk_Budget': target_risk,
                'Portfolio_Risk_Weight': risk_weight_of(current_weights),
                'Signal': '; '.join(actions) if actions else '밴드 내 유지',
                'Actions': '; '.join(actions) if actions else '밴드 내 유지',
                'Turnover': turnover,
                'TradingCost': cost if apply_costs else 0.0,
                'TradingCostRate': cost_rate if apply_costs else 0.0,
                **signal_info,
            })

        nav_rows.append((dt, nav))

    nav = pd.Series(dict(nav_rows), name=portfolio_name).sort_index()
    schedule = pd.DataFrame(schedule_rows).drop_duplicates('Date').set_index('Date')
    return nav, schedule

static_weights = BACKTEST_INITIAL_PORTFOLIO.set_index('Ticker')['Weight']

# 계획서 규칙형: 월중 조정은 VKOSPI 고변동성+추세 훼손/NAV/MDD 또는 완화 밴드의 hard 조건에서만 발생한다.
strategy_tv_gross, targetvol_schedule_gross = plan_rules_nav(
    price, apply_costs=False, portfolio_name='Strategy_TargetVol_Rule'
)
strategy_tv_net, targetvol_schedule = plan_rules_nav(
    price, apply_costs=True, portfolio_name='Strategy_TargetVol_Rule'
)

bench_65_35_gross, costs_bench_gross = fixed_weight_nav(
    price, pd.Series({'A102110': 0.65, 'A459580': 0.35}), rebalance_monthly=False, apply_costs=False, portfolio_name='Benchmark_65_35_TIGER200_CD'
)
bench_65_35_net, costs_bench_net = fixed_weight_nav(
    price, pd.Series({'A102110': 0.65, 'A459580': 0.35}), rebalance_monthly=False, apply_costs=True, portfolio_name='Benchmark_65_35_TIGER200_CD'
)

static_no_rebal_gross, costs_no_rebal_gross = fixed_weight_nav(
    price, static_weights, rebalance_monthly=False, apply_costs=False, portfolio_name='Static_65_35_NoRebalance'
)
static_no_rebal_net, costs_no_rebal_net = fixed_weight_nav(
    price, static_weights, rebalance_monthly=False, apply_costs=True, portfolio_name='Static_65_35_NoRebalance'
)

tiger200_100_gross = normalized_series(price['A102110'], name='TIGER200_100')
tiger200_100_net, costs_tiger200_100_net = fixed_weight_nav(
    price, pd.Series({'A102110': 1.0}), rebalance_monthly=False, apply_costs=True, portfolio_name='TIGER200_100'
)
kospi200_100 = normalized_series(market_index['코스피 200'], name='KOSPI200_100')

nav_gross = pd.concat([
    strategy_tv_gross,
    bench_65_35_gross,
    tiger200_100_gross,
    kospi200_100,
    static_no_rebal_gross,
], axis=1).dropna()

# KOSPI200은 지수 벤치마크이므로 비용을 적용하지 않는다. TIGER 200 단일 ETF는 최초 매수 수수료를 반영한다.
nav_net = pd.concat([
    strategy_tv_net,
    bench_65_35_net,
    tiger200_100_net,
    kospi200_100,
    static_no_rebal_net,
], axis=1).dropna()

cost_events = pd.concat([
    targetvol_schedule.reset_index()[['Date', 'Turnover', 'TradingCost', 'TradingCostRate']].assign(Portfolio='Strategy_TargetVol_Rule'),
    costs_bench_net,
    costs_tiger200_100_net,
    costs_no_rebal_net,
], ignore_index=True, sort=False)

PORTFOLIO_LABELS = {
    'Strategy_TargetVol_Rule': '계획서형 Target Vol Overlay 규칙 적용',
    'Static_65_35_NoRebalance': '초기 65:35 포트폴리오 (리밸런싱 없음)',
    'Benchmark_65_35_TIGER200_CD': 'TIGER 200 65% + CD금리 35% (리밸런싱 없음)',
    'TIGER200_100': 'TIGER 200 100%',
    'KOSPI200_100': 'KOSPI200 100%',
}
REQUESTED_COLUMNS = ['Portfolio', 'TotalReturn', 'CAGR', 'Vol', 'Sharpe', 'Sortino', 'MDD', 'Calmar']


def max_drawdown(nav_series, initial_base=1.0):
    # 최초 매수 수수료가 첫 NAV에 반영되면 시작 원금 1.0 대비 낙폭도 MDD에 포함한다.
    running_peak = nav_series.cummax().clip(lower=initial_base)
    dd = nav_series / running_peak - 1
    return float(dd.min())


def period_return(nav_series, periods):
    s = nav_series.dropna()
    if len(s) <= periods:
        return np.nan
    return float(s.iloc[-1] / s.iloc[-periods - 1] - 1)


def ytd_return(nav_series):
    s = nav_series.dropna()
    y = s.loc[s.index >= pd.Timestamp('2026-01-01')]
    if len(y) < 2:
        return np.nan
    return float(y.iloc[-1] / y.iloc[0] - 1)


def downside_deviation(daily_returns, mar_daily):
    downside = np.minimum(daily_returns - mar_daily, 0.0)
    if len(downside) == 0:
        return np.nan
    return float(np.sqrt(np.mean(np.square(downside))) * np.sqrt(TRADING_DAYS))


def perf_stats(nav_df, labels=None, scenario=None, rf_ann_used=rf_ann, mar_ann_used=mar_ann, mar_daily_used=mar_daily):
    labels = labels or PORTFOLIO_LABELS
    rows = []
    for col in nav_df.columns:
        s = nav_df[col].dropna()
        r = s.pct_change().dropna()
        total = float(s.iloc[-1] / 1.0 - 1)
        cagr = float((1 + total) ** (TRADING_DAYS / len(r)) - 1) if len(r) else np.nan
        ann_vol = float(r.std() * np.sqrt(TRADING_DAYS)) if len(r) else np.nan
        sharpe = float((cagr - rf_ann_used) / ann_vol) if ann_vol and ann_vol > 0 else np.nan
        dd = max_drawdown(s)
        calmar = float(cagr / abs(dd)) if dd < 0 else np.nan
        down_dev = downside_deviation(r, mar_daily_used)
        sortino = float((cagr - mar_ann_used) / down_dev) if down_dev and down_dev > 0 else np.nan
        rows.append({
            'Scenario': scenario,
            'Portfolio_ID': col,
            'Portfolio': labels.get(col, col),
            'Start': s.index[0].date(),
            'End': s.index[-1].date(),
            'Obs': len(s),
            'TotalReturn': total,
            'CAGR': cagr,
            'Vol': ann_vol,
            'Sharpe': sharpe,
            'Sortino': sortino,
            'MDD': dd,
            'Calmar': calmar,
            'Ret_1M': period_return(s, 21),
            'Ret_3M': period_return(s, 63),
            'Ret_6M': period_return(s, 126),
            'YTD_2026': ytd_return(s),
            'Final_NAV': float(s.iloc[-1]),
        })
    return pd.DataFrame(rows)


def format_metric_table(df):
    out = df[REQUESTED_COLUMNS].copy()
    for c in ['TotalReturn', 'CAGR', 'Vol', 'MDD']:
        out[c] = out[c].map(lambda x: '' if pd.isna(x) else f'{x:.2%}')
    for c in ['Sharpe', 'Sortino', 'Calmar']:
        out[c] = out[c].map(lambda x: '' if pd.isna(x) else f'{x:.2f}')
    return out


performance_gross = perf_stats(nav_gross, scenario='비용 반영 전')
performance_net = perf_stats(nav_net, scenario='비용 반영 후')

portfolio_order = list(PORTFOLIO_LABELS.values())
result_table_gross = performance_gross.set_index('Portfolio').loc[portfolio_order].reset_index()[REQUESTED_COLUMNS]
result_table_net = performance_net.set_index('Portfolio').loc[portfolio_order].reset_index()[REQUESTED_COLUMNS]
result_table_gross_fmt = format_metric_table(result_table_gross)
result_table_net_fmt = format_metric_table(result_table_net)

# 기존 저장 로직과 호환되도록 대표 성과표 변수도 유지한다.
performance_summary = performance_net.copy()
trailing_returns = performance_gross[['Portfolio', 'Ret_1M', 'Ret_3M', 'Ret_6M', 'YTD_2026', 'TotalReturn']].copy()
nav = nav_net.copy()


def build_effect_decomposition(perf_df, scenario):
    p = perf_df.set_index('Portfolio')
    pairs = [
        ('분산/섹터선정 효과', 'TIGER 200 65% + CD금리 35% (리밸런싱 없음)', '초기 65:35 포트폴리오 (리밸런싱 없음)'),
        ('Target Vol/운용규칙 효과', '초기 65:35 포트폴리오 (리밸런싱 없음)', '계획서형 Target Vol Overlay 규칙 적용'),
    ]
    rows = []
    for effect, from_name, to_name in pairs:
        row = {'Scenario': scenario, 'Effect': effect, 'From': from_name, 'To': to_name}
        for metric in ['Vol', 'MDD', 'Sharpe', 'Sortino']:
            row[f'{metric}_Diff'] = p.loc[to_name, metric] - p.loc[from_name, metric]
        rows.append(row)
    return pd.DataFrame(rows)


effect_decomposition = pd.concat([
    build_effect_decomposition(performance_gross, '비용 반영 전'),
    build_effect_decomposition(performance_net, '비용 반영 후'),
], ignore_index=True)

def format_effect_table(df):
    out = df.copy()
    for c in ['Vol_Diff', 'MDD_Diff']:
        out[c] = out[c].map(lambda x: '' if pd.isna(x) else f'{x:+.2%}')
    for c in ['Sharpe_Diff', 'Sortino_Diff']:
        out[c] = out[c].map(lambda x: '' if pd.isna(x) else f'{x:+.2f}')
    return out


effect_decomposition_fmt = format_effect_table(effect_decomposition)

lookahead_audit = targetvol_schedule.reset_index()[[
    'Date', 'Signal_Date', 'VKOSPI', 'VKOSPI_20D_MA', 'VKOSPI_60D_MA',
    'KOSPI200_Close', 'KOSPI200_MA20', 'KOSPI200_MA60',
    'KOSPI200_Above_MA20', 'KOSPI200_Above_MA60', 'Risk_Budget', 'Signal',
    'Turnover', 'TradingCostRate'
]].copy()
lookahead_audit['Execution_Assumption'] = SIGNAL_EXECUTION_ASSUMPTION

backtest_footnotes = pd.DataFrame([
    {'Item': '백테스트 기간', 'Note': f'{BACKTEST_START.date()} ~ {BACKTEST_END.date()} ({len(nav_gross)}거래일). 2025년 5월 말 정보로 구성한 포트폴리오를 다음 영업일부터 적용한 약 12개월 상승장 구간.'},
    {'Item': 'Rf/Sharpe', 'Note': f'{rf_source}. 적용 Rf={rf_ann:.2%}.'},
    {'Item': 'Sortino MAR', 'Note': f'{mar_source}. 적용 MAR={mar_ann:.2%}, 하방편차는 일별 수익률이 일별 MAR을 밑돈 부분만 사용.'},
    {'Item': '리밸런싱 가격', 'Note': f'{REBALANCE_EXECUTION_ASSUMPTION}. {PRICE_EXECUTION_PROXY}.'},
    {'Item': 'VKOSPI 신호', 'Note': '체결일 Date에는 Date 직전 Signal_Date의 종가 신호만 사용한다. 20일/60일 이동평균도 Signal_Date까지의 trailing 데이터로 계산한다.'},
    {'Item': '계획서 규칙형 Target Vol', 'Note': '월중 조정은 VKOSPI 70 초과가 KOSPI200 추세 훼손 또는 NAV/MDD 손실과 결합될 때, NAV -5%/MDD -6% 방어 전환, NAV -7%/MDD -8% 강제 방어, 또는 완화 밴드 hard 조건에서만 수행. 섹터 로테이션은 월간 점검일에만 실행.'},
    {'Item': '위험확대 70% 조건', 'Note': 'Target Vol 효과 분리표에서는 65% 초과 위험확대를 OFF로 두어 방어 overlay만 검증한다. 별도 민감도 분석에서 켤 경우 VKOSPI 60 이하, KOSPI200 20/60일선 상회, VKOSPI < 20일 평균 < 60일 평균일 때만 허용.'},
    {'Item': '비교 벤치마크', 'Note': '초기 65:35 포트폴리오와 TIGER 200 65%+CD 35% 벤치마크는 월별 목표비중 복원이 아니라 최초 비중 매수 후 리밸런싱 없이 보유하는 방식으로 계산.'},
    {'Item': '거래비용', 'Note': f'대회 기준 국내 ETF 거래수수료 {COMMISSION_RATE:.2%}, 제세금 {TAX_RATE:.2%}. 슬리피지는 공식 NAV 비용으로 차감하지 않음. 최초 진입 및 이후 회전 금액에 비례해 차감.'},
    {'Item': '비용 전/후 비교', 'Note': '동일 신호·동일 거래 경로에서 비용만 분리해 비교한다. KOSPI200 지수 벤치마크는 비용 미적용, TIGER 200 ETF 벤치마크는 최초 매수 수수료 반영.'},
    {'Item': '회전율 공식', 'Note': '회전율(%) = 총 매매금액 / 평균자산 × 100. 대회 시작일부터 5영업일 이내 80% 이상, 월 회전율 10% 이상, 3개월 누적 기준 없음.'},
    {'Item': 'KOSPI200 100%', 'Note': '지수 벤치마크이므로 거래비용을 적용하지 않음. TIGER 200 100% ETF 벤치마크에는 최초 매수 수수료를 반영.'},
])

data_coverage = pd.DataFrame({
    'Ticker': first_dates.index,
    'Name': [code_names.get(t) for t in first_dates.index],
    'First_Date': first_dates.values,
    'Last_Date': last_dates.values,
}).sort_values('First_Date', ascending=False)

period_limitations = pd.DataFrame([
    {'Check': '주 백테스트 구간', 'Result': f'{BACKTEST_START.date()} ~ {BACKTEST_END.date()}. 2025년 5월 말까지의 정보로 정한 백테스트 포트폴리오를 다음 영업일부터 적용.'},
    {'Check': '원 포트폴리오 장기 검증 제약', 'Result': f'백테스트 전용 초기 포트폴리오의 공통 시작일은 {first_dates.max().date()}이며, 시작 요청일은 2025년 5월 말 정보가 반영되는 {BACKTEST_START_REQUESTED.date()}로 설정.'},
    {'Check': '다른 국면 검증 한계', 'Result': '2022 하락장, 2020 변동성 급등 구간은 현재 편입 ETF 일부가 상장 전이라 원 포트폴리오 그대로는 재현 불가. 대체 ETF 프록시 장기 백테스트는 제출본의 주 성과표에서 제외.'},
])


print('\n비용 반영 전 성과표')
print(result_table_gross_fmt.to_string(index=False))
print('\n대회 거래비용 반영 후 성과표')
print(result_table_net_fmt.to_string(index=False))
print('\n효과 분해')
print(effect_decomposition_fmt.to_string(index=False))

# 백테스트 결과 저장
with pd.ExcelWriter(BACKTEST_OUTPUT_PATH, engine='openpyxl') as writer:
    performance_gross.to_excel(writer, sheet_name='Perf_Gross', index=False)
    performance_net.to_excel(writer, sheet_name='Perf_Net', index=False)
    result_table_gross.to_excel(writer, sheet_name='Table_Gross', index=False)
    result_table_net.to_excel(writer, sheet_name='Table_Net', index=False)
    effect_decomposition.to_excel(writer, sheet_name='Effects', index=False)
    trailing_returns.to_excel(writer, sheet_name='Trailing_Returns', index=False)
    nav_gross.to_excel(writer, sheet_name='Daily_NAV_Gross')
    nav_net.to_excel(writer, sheet_name='Daily_NAV_Net')
    targetvol_schedule.to_excel(writer, sheet_name='TargetVol_Schedule')
    targetvol_schedule.to_excel(writer, sheet_name='PlanRule_Schedule')
    lookahead_audit.to_excel(writer, sheet_name='Lookahead_Audit', index=False)
    backtest_footnotes.to_excel(writer, sheet_name='Footnotes', index=False)
    period_limitations.to_excel(writer, sheet_name='Period_Limits', index=False)
    cost_events.to_excel(writer, sheet_name='Cost_Events', index=False)
    portfolio_meta.to_excel(writer, sheet_name='Initial_Portfolio', index=False)
    data_coverage.to_excel(writer, sheet_name='Data_Coverage', index=False)

print(f'저장 완료: {BACKTEST_OUTPUT_PATH}')
