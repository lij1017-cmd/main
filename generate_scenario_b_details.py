import pandas as pd
import numpy as np
from backtest_atr import clean_data, BacktesterATR, calculate_metrics

def generate_details():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterATR(prices, volumes, code_to_name)

    # Scenario B: ROC=10, SMA=303, ATR_P=15, ATR_M=4.3
    p = {'roc': 10, 'reb': 9, 'atr_p': 15, 'atr_m': 4.3, 'mkt_t': 0.42, 'mkt_s': 14}

    equity, trades, holdings, trades2, daily = bt.run(
        sma_period=303,
        roc_period=p['roc'],
        stop_loss_type='atr',
        atr_period=p['atr_p'],
        atr_multiplier=p['atr_m'],
        rebalance_interval=9,
        use_market_filter=True,
        breadth_threshold=p['mkt_t'],
        mkt_sma_window=p['mkt_s'],
        breadth_window=290
    )

    # 1. 年度績效
    equity['Year'] = equity['日期'].dt.year
    annual_metrics = []
    for year, group in equity.groupby('Year'):
        y_cagr, y_mdd, y_calmar, y_ret = calculate_metrics(group)
        annual_metrics.append({
            'Year': year, 'CAGR': y_cagr, 'MaxDD': y_mdd, 'Calmar': y_calmar, 'Return': y_ret
        })
    df_annual = pd.DataFrame(annual_metrics)
    print("--- 2019-2025 年度績效 (Scenario B) ---")
    print(df_annual.to_string(index=False))

    # 2. 尋找 ATR 使用案例 (尋找賣出原因包含 ATR 的交易)
    atr_trades = trades2[trades2['賣出原因'].str.contains('ATR', na=False)]
    if not atr_trades.empty:
        print("\n--- ATR 停損案例 (精選) ---")
        # 挑選一筆具有代表性的，例如 2022 年或 2021 年的
        example = atr_trades.sort_values('損益', ascending=False).iloc[0] # 獲利回吐出場的案例通常很有說服力
        print(example)

    # 3. 產出高原參數表 (鄰近 ATR 參數)
    periods = [10, 15, 20, 25]
    multipliers = [3.5, 4.0, 4.3, 4.5, 5.0]

    plateau = []
    for ap in periods:
        for am in multipliers:
            e, _, _, _, _ = bt.run(
                sma_period=303, roc_period=10, stop_loss_type='atr',
                atr_period=ap, atr_multiplier=am,
                rebalance_interval=9, use_market_filter=True, breadth_threshold=0.42, mkt_sma_window=14
            )
            c, m, ca, r = calculate_metrics(e)
            plateau.append({'ATR_P': ap, 'ATR_M': am, 'CAGR': c, 'MaxDD': m, 'Calmar': ca})

    df_plateau = pd.DataFrame(plateau)
    pivot_calmar = df_plateau.pivot(index='ATR_P', columns='ATR_M', values='Calmar')
    print("\n--- ATR 參數高原表 (Calmar Ratio) ---")
    print(pivot_calmar)

    pivot_cagr = df_plateau.pivot(index='ATR_P', columns='ATR_M', values='CAGR')
    print("\n--- ATR 參數高原表 (CAGR) ---")
    print(pivot_cagr)

if __name__ == "__main__":
    generate_details()
