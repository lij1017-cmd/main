import pandas as pd
from backtest_atr import clean_data, BacktesterATR, calculate_metrics

def verify():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterATR(prices, volumes, code_to_name)

    # Candidate 1: ROC=10, REB=9, ATR_P=15, ATR_M=4.3, MKT_T=0.42, MKT_S=14
    p = {'roc': 10, 'reb': 9, 'atr_p': 15, 'atr_m': 4.3, 'mkt_t': 0.42, 'mkt_s': 14}

    equity, trades, holdings, trades2, daily = bt.run(
        sma_period=303,
        roc_period=p['roc'],
        stop_loss_type='atr',
        atr_period=p['atr_p'],
        atr_multiplier=p['atr_m'],
        rebalance_interval=p['reb'],
        use_market_filter=True,
        breadth_threshold=p['mkt_t'],
        mkt_sma_window=p['mkt_s'],
        breadth_window=290
    )

    cagr, mdd, calmar, total_ret = calculate_metrics(equity)
    print(f"--- 全期間績效 (2019-2025) ---")
    print(f"CAGR: {cagr:.2%}")
    print(f"MaxDD: {mdd:.2%}")
    print(f"Calmar: {calmar:.2f}")
    print(f"Total Return: {total_ret:.2%}")

    print(f"\n--- 年度績效 ---")
    equity['Year'] = equity['日期'].dt.year
    for year, group in equity.groupby('Year'):
        y_cagr, y_mdd, y_calmar, y_ret = calculate_metrics(group)
        print(f"{year}: CAGR {y_cagr:7.2%}, MaxDD {y_mdd:7.2%}, Calmar {y_calmar:5.2f}, Return {y_ret:7.2%}")

if __name__ == "__main__":
    verify()
