import pandas as pd
import numpy as np
from backtest_vol import BacktesterVol, calculate_metrics_dual, clean_data
import os

def main():
    filepath = '樣本集-1.xlsx'
    if not os.path.exists(filepath):
        print(f"Error: {filepath} not found.")
        return

    print("Loading data...")
    prices, volumes, code_to_name = clean_data(filepath)
    # 使用 30M 交易資金與 150M 授權資金
    TRADING_CAP = 30000000
    AUTH_CAP = 150000000
    bt = BacktesterVol(prices, volumes, code_to_name, trading_capital=TRADING_CAP, authorized_capital=AUTH_CAP)

    # 固定參數
    SMA_PERIOD = 303
    ROC_PERIOD = 10
    REBALANCE = 9
    VOL_PERIOD = 15
    BREADTH_THRESHOLD = 0.42
    MKT_SMA = 14
    BREADTH_WINDOW = 290

    TRAIN_START = '2019-01-01'
    TRAIN_END = '2023-12-31'

    results = []

    # 使用 0.2 步長
    multipliers = np.arange(1.5, 3.7, 0.2)

    print(f"Starting Grid Search for Multiplier on Training Set ({TRAIN_START} to {TRAIN_END})...")

    for m in multipliers:
        m = round(m, 1)
        print(f"Testing Multiplier: {m}")
        eq_curve, trades, trades2, details = bt.run(
            sma_period=SMA_PERIOD,
            roc_period=ROC_PERIOD,
            vol_period=VOL_PERIOD,
            vol_multiplier=m,
            rebalance_interval=REBALANCE,
            breadth_threshold=BREADTH_THRESHOLD,
            mkt_sma_window=MKT_SMA,
            breadth_window=BREADTH_WINDOW,
            start_date=TRAIN_START,
            end_date=TRAIN_END
        )

        metrics = calculate_metrics_dual(eq_curve, TRADING_CAP, AUTH_CAP)
        results.append({
            'Multiplier': m,
            'Trading_CAGR': metrics['Trading CAGR'],
            'Auth_CAGR': metrics['Authorized CAGR'],
            'Std_MaxDD': metrics['Standard MaxDD'],
            'Fixed_MaxDD': metrics['Fixed Base MaxDD'],
            'Calmar': metrics['Trading Calmar']
        })

    results_df = pd.DataFrame(results)
    results_df.to_csv('multiplier_plateau_adj1.csv', index=False)
    print("\nResults saved to multiplier_plateau_adj1.csv")
    print(results_df.to_string(index=False))

if __name__ == "__main__":
    main()
