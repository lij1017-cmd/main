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
    TRADING_CAP = 30000000
    AUTH_CAP = 150000000
    bt = BacktesterVol(prices, volumes, code_to_name, trading_capital=TRADING_CAP, authorized_capital=AUTH_CAP)

    # 固定參數
    SMA_PERIOD = 303
    ROC_PERIOD = 10
    REBALANCE = 9
    VOL_PERIOD = 15
    VOL_MULTIPLIER = 2.7
    BREADTH_THRESHOLD = 0.42
    MKT_SMA = 14

    TRAIN_START = '2019-01-01'
    TRAIN_END = '2023-12-31'

    results = []

    # 測試指定區間
    windows = [200, 240, 290, 350]

    print(f"Starting Sensitivity Analysis for Breadth Window ({TRAIN_START} to {TRAIN_END})...")

    for bw in windows:
        print(f"Testing Breadth Window: {bw}")
        eq_curve, trades, trades2, details = bt.run(
            sma_period=SMA_PERIOD,
            roc_period=ROC_PERIOD,
            vol_period=VOL_PERIOD,
            vol_multiplier=VOL_MULTIPLIER,
            rebalance_interval=REBALANCE,
            breadth_threshold=BREADTH_THRESHOLD,
            mkt_sma_window=MKT_SMA,
            breadth_window=bw,
            start_date=TRAIN_START,
            end_date=TRAIN_END
        )

        metrics = calculate_metrics_dual(eq_curve, TRADING_CAP, AUTH_CAP)
        results.append({
            'BreadthWindow': bw,
            'Trading_CAGR': metrics['Trading CAGR'],
            'Auth_CAGR': metrics['Authorized CAGR'],
            'Std_MaxDD': metrics['Standard MaxDD'],
            'Fixed_MaxDD': metrics['Fixed Base MaxDD'],
            'Calmar': metrics['Trading Calmar']
        })

    results_df = pd.DataFrame(results)
    results_df.to_csv('breadth_sensitivity_adj1.csv', index=False)
    print("\nResults saved to breadth_sensitivity_adj1.csv")
    print(results_df.to_string(index=False))

if __name__ == "__main__":
    main()
