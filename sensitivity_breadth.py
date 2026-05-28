import pandas as pd
import numpy as np
from backtest_vol import BacktesterVol, calculate_metrics, clean_data
import os

def main():
    filepath = '樣本集-1.xlsx'
    if not os.path.exists(filepath):
        print(f"Error: {filepath} not found.")
        return

    print("Loading data...")
    prices, volumes, code_to_name = clean_data(filepath)
    bt = BacktesterVol(prices, volumes, code_to_name)

    # 固定參數 (使用剛選出的 Multiplier 2.7)
    SMA_PERIOD = 303
    ROC_PERIOD = 10
    REBALANCE = 9
    VOL_PERIOD = 15
    VOL_MULTIPLIER = 2.7
    BREADTH_THRESHOLD = 0.42
    MKT_SMA = 14

    TRAIN_START = '2019-01-01'
    TRAIN_END = '2023-12-31'

    breadth_windows = [200, 240, 290, 350]
    results = []

    print(f"Starting Breadth Window Sensitivity Analysis ({TRAIN_START} to {TRAIN_END})...")

    for bw in breadth_windows:
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

        cagr, mdd, fmdd, calmar, total_ret = calculate_metrics(eq_curve)
        results.append({
            'BreadthWindow': bw,
            'CAGR': cagr,
            'MaxDD': mdd,
            'Calmar': calmar,
            'TotalReturn': total_ret
        })

    results_df = pd.DataFrame(results)
    print("\nBreadth Window Sensitivity Results (Training Set 2019-2023):")
    print(results_df.to_string(index=False))

    results_df.to_csv('breadth_sensitivity_results.csv', index=False)

if __name__ == "__main__":
    main()
