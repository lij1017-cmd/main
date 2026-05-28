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

    windows = range(180, 410, 10)

    print(f"Starting Broad Grid Search for Breadth Window on Training Set ({TRAIN_START} to {TRAIN_END})...")

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

        cagr, mdd, fmdd, calmar, total_ret = calculate_metrics(eq_curve)
        results.append({
            'BreadthWindow': bw,
            'CAGR': cagr,
            'MaxDD': mdd,
            'Calmar': calmar,
            'TotalReturn': total_ret
        })

    results_df = pd.DataFrame(results)
    results_df.to_csv('breadth_plateau.csv', index=False)
    print("\nResults saved to breadth_plateau.csv")

if __name__ == "__main__":
    main()
