import pandas as pd
import numpy as np
from backtest_vol import BacktesterVol, calculate_metrics, clean_data
import os

def main():
    filepath = '樣本集-1.xlsx'
    if not os.path.exists(filepath):
        print(f"Error: {filepath} not found.")
        return

    prices, volumes, code_to_name = clean_data(filepath)
    bt = BacktesterVol(prices, volumes, code_to_name)

    # 選定的最優/穩健參數
    SMA_PERIOD = 303
    ROC_PERIOD = 10
    REBALANCE = 9
    VOL_PERIOD = 15
    VOL_MULTIPLIER = 2.7
    BREADTH_THRESHOLD = 0.42
    MKT_SMA = 14
    BREADTH_WINDOW = 290

    periods = [
        ('In-Sample (Train)', '2019-01-01', '2023-12-31'),
        ('Out-of-Sample (Test)', '2024-01-01', '2025-12-31'),
        ('Full Period', '2019-01-01', '2025-12-31')
    ]

    results = []

    for name, start, end in periods:
        print(f"Running {name}: {start} to {end}")
        eq_curve, trades, trades2, details = bt.run(
            sma_period=SMA_PERIOD,
            roc_period=ROC_PERIOD,
            vol_period=VOL_PERIOD,
            vol_multiplier=VOL_MULTIPLIER,
            rebalance_interval=REBALANCE,
            breadth_threshold=BREADTH_THRESHOLD,
            mkt_sma_window=MKT_SMA,
            breadth_window=BREADTH_WINDOW,
            start_date=start,
            end_date=end
        )

        cagr, mdd, fmdd, calmar, total_ret = calculate_metrics(eq_curve)
        results.append({
            'Period': name,
            'Start': start,
            'End': end,
            'CAGR': cagr,
            'MaxDD': mdd,
            'Calmar': calmar,
            'TotalReturn': total_ret
        })

    results_df = pd.DataFrame(results)
    print("\nIS/OOS Validation Results:")
    print(results_df.to_string(index=False))

    results_df.to_csv('is_oos_validation.csv', index=False)

if __name__ == "__main__":
    main()
