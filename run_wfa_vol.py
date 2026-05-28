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

    # 9-period WFA 區間
    periods = [
        ('2024-06-01', '2025-12-31'),
        ('2024-01-02', '2025-05-31'),
        ('2023-01-02', '2024-12-31'),
        ('2022-01-02', '2024-05-31'),
        ('2021-06-01', '2023-12-31'),
        ('2021-01-02', '2023-05-31'),
        ('2020-01-02', '2022-12-31'),
        ('2019-06-01', '2022-05-31'),
        ('2019-01-02', '2021-12-31'),
    ]

    results = []

    for start, end in periods:
        print(f"Running WFA Period: {start} to {end}")
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
            'Period': f"{start} - {end}",
            'CAGR': cagr,
            'MaxDD': mdd,
            'Calmar': calmar,
            'TotalReturn': total_ret
        })

    results_df = pd.DataFrame(results)
    print("\nWalk-Forward Analysis Results:")
    print(results_df.to_string(index=False))

    results_df.to_csv('wfa_results_vol.csv', index=False)

if __name__ == "__main__":
    main()
