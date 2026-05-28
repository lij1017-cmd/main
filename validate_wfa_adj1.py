import pandas as pd
import numpy as np
from backtest_vol import BacktesterVol, calculate_metrics_dual, clean_data
import os

def main():
    filepath = '樣本集-1.xlsx'
    if not os.path.exists(filepath):
        print(f"Error: {filepath} not found.")
        return

    prices, volumes, code_to_name = clean_data(filepath)
    TRADING_CAP = 30000000
    AUTH_CAP = 150000000
    bt = BacktesterVol(prices, volumes, code_to_name, trading_capital=TRADING_CAP, authorized_capital=AUTH_CAP)

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

        metrics = calculate_metrics_dual(eq_curve, TRADING_CAP, AUTH_CAP)
        results.append({
            'Period': name,
            'Start': start,
            'End': end,
            'Trading_CAGR': metrics['Trading CAGR'],
            'Auth_CAGR': metrics['Authorized CAGR'],
            'Std_MaxDD': metrics['Standard MaxDD'],
            'Fixed_MaxDD': metrics['Fixed Base MaxDD'],
            'Calmar': metrics['Trading Calmar']
        })

    results_df = pd.DataFrame(results)
    print("\nIS/OOS Validation Results:")
    print(results_df.to_string(index=False))
    results_df.to_csv('validate_is_oos_adj1.csv', index=False)

    # WFA
    wfa_periods = [
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
    wfa_results = []
    for start, end in wfa_periods:
        print(f"Running WFA Period: {start} to {end}")
        eq_curve, _, _, _ = bt.run(SMA_PERIOD, ROC_PERIOD, start_date=start, end_date=end)
        metrics = calculate_metrics_dual(eq_curve, TRADING_CAP, AUTH_CAP)
        wfa_results.append({
            'Period': f"{start} - {end}",
            'Trading_CAGR': metrics['Trading CAGR'],
            'Auth_CAGR': metrics['Authorized CAGR'],
            'Std_MaxDD': metrics['Standard MaxDD'],
            'Fixed_MaxDD': metrics['Fixed Base MaxDD'],
            'Calmar': metrics['Trading Calmar']
        })
    wfa_df = pd.DataFrame(wfa_results)
    print("\nWFA Results:")
    print(wfa_df.to_string(index=False))
    wfa_df.to_csv('run_wfa_adj1.csv', index=False)

if __name__ == "__main__":
    main()
