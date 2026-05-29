import pandas as pd
import numpy as np
from backtest_vol import BacktesterVol, calculate_metrics_dual, clean_data

def main():
    filepath = '樣本集-1.xlsx'
    prices, volumes, code_to_name = clean_data(filepath)
    TRADING_CAP = 30000000
    AUTH_CAP = 150000000
    bt = BacktesterVol(prices, volumes, code_to_name, trading_capital=TRADING_CAP, authorized_capital=AUTH_CAP)

    params = [
        (303, 10, 2.7, 290),
        (303, 14, 2.7, 290)
    ]

    for sma, roc, mult, b_win in params:
        print(f"\n--- Testing SMA {sma}, ROC {roc}, Mult {mult}, Breadth {b_win} ---")
        eq, _, _, _ = bt.run(sma, roc, vol_multiplier=mult, breadth_window=b_win, use_breadth_weight=True)
        m = calculate_metrics_dual(eq, TRADING_CAP, AUTH_CAP)
        print(f"Full CAGR (Trading): {m['Trading CAGR']:.2%}")
        print(f"Full MaxDD (Std): {m['Standard MaxDD']:.2%}")
        print(f"Calmar: {m['Trading Calmar']:.2f}")
        print(f"2022 Return: {m['Yearly Performance'].loc[2022, '年度報酬率']:.2%}")

        # WFA Check
        wfa_periods = [
            ('2024-06-01', '2025-12-31'), ('2024-01-02', '2025-05-31'), ('2023-01-02', '2024-12-31'),
            ('2022-01-02', '2024-05-31'), ('2021-06-01', '2023-12-31'), ('2021-01-02', '2023-05-31'),
            ('2020-01-02', '2022-12-31'), ('2019-06-01', '2022-05-31'), ('2019-01-02', '2021-12-31'),
        ]
        cagrs = []
        for s, e in wfa_periods:
            ew, _, _, _ = bt.run(sma, roc, vol_multiplier=mult, breadth_window=b_win, start_date=s, end_date=e, use_breadth_weight=True)
            mw = calculate_metrics_dual(ew, TRADING_CAP, AUTH_CAP)
            cagrs.append(mw['Trading CAGR'])

        print(f"WFA CAGR Min: {min(cagrs):.2%}")
        print(f"WFA CAGR Max: {max(cagrs):.2%}")
        print(f"WFA CAGR Mean: {np.mean(cagrs):.2%}")
        print(f"WFA CAGR Std: {np.std(cagrs):.2%}")

if __name__ == "__main__":
    main()
