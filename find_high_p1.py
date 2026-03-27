import pandas as pd
import numpy as np
from backtest_v2 import clean_data, BacktesterV2, calculate_metrics

def main():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterV2(prices, volumes, code_to_name)
    p1_s, p1_e = pd.to_datetime('2019-01-01'), pd.to_datetime('2023-12-31')
    p2_s, p2_e = pd.to_datetime('2024-01-01'), pd.to_datetime('2025-12-31')

    # Look for parameters where P1 is extremely high
    configs = []
    for sma in range(30, 151, 20):
        for roc in range(30, 151, 20):
            for reb in [5, 7, 9]:
                configs.append((sma, roc, 0.1, reb, 'peak', 10))

    best_min_calmar = -1
    best_p = None

    for p in configs:
        eq, trades, _, _, _ = bt.run(*p)
        eq1 = eq[(eq['日期'] >= p1_s) & (eq['日期'] <= p1_e)]
        if eq1.empty: continue
        _, mdd1, c1, _ = calculate_metrics(eq1)

        if c1 > 3.5:
            eq2 = eq[(eq['日期'] >= p2_s) & (eq['日期'] <= p2_e)]
            if eq2.empty: continue
            _, mdd2, c2, _ = calculate_metrics(eq2)
            min_calmar = min(c1, c2)
            if mdd1 < -0.25 or mdd2 < -0.25: min_calmar = -1
            if min_calmar > best_min_calmar:
                best_min_calmar = min_calmar
                best_p = p
                print(f"High P1 Candidate: {best_p} -> Min Calmar: {best_min_calmar:.2f} (P1: {c1:.2f}, P2: {c2:.2f})")

if __name__ == "__main__":
    main()
