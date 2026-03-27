import pandas as pd
import numpy as np
from backtest_v2 import clean_data, BacktesterV2, calculate_metrics

def main():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterV2(prices, volumes, code_to_name)
    p1_s, p1_e = pd.to_datetime('2019-01-01'), pd.to_datetime('2023-12-31')
    p2_s, p2_e = pd.to_datetime('2024-01-01'), pd.to_datetime('2025-12-31')

    # Deep scan ROC for (80, ROC, 0.1, 13, 'ma', 60)
    best_min_calmar = -1
    best_p = None

    for roc in range(1, 151):
        eq, trades, _, _, _ = bt.run(80, roc, 0.1, 13, 'ma', 60)
        eq1 = eq[(eq['日期'] >= p1_s) & (eq['日期'] <= p1_e)]
        if eq1.empty: continue
        _, mdd1, c1, _ = calculate_metrics(eq1)
        eq2 = eq[(eq['日期'] >= p2_s) & (eq['日期'] <= p2_e)]
        if eq2.empty: continue
        _, mdd2, c2, _ = calculate_metrics(eq2)
        min_calmar = min(c1, c2)
        if mdd1 < -0.25 or mdd2 < -0.25: min_calmar = -1
        if min_calmar > best_min_calmar:
            best_min_calmar = min_calmar
            best_p = (80, roc, 0.1, 13, 'ma', 60)
            print(f"New Best: {best_p} -> Min Calmar: {best_min_calmar:.2f} (P1: {c1:.2f}, P2: {c2:.2f})")

if __name__ == "__main__":
    main()
