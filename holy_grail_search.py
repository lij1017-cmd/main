import pandas as pd
import numpy as np
from backtest_v2 import clean_data, BacktesterV2, calculate_metrics

def main():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterV2(prices, volumes, code_to_name)
    p1_s, p1_e = pd.to_datetime('2019-01-01'), pd.to_datetime('2023-12-31')
    p2_s, p2_e = pd.to_datetime('2024-01-01'), pd.to_datetime('2025-12-31')

    # Targeted Search for the "Holy Grail" parameters
    best_min_calmar = -1
    best_p = None

    # Range 1: Exploring MA Stop 60 neighborhood
    print("Scanning Range 1 (MA Stop 60)...")
    for sma in range(70, 111, 2):
        for roc in range(5, 31, 1):
            for reb in [11, 12, 13, 14, 15]:
                # We fix sl=0.1 as it didn't seem to matter for MA Stop
                eq, trades, _, _, _ = bt.run(sma, roc, 0.1, reb, 'ma', 60)
                eq1 = eq[(eq['日期'] >= p1_s) & (eq['日期'] <= p1_e)]
                if eq1.empty: continue
                _, mdd1, c1, _ = calculate_metrics(eq1)

                eq2 = eq[(eq['日期'] >= p2_s) & (eq['日期'] <= p2_e)]
                if eq2.empty: continue
                _, mdd2, c2, _ = calculate_metrics(eq2)

                if mdd1 < -0.25 or mdd2 < -0.25: continue

                min_c = min(c1, c2)
                if min_c > best_min_calmar:
                    best_min_calmar = min_c
                    best_p = (sma, roc, 0.1, reb, 'ma', 60)
                    print(f"New Best: {best_p} -> Min Calmar: {best_min_calmar:.2f} (P1: {c1:.2f}, P2: {c2:.2f})")
                    if c1 > 3.0 and c2 > 3.0:
                        print("!!! GOAL MET !!!")
                        return

    # Range 2: Exploring Peak Stop neighborhood
    print("Scanning Range 2 (Peak Stop)...")
    for sma in range(100, 141, 2):
        for roc in range(15, 41, 1):
            for sl in [0.08, 0.09, 0.095]: # Peak stop < 10%
                for reb in [11, 12, 13, 14, 15]:
                    eq, trades, _, _, _ = bt.run(sma, roc, sl, reb, 'peak', 5)
                    eq1 = eq[(eq['日期'] >= p1_s) & (eq['日期'] <= p1_e)]
                    if eq1.empty: continue
                    _, mdd1, c1, _ = calculate_metrics(eq1)

                    eq2 = eq[(eq['日期'] >= p2_s) & (eq['日期'] <= p2_e)]
                    if eq2.empty: continue
                    _, mdd2, c2, _ = calculate_metrics(eq2)

                    if mdd1 < -0.25 or mdd2 < -0.25: continue

                    min_c = min(c1, c2)
                    if min_c > best_min_calmar:
                        best_min_calmar = min_c
                        best_p = (sma, roc, sl, reb, 'peak', 5)
                        print(f"New Best: {best_p} -> Min Calmar: {best_min_calmar:.2f} (P1: {c1:.2f}, P2: {c2:.2f})")
                        if c1 > 3.0 and c2 > 3.0:
                            print("!!! GOAL MET !!!")
                            return

if __name__ == "__main__":
    main()
