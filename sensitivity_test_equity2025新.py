import pandas as pd
import numpy as np
from tabulate import tabulate
from run_backtest_equity2025新 import Backtester, clean_data, calculate_metrics

def main():
    data_file = '個股合-1.xlsx'
    prices, code_to_name = clean_data(data_file)
    bt = Backtester(prices, code_to_name)

    base_sma = 87
    base_roc = 54
    sl = 0.09
    reb = 6

    sma_tests = [84, 85, 87, 89, 90]
    roc_tests = [51, 52, 54, 56, 57]

    print("### SMA Sensitivity Analysis (ROC=54)")
    sma_results = []
    for s in sma_tests:
        eq, _, _ = bt.run(s, base_roc, sl, reb)
        cagr, mdd, calmar, _ = calculate_metrics(eq)
        sma_results.append([s, f"{cagr:.2%}", f"{mdd:.2%}", f"{calmar:.2f}"])

    print(tabulate(sma_results, headers=["SMA", "CAGR", "MaxDD", "Calmar"], tablefmt="pipe"))
    print("\n")

    print("### ROC Sensitivity Analysis (SMA=87)")
    roc_results = []
    for r in roc_tests:
        eq, _, _ = bt.run(base_sma, r, sl, reb)
        cagr, mdd, calmar, _ = calculate_metrics(eq)
        roc_results.append([r, f"{cagr:.2%}", f"{mdd:.2%}", f"{calmar:.2f}"])

    print(tabulate(roc_results, headers=["ROC", "CAGR", "MaxDD", "Calmar"], tablefmt="pipe"))

if __name__ == "__main__":
    main()
