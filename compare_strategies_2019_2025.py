import pandas as pd
import numpy as np
from tabulate import tabulate
from run_backtest_equity2025新_1 import Backtester, clean_data, calculate_metrics

def main():
    data_file = '個股合-1.xlsx'
    prices, code_to_name = clean_data(data_file)
    bt = Backtester(prices, code_to_name)

    # Strategy 1: Optimized for 6-day rebalance
    sma1, roc1, sl1, reb1 = 87, 54, 0.09, 6
    # Strategy 2: Optimized for 8-day rebalance
    sma2, roc2, sl2, reb2 = 27, 98, 0.075, 8

    # Run Strategy 1
    eq1, trades1, _ = bt.run(sma1, roc1, sl1, reb1)
    cagr1, mdd1, calmar1, ret1 = calculate_metrics(eq1)
    cost1 = trades1['買入手續費'].sum() + trades1['賣出手續費'].sum() + trades1['賣出交易稅'].sum()

    # Run Strategy 2
    eq2, trades2, _ = bt.run(sma2, roc2, sl2, reb2)
    cagr2, mdd2, calmar2, ret2 = calculate_metrics(eq2)
    cost2 = trades2['買入手續費'].sum() + trades2['賣出手續費'].sum() + trades2['賣出交易稅'].sum()

    results = [
        ["參數組合", "SMA:87, ROC:54, SL:9%", "SMA:27, ROC:98, SL:7.5%"],
        ["再平衡週期", "6日", "8日"],
        ["CAGR", f"{cagr1:.2%}", f"{cagr2:.2%}"],
        ["MaxDD", f"{mdd1:.2%}", f"{mdd2:.2%}"],
        ["Calmar Ratio", f"{calmar1:.2f}", f"{calmar2:.2f}"],
        ["總報酬率", f"{ret1:.2%}", f"{ret2:.2%}"],
        ["交易筆數", len(trades1), len(trades2)],
        ["交易成本總計", f"{int(cost1):,}", f"{int(cost2):,}"]
    ]

    # Transpose for horizontal display
    print(tabulate(results, tablefmt="pipe"))

if __name__ == "__main__":
    main()
