import pandas as pd
import numpy as np
from backtest_vol import BacktesterVol, calculate_metrics_dual, clean_data
import os

def main():
    filepath = '樣本集-1.xlsx'
    prices, volumes, code_to_name = clean_data(filepath)
    TRADING_CAP = 30000000
    AUTH_CAP = 150000000

    # 測試 ROC 14, Multiplier 2.7, Breadth 290 (搜索出的最優組合)
    bt = BacktesterVol(prices, volumes, code_to_name, trading_capital=TRADING_CAP, authorized_capital=AUTH_CAP)

    print("測試優化參數 (SMA 303, ROC 14, Mult 2.7, Breadth 290)...")
    eq, _, _, _ = bt.run(303, 14, vol_multiplier=2.7, breadth_window=290)
    m = calculate_metrics_dual(eq, TRADING_CAP, AUTH_CAP)
    print(f"全期間 Calmar: {m['Trading Calmar']:.2f}, MDD: {m['Standard MaxDD']:.2%}")
    print(f"2022 報酬: {m['Yearly Performance'].loc[2022, '年度報酬率']:.2%}")

if __name__ == "__main__":
    main()
