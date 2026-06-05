
import pandas as pd
import numpy as np
import os
from backtest_v2 import clean_data as clean_legacy
from backtest_v2 import BacktesterV2
from backtest_adj4 import BacktesterVol

def run_diagnostics():
    """
    Consolidated diagnostic tool to verify B303 engine calibration and data source differences.
    """
    sma, roc, sl, reb = 303, 14, 0.0999, 9

    files_to_check = ['樣本集-1.xlsx', '資料26Q2-1.xlsx']

    for data_file in files_to_check:
        if not os.path.exists(data_file):
            print(f"File {data_file} not found, skipping...")
            continue

        print(f"\n=== Testing File: {data_file} ===")
        prices, volumes, code_to_name = clean_legacy(data_file)

        # 1. Legacy Engine
        print("Running Legacy Engine (BacktesterV2)...")
        bt_v2 = BacktesterV2(prices, volumes, code_to_name, initial_capital=30000000)
        eq_v2, _, _, _, _ = bt_v2.run(sma, roc, sl, reb, 'peak', 10)

        # 2. New Engine Baseline
        print("Running New Engine (BacktesterVol) - Baseline Configuration...")
        bt_vol = BacktesterVol(prices, volumes, code_to_name, trading_capital=30000000, authorized_capital=150000000)
        eq_vol_base, _, _, _ = bt_vol.run(
            sma_period=sma, roc_period=roc, stop_loss_type='fixed', stop_loss_val=sl,
            rebalance_interval=reb, use_market_filter=False, start_date='2019-01-01', end_date='2025-12-31'
        )

        # 3. New Engine Scenario C
        print("Running New Engine (BacktesterVol) - Scenario C Configuration...")
        eq_vol_c, _, _, _ = bt_vol.run(
            sma_period=sma, roc_period=roc, stop_loss_type='vol', vol_multiplier=2.7,
            use_breadth_weight=True, rebalance_interval=reb, use_market_filter=True,
            breadth_threshold=0.42, breadth_window=290, start_date='2019-01-01', end_date='2025-12-31'
        )

        def print_metrics(label, eq):
            equity = eq['權益']
            ret = (equity.iloc[-1] / equity.iloc[0]) - 1
            years = (eq['日期'].iloc[-1] - eq['日期'].iloc[0]).days / 365.25
            cagr = (1 + ret) ** (1 / years) - 1
            mdd = ((equity - equity.cummax()) / equity.cummax()).min()
            print(f"  {label:30s} | CAGR: {cagr:6.2%} | MaxDD: {mdd:7.2%}")

        print_metrics("Legacy Engine (Original)", eq_v2[eq_v2['日期'] >= '2019-01-01'])
        print_metrics("New Engine (Baseline)", eq_vol_base)
        print_metrics("New Engine (Scenario C)", eq_vol_c)

if __name__ == "__main__":
    run_diagnostics()
