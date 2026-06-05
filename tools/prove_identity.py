
import pandas as pd
import numpy as np
from backtest_v2 import clean_data as clean_legacy
from backtest_adj4 import BacktesterVol

def prove_logic_identity():
    # 1. Load the LEGACY data (the one the user THINKS both used)
    data_file = '樣本集-1.xlsx'
    print(f"Testing logic on {data_file}...")
    prices, volumes, code_to_name = clean_legacy(data_file)

    # 2. Run NEW engine on LEGACY data
    sma, roc, sl, reb = 303, 14, 0.0999, 9
    bt_vol = BacktesterVol(prices, volumes, code_to_name, trading_capital=30000000, authorized_capital=150000000)
    _, _, t2_new, _ = bt_vol.run(
        sma_period=sma, roc_period=roc, stop_loss_type='fixed', stop_loss_val=sl,
        rebalance_interval=reb, use_market_filter=False, start_date='2019-01-01', end_date='2025-12-31'
    )

    # 3. Load the LEGACY results
    df_legacy = pd.read_excel('trendstrategy_results_equityV.xlsx', sheet_name='Trades2')

    # Standardize for comparison
    t2_new['key'] = pd.to_datetime(t2_new['買進訊號日期']).dt.strftime('%Y-%m-%d') + "_" + t2_new['股票代號'].astype(str)
    df_legacy['key'] = pd.to_datetime(df_legacy['買進訊號日期']).dt.strftime('%Y-%m-%d') + "_" + df_legacy['股票代號'].astype(str)

    keys_new = set(t2_new['key'])
    keys_legacy = set(df_legacy['key'])

    diff_new = keys_new - keys_legacy
    diff_legacy = keys_legacy - keys_new

    print(f"Total Legacy trades: {len(keys_legacy)}")
    print(f"Total New engine trades on same data: {len(keys_new)}")
    print(f"Common trades: {len(keys_new.intersection(keys_legacy))}")
    print(f"Discrepancies: {len(diff_new) + len(diff_legacy)}")

    if diff_new or diff_legacy:
        print("\nLogic Discrepancy Found!")
        if diff_new: print(f"Only in New: {list(diff_new)[:5]}")
        if diff_legacy: print(f"Only in Legacy: {list(diff_legacy)[:5]}")
    else:
        print("\n✅ SUCCESS: Logic is identical when data sources are aligned.")

if __name__ == "__main__":
    prove_logic_identity()
