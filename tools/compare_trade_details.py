
import pandas as pd
import numpy as np

def compare_trades(file_legacy, file_new):
    print(f"Comparing round-trip trades between:\n  Legacy: {file_legacy}\n  New:    {file_new}\n")

    def load_trades(filepath):
        # Load Trades2 sheet which usually contains the standardized trade records
        df = pd.read_excel(filepath, sheet_name='Trades2')
        # Standardize columns
        # Expected: 買進訊號日期, 股票代號, 股票名稱, 賣出訊號日期
        df['買進訊號日期'] = pd.to_datetime(df['買進訊號日期']).dt.strftime('%Y-%m-%d')
        df['賣出訊號日期'] = pd.to_datetime(df['賣出訊號日期']).dt.strftime('%Y-%m-%d')
        df['股票代號'] = df['股票代號'].astype(str)
        return df

    try:
        df_l = load_trades(file_legacy)
        df_n = load_trades(file_new)
    except Exception as e:
        print(f"Error loading files: {e}")
        return

    # Create a unique key for each trade: Date_In + Symbol + Date_Out
    df_l['trade_key'] = df_l['買進訊號日期'] + "_" + df_l['股票代號'] + "_" + df_l['賣出訊號日期']
    df_n['trade_key'] = df_n['買進訊號日期'] + "_" + df_n['股票代號'] + "_" + df_n['賣出訊號日期']

    keys_l = set(df_l['trade_key'])
    keys_n = set(df_n['trade_key'])

    common = keys_l.intersection(keys_n)
    only_l = keys_l - keys_n
    only_n = keys_n - keys_l

    total_l = len(keys_l)
    total_n = len(keys_n)

    print(f"Total trades in Legacy: {total_l}")
    print(f"Total trades in New:    {total_n}")
    print(f"Common trades:          {len(common)} ({len(common)/total_l:.2%} of Legacy, {len(common)/total_n:.2%} of New)")

    if only_l:
        print(f"\nTrades ONLY in Legacy ({len(only_l)}):")
        print(df_l[df_l['trade_key'].isin(list(only_l)[:10])][['買進訊號日期', '股票代號', '賣出訊號日期', '報酬率']])

    if only_n:
        print(f"\nTrades ONLY in New ({len(only_n)}):")
        print(df_n[df_n['trade_key'].isin(list(only_n)[:10])][['買進訊號日期', '股票代號', '賣出訊號日期', '報酬率']])

    # Compare execution prices for common trades
    print("\nPrice Comparison for Common Trades:")
    merged = pd.merge(df_l, df_n, on='trade_key', suffixes=('_L', '_N'))
    merged['price_diff_in'] = (merged['T+1日買進價格_L'] - merged['T+1日買進價格_N']).abs()
    merged['price_diff_out'] = (merged['T+1日賣出價格_L'] - merged['T+1日賣出價格_N']).abs()

    p_diff_in = (merged['price_diff_in'] > 0.01).sum()
    p_diff_out = (merged['price_diff_out'] > 0.01).sum()

    print(f"Common trades with Entry Price difference: {p_diff_in}")
    print(f"Common trades with Exit Price difference:  {p_diff_out}")

    if p_diff_in > 0:
        print("\nExample Entry Price difference:")
        print(merged[merged['price_diff_in'] > 0.01][['trade_key', 'T+1日買進價格_L', 'T+1日買進價格_N']].head(5))

if __name__ == "__main__":
    compare_trades('trendstrategy_results_equityV.xlsx', 'trendstrategy_results_equityV-adj4-1-251231.xlsx')
