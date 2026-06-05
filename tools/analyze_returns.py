
import pandas as pd
import numpy as np

def analyze_returns_diff(file1, file2):
    print(f"Analyzing Returns difference between {file1} and {file2}...")

    def load_data(filepath):
        df_prices = pd.read_excel(filepath, sheet_name='還原收盤價', header=None)
        stock_codes = df_prices.iloc[0, 1:].values
        date_strings = df_prices.iloc[2:, 0].astype(str).str[:8]
        dates = pd.to_datetime(date_strings, format='%Y%m%d')
        prices = df_prices.iloc[2:, 1:].astype(float)
        prices.index = dates
        prices.columns = stock_codes
        return prices

    p1 = load_data(file1)
    p2 = load_data(file2)

    common_dates = p1.index.intersection(p2.index)
    common_stocks = p1.columns.intersection(p2.columns)

    p1 = p1.loc[common_dates, common_stocks]
    p2 = p2.loc[common_dates, common_stocks]

    # Daily Returns
    ret1 = p1.pct_change()
    ret2 = p2.pct_change()

    ret_diff = (ret1 - ret2).abs()

    # Significant return differences
    significant_ret_diff = ret_diff > 1e-5
    diff_count = significant_ret_diff.sum().sum()

    print(f"Cells with significant daily return differences (>0.001%): {diff_count}")

    if diff_count > 0:
        # Stocks with most return differences
        stock_ret_diff = significant_ret_diff.sum(axis=0).sort_values(ascending=False)
        print("\nTop 5 stocks with return differences:")
        print(stock_ret_diff.head(5))

        # Example of a mismatch
        worst_stock = stock_ret_diff.index[0]
        mismatch_dates = significant_ret_diff[worst_stock]
        mismatch_dates = mismatch_dates[mismatch_dates].index

        print(f"\nExample mismatch for stock {worst_stock} on {mismatch_dates[0]}:")
        print(f"  File 1 Price: {p1.loc[mismatch_dates[0], worst_stock]:.4f}, Return: {ret1.loc[mismatch_dates[0], worst_stock]:.6%}")
        print(f"  File 2 Price: {p2.loc[mismatch_dates[0], worst_stock]:.4f}, Return: {ret2.loc[mismatch_dates[0], worst_stock]:.6%}")

    # Check SMA impact
    sma1 = p1.rolling(window=303).mean()
    sma2 = p2.rolling(window=303).mean()

    # Signal: Price > SMA
    sig1 = p1 > sma1
    sig2 = p2 > sma2

    signal_mismatch = (sig1 != sig2) & p1.notna() & p2.notna() & sma1.notna() & sma2.notna()
    mismatch_total = signal_mismatch.sum().sum()

    print(f"\nSignal Mismatches (Price > SMA303): {mismatch_total} cells")
    if mismatch_total > 0:
        stocks_with_mismatch = signal_mismatch.sum(axis=0).sort_values(ascending=False)
        print("Top 5 stocks with signal mismatches:")
        print(stocks_with_mismatch.head(5))

if __name__ == "__main__":
    analyze_returns_diff('資料26Q2-1.xlsx', '樣本集-1.xlsx')
