
import pandas as pd
import numpy as np

def compare_excel_files(file1, file2):
    print(f"Comparing {file1} and {file2}...")

    def load_data(filepath):
        df_prices = pd.read_excel(filepath, sheet_name='還原收盤價', header=None)
        stock_codes = df_prices.iloc[0, 1:].values
        stock_names = df_prices.iloc[1, 1:].values
        date_strings = df_prices.iloc[2:, 0].astype(str).str[:8]
        dates = pd.to_datetime(date_strings, format='%Y%m%d')
        prices = df_prices.iloc[2:, 1:].astype(float)
        prices.index = dates
        prices.columns = stock_codes
        return prices, stock_names

    prices1, names1 = load_data(file1)
    prices2, names2 = load_data(file2)

    # Filter for 2019-2025
    mask1 = (prices1.index >= '2019-01-01') & (prices1.index <= '2025-12-31')
    mask2 = (prices2.index >= '2019-01-01') & (prices2.index <= '2025-12-31')
    p1 = prices1.loc[mask1]
    p2 = prices2.loc[mask2]

    print(f"File 1 shape: {p1.shape}")
    print(f"File 2 shape: {p2.shape}")

    # Check common dates
    common_dates = p1.index.intersection(p2.index)
    print(f"Common dates: {len(common_dates)}")

    # Check common stocks
    common_stocks = p1.columns.intersection(p2.columns)
    print(f"Common stocks: {len(common_stocks)}")

    p1_common = p1.loc[common_dates, common_stocks]
    p2_common = p2.loc[common_dates, common_stocks]

    # Calculate difference
    diff = p1_common - p2_common
    rel_diff = (p1_common - p2_common) / p2_common

    # Summary of differences
    has_diff = np.abs(rel_diff) > 1e-6
    diff_count = has_diff.sum().sum()
    total_count = p1_common.size

    print(f"Total cells compared: {total_count}")
    print(f"Cells with differences (>0.0001%): {diff_count} ({diff_count/total_count:.2%})")

    if diff_count > 0:
        # Top stocks with differences
        stock_diff_counts = has_diff.sum(axis=0).sort_values(ascending=False)
        print("\nTop 10 stocks with most price differences:")
        print(stock_diff_counts.head(10))

        # Max relative difference
        max_rel_diff = np.abs(rel_diff).max().max()
        print(f"\nMaximum relative difference: {max_rel_diff:.4%}")

        # Check if the difference is a constant multiplier per stock
        print("\nChecking if differences are constant multipliers (different adjustment base):")
        for stock in stock_diff_counts.head(5).index:
            ratios = p1_common[stock] / p2_common[stock]
            # Filter out NaNs and Zeros
            ratios = ratios[np.isfinite(ratios) & (ratios != 0)]
            std_ratio = ratios.std()
            print(f"Stock {stock}: Ratio Mean={ratios.mean():.4f}, Std={std_ratio:.6f}")
            if std_ratio < 1e-4:
                print(f"  -> Likely different adjustment base (constant multiplier).")
            else:
                print(f"  -> Dynamic differences (different adjustment logic or data points).")

    # Check for missing stocks
    missing_in_2 = p1.columns.difference(p2.columns)
    missing_in_1 = p2.columns.difference(p1.columns)
    if not missing_in_2.empty:
        print(f"\nStocks in {file1} but missing in {file2}: {len(missing_in_2)}")
        print(missing_in_2[:10].tolist())
    if not missing_in_1.empty:
        print(f"\nStocks in {file2} but missing in {file1}: {len(missing_in_1)}")
        print(missing_in_1[:10].tolist())

if __name__ == "__main__":
    compare_excel_files('資料26Q2-1.xlsx', '樣本集-1.xlsx')
