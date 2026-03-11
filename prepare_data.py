import pandas as pd
import numpy as np

def clean_data(filepath):
    # Load raw data
    # Row 0: Stock Codes, Row 1: Stock Names, Col 0: Index info (ignore), Col 1: Dates
    df_raw = pd.read_excel(filepath, header=None)

    # Extract headers
    stock_codes = df_raw.iloc[0, 2:].values
    stock_names = df_raw.iloc[1, 2:].values
    dates = pd.to_datetime(df_raw.iloc[2:, 1])

    # Extract prices
    prices = df_raw.iloc[2:, 2:].astype(float)
    prices.index = dates
    prices.columns = stock_codes

    code_to_name = dict(zip(stock_codes, stock_names))

    # Data Cleaning Rules from V6.md:
    # 較晚上市標的：以首次出現的價格往前期填補 (backfill)
    # 中途暫停交易標的：以前一日收盤價填補中間空白區域 (forward fill)
    prices = prices.ffill().bfill()

    # Drop columns that are entirely NaN (if any remain)
    prices = prices.dropna(axis=1, how='all')

    return prices, code_to_name

if __name__ == "__main__":
    prices, code_to_name = clean_data('個股1.xlsx')
    print(f"Number of assets: {len(prices.columns)}")
    print(f"Date range: {prices.index[0]} to {prices.index[-1]}")
    print(f"First 5 stock codes: {list(prices.columns[:5])}")
    print(f"First 5 stock names: {[code_to_name[c] for c in prices.columns[:5]]}")

    # Save cleaned data for later steps
    import pickle
    with open('cleaned_data.pkl', 'wb') as f:
        pickle.dump((prices, code_to_name), f)
    print("Cleaned data saved to cleaned_data.pkl")
