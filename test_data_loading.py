from backtest_equity2MA import clean_data
import pandas as pd

try:
    prices, volumes, names = clean_data('樣本集-1.xlsx')
    print(f"Data loaded successfully.")
    print(f"Prices shape: {prices.shape}")
    print(f"Volumes shape: {volumes.shape}")
    print(f"Number of stocks: {len(names)}")
    print(f"First 5 dates: {prices.index[:5].tolist()}")
    print(f"Last 5 dates: {prices.index[-5:].tolist()}")
except Exception as e:
    print(f"Error: {e}")
