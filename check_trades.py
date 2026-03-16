import pandas as pd
import numpy as np

# Load the generated results to inspect
try:
    df = pd.read_excel('trendstrategy_results_equity2025新.xlsx', sheet_name='Trades')
    print("First 10 Buy Trades:")
    buys = df[df['狀態'] == '買進'].head(10)
    for idx, row in buys.iterrows():
        amt = row['價格'] * row['股數']
        print(f"Date: {row['日期']}, Stock: {row['股票代號']}, Price: {row['價格']}, Shares: {row['股數']}, Amount: {amt:,.0f}, Fee: {row['買入手續費']:,.0f}")
except Exception as e:
    print(f"Error: {e}")
