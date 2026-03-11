import pandas as pd
import os

files = ['trendstrategy_results_equity25(成).xlsx', 'reproduce_report25(成).md']
for f in files:
    if os.path.exists(f):
        print(f"{f} exists.")
    else:
        print(f"{f} MISSING.")

# Check Excel sheets
xls = pd.ExcelFile('trendstrategy_results_equity25(成).xlsx')
print(f"Sheets: {xls.sheet_names}")

summary = pd.read_excel('trendstrategy_results_equity25(成).xlsx', sheet_name='Summary')
print("Summary metrics:")
print(summary)

trades = pd.read_excel('trendstrategy_results_equity25(成).xlsx', sheet_name='Trades')
print(f"Total trades logged: {len(trades)}")
print(trades.head())

# Check for transaction cost logic impact in Trades (if any, though not explicit in columns)
# Mostly we check if prices and shares look reasonable.
