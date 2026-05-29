import pandas as pd

def check_excel(filepath):
    try:
        df = pd.read_excel(filepath, sheet_name='Summary')
        print(f"--- Summary of {filepath} ---")
        print(df)
    except Exception as e:
        print(f"Could not read {filepath}: {e}")

check_excel('equityV-adj1.xlsx')
print("\n" + "="*50 + "\n")

# For the original one, it might have a different sheet name or structure
try:
    df_orig = pd.read_excel('trendstrategy_results_equityV.xlsx')
    print("--- First few rows of trendstrategy_results_equityV.xlsx ---")
    print(df_orig.head())
except Exception as e:
    print(f"Could not read trendstrategy_results_equityV.xlsx: {e}")
