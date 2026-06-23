import pandas as pd

def analyze_trades(filepath, name):
    try:
        # Check if it's the new format (sheet 'Trades') or old format
        try:
            df = pd.read_excel(filepath, sheet_name='Trades')
        except:
            df = pd.read_excel(filepath)

        print(f"--- Trade Analysis for {name} ---")
        print(f"Total Trades: {len(df)}")

        # Calculate approximate turnover if possible
        # Or just look at the distribution of reasons
        if '原因' in df.columns:
            print(df['原因'].value_counts())
        elif '狀態' in df.columns:
             print(df['狀態'].value_counts())

    except Exception as e:
        print(f"Error analyzing {name}: {e}")

analyze_trades('equityV-adj1.xlsx', 'Adj1')
analyze_trades('trendstrategy_results_equityV.xlsx', 'Original')
