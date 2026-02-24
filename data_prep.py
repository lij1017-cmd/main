import pandas as pd
import numpy as np
def clean_data(filepath):
    df_raw = pd.read_excel(filepath, header=None)
    stock_codes = df_raw.iloc[0, 2:].values
    stock_names = df_raw.iloc[1, 2:].values
    dates = pd.to_datetime(df_raw.iloc[2:, 1])
    prices = df_raw.iloc[2:, 2:].astype(float)
    prices.index = dates
    prices.columns = stock_codes
    code_to_name = dict(zip(stock_codes, stock_names))
    prices = prices.bfill().ffill().dropna(axis=1, how='all')
    return prices, code_to_name
if __name__ == "__main__":
    prices, code_to_name = clean_data('個股1.xlsx')
    prices.to_pickle('prices_cleaned.pkl')
    import pickle
    with open('code_to_name.pkl', 'wb') as f:
        pickle.dump(code_to_name, f)
