import pandas as pd
import numpy as np
import pickle
from backtester import Backtester, calculate_metrics

def verify():
    prices = pd.read_pickle('prices_cleaned.pkl')
    with open('best_params.pkl', 'rb') as f:
        best_params = pickle.load(f)

    sma_p, roc_p, sl_p = best_params
    bt = Backtester(prices)
    eq, trades, holdings, rebalance_log = bt.run(sma_p, roc_p, sl_p)

    # 1. Verification for 2019/05/02
    target_date = pd.Timestamp('2019-05-02')

    # Calculate indicators manually for that date
    sma = prices.rolling(window=sma_p).mean()
    roc = prices.pct_change(periods=roc_p)

    if target_date in prices.index:
        current_prices = prices.loc[target_date]
        current_sma = sma.loc[target_date]
        current_roc = roc.loc[target_date]

        # Sort all assets by ROC
        roc_sorted = current_roc.sort_values(ascending=False)
        top_5_codes = roc_sorted.head(5).index

        print("--- 2019/05/02 Top 5 Assets by ROC ---")
        verification_table = []
        for code in top_5_codes:
            p = current_prices[code]
            s = current_sma[code]
            r = current_roc[code]
            meet_criteria = (p > s) and (r > 0)
            verification_table.append({
                '股票代號': code,
                'ROC': f"{r:.2%}",
                'Price': p,
                'SMA': f"{s:.2f}",
                'Price > SMA': meet_criteria
            })
        print(pd.DataFrame(verification_table))

        # Check holdings on that date
        holdings_on_date = holdings[holdings['Date'] == target_date]
        if not holdings_on_date.empty:
            print(f"\nHoldings on {target_date}: {holdings_on_date.iloc[0]['Holdings']}")
        else:
            print(f"\nNo holdings record found for {target_date}")

    # 2. Stop Loss Examples
    sl_trades = trades[trades['Reason'] == 'Stop Loss']
    print("\n--- Stop Loss Trade Examples ---")
    if not sl_trades.empty:
        print(sl_trades[['Asset', 'Sell_Date', 'Sell_Price', 'Shares']].head(2))
    else:
        print("No Stop Loss trades found in the log.")

if __name__ == "__main__":
    verify()
