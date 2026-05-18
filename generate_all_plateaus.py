import pandas as pd
import numpy as np
from backtest_atr import clean_data, BacktesterATR, calculate_metrics

def generate_plateaus():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterATR(prices, volumes, code_to_name)

    periods = [10, 15, 20, 25]
    multipliers = [3.5, 4.0, 4.3, 4.5, 5.0]

    for roc_val in [10, 14]:
        plateau = []
        for ap in periods:
            for am in multipliers:
                e, _, _, _, _ = bt.run(
                    sma_period=303, roc_period=roc_val, stop_loss_type='atr',
                    atr_period=ap, atr_multiplier=am,
                    rebalance_interval=9, use_market_filter=True, breadth_threshold=0.42, mkt_sma_window=14
                )
                c, m, ca, r = calculate_metrics(e)
                plateau.append({'ATR_P': ap, 'ATR_M': am, 'CAGR': c, 'MaxDD': m, 'Calmar': ca})

        df = pd.DataFrame(plateau)
        print(f"\n--- ROC {roc_val} ATR 參數高原表 (Calmar Ratio) ---")
        print(df.pivot(index='ATR_P', columns='ATR_M', values='Calmar'))
        print(f"\n--- ROC {roc_val} ATR 參數高原表 (CAGR) ---")
        print(df.pivot(index='ATR_P', columns='ATR_M', values='CAGR'))

if __name__ == "__main__":
    generate_plateaus()
