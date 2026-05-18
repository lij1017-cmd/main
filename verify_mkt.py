import pandas as pd
import numpy as np
from backtest_atr import clean_data, BacktesterATR, calculate_metrics
import itertools

def run_sim(bt, p):
    equity, _, _, _, _ = bt.run(
        sma_period=303,
        roc_period=p['roc'],
        stop_loss_type='atr',
        atr_period=p['atr_p'],
        atr_multiplier=p['atr_m'],
        rebalance_interval=9,
        use_market_filter=True,
        breadth_threshold=0.45, # 稍微提高
        mkt_sma_window=20,     # 稍微拉長
        breadth_window=290
    )
    if equity.empty: return None
    cagr, mdd, calmar, _ = calculate_metrics(equity)
    y_rets = {}
    equity['Year'] = equity['日期'].dt.year
    for year, group in equity.groupby('Year'):
        _, _, _, yr = calculate_metrics(group)
        y_rets[year] = yr
    return {'Full_CAGR': cagr, 'Full_Calmar': calmar, 'Full_MaxDD': mdd, 'years': y_rets}

def main():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterATR(prices, volumes, code_to_name)

    rocs = [12, 14, 16]
    aps = [10, 15, 20]
    ams = [3.5, 4.0, 4.5, 5.0]

    results = []
    for roc, ap, am in itertools.product(rocs, aps, ams):
        p = {'roc': roc, 'atr_p': ap, 'atr_m': am}
        res = run_sim(bt, p)
        if res:
            row = {**p, 'Full_CAGR': res['Full_CAGR'], 'Full_Calmar': res['Full_Calmar'], 'Full_MaxDD': res['Full_MaxDD']}
            for y, r in res['years'].items():
                row[f'Ret_{y}'] = r
            results.append(row)

    df = pd.DataFrame(results)
    print("Candidates (Mkt Filter 0.45/20):")
    cond = (df['Full_CAGR'] > 0.33) & (df['Ret_2022'] > 0)
    print(df[cond].sort_values('Full_Calmar', ascending=False).to_string())

if __name__ == "__main__":
    main()
