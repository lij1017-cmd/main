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
        rebalance_interval=p['reb'],
        use_market_filter=True,
        breadth_threshold=p['mkt_t'],
        mkt_sma_window=14,
        breadth_window=290
    )
    if equity.empty: return None
    cagr, mdd, calmar, _ = calculate_metrics(equity)
    y_metrics = {}
    equity['Year'] = equity['日期'].dt.year
    for year, group in equity.groupby('Year'):
        yc, ym, yca, yr = calculate_metrics(group)
        y_metrics[year] = yr
    return {**p, 'Full_CAGR': cagr, 'Full_Calmar': calmar, 'years': y_metrics}

def main():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterATR(prices, volumes, code_to_name)

    rocs = [8, 10, 12, 14]
    rebs = [5, 7, 9, 11]
    atr_ps = [10, 15, 20]
    atr_ms = [3.5, 4.0, 4.5]
    mkt_ts = [0.42, 0.45, 0.48]

    param_list = list(itertools.product(rocs, rebs, atr_ps, atr_ms, mkt_ts))
    print(f"Total: {len(param_list)}")

    results = []
    for i, (roc, reb, ap, am, mt) in enumerate(param_list):
        if i % 100 == 0: print(f"Progress: {i}/{len(param_list)}")
        p = {'roc': roc, 'reb': reb, 'atr_p': ap, 'atr_m': am, 'mkt_t': mt}
        res = run_sim(bt, p)
        if res: results.append(res)

    flat = []
    for r in results:
        row = {k: v for k, v in r.items() if k != 'years'}
        for y, ret in r['years'].items():
            row[f'Ret_{y}'] = ret
        flat.append(row)

    df = pd.DataFrame(flat)
    df.to_csv('opt_v3.csv', index=False)

    cond = (df['Full_CAGR'] > 0.35) & (df['Ret_2022'] > 0)
    print("Top candidates by Full_Calmar:")
    print(df[cond].sort_values('Full_Calmar', ascending=False).head(20).to_string())

if __name__ == "__main__":
    main()
