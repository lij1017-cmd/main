import pandas as pd
import numpy as np
from backtest_atr import clean_data, BacktesterATR, calculate_metrics
import itertools

def run_sim(bt, p):
    equity, _, _, _, _ = bt.run(
        sma_period=p['sma'],
        roc_period=p['roc'],
        stop_loss_type='atr',
        atr_period=p['atr_p'],
        atr_multiplier=p['atr_m'],
        rebalance_interval=p['reb'],
        use_market_filter=True,
        breadth_threshold=p['mkt_t'],
        mkt_sma_window=p['mkt_s'],
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

    # 探索 SMA 變化
    smas = [150, 200, 250, 300]
    rocs = [10, 15, 20]
    atr_ps = [10, 20]
    atr_ms = [3.0, 4.0, 5.0]
    rebs = [5, 9, 14]
    mkt_ts = [0.42, 0.45]
    mkt_ss = [14, 20]

    param_list = list(itertools.product(smas, rocs, atr_ps, atr_ms, rebs, mkt_ts, mkt_ss))
    print(f"Total: {len(param_list)}")

    results = []
    for i, (sma, roc, ap, am, reb, mt, ms) in enumerate(param_list):
        if i % 100 == 0: print(f"Progress: {i}/{len(param_list)}")
        p = {'sma': sma, 'roc': roc, 'atr_p': ap, 'atr_m': am, 'reb': reb, 'mkt_t': mt, 'mkt_s': ms}
        res = run_sim(bt, p)
        if res: results.append(res)

    flat = []
    for r in results:
        row = {k: v for k, v in r.items() if k != 'years'}
        for y, ret in r['years'].items():
            row[f'Ret_{y}'] = ret
        flat.append(row)

    df = pd.DataFrame(flat)
    df.to_csv('opt_sma.csv', index=False)

    # 篩選
    cond = (df['Full_CAGR'] > 0.35) & (df['Ret_2022'] > 0)
    print("Top candidates by Full_Calmar:")
    print(df[cond].sort_values('Full_Calmar', ascending=False).head(10).to_string())

if __name__ == "__main__":
    main()
