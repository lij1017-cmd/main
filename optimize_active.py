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
        mkt_sma_window=p['mkt_s'],
        breadth_window=290
    )
    if equity.empty: return None
    cagr, mdd, calmar, _ = calculate_metrics(equity)
    y_rets = {}
    equity['Year'] = equity['日期'].dt.year
    for year, group in equity.groupby('Year'):
        _, _, _, yr = calculate_metrics(group)
        y_rets[year] = yr
    return {**p, 'Full_CAGR': cagr, 'Full_Calmar': calmar, 'Full_MaxDD': mdd, 'years': y_rets}

def main():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterATR(prices, volumes, code_to_name)

    # 專攻 2023/2025
    rocs = [8, 10, 12]
    rebs = [3, 5, 7]
    atr_ps = [10, 20]
    atr_ms = [2.0, 2.5, 3.0]
    mkt_ts = [0.35, 0.40, 0.42]
    mkt_ss = [10, 14, 20]

    param_list = list(itertools.product(rocs, rebs, atr_ps, atr_ms, mkt_ts, mkt_ss))
    print(f"Total: {len(param_list)}")

    results = []
    for i, (roc, reb, ap, am, mt, ms) in enumerate(param_list):
        if i % 100 == 0: print(f"Progress: {i}/{len(param_list)}")
        p = {'roc': roc, 'reb': reb, 'atr_p': ap, 'atr_m': am, 'mkt_t': mt, 'mkt_s': ms}
        res = run_sim(bt, p)
        if res:
            row = {**p, 'Full_CAGR': res['Full_CAGR'], 'Full_Calmar': res['Full_Calmar'], 'Full_MaxDD': res['Full_MaxDD']}
            for y, r in res['years'].items():
                row[f'Ret_{y}'] = r
            results.append(row)

    df = pd.DataFrame(results)
    df.to_csv('opt_active.csv', index=False)

    # 篩選
    cond = (df['Ret_2022'] > 0) & (df['Ret_2023'] > 0.20) & (df['Ret_2025'] > 0.20)
    print("Candidates (2022>0, 2023>20%, 2025>20%):")
    print(df[cond].sort_values('Full_Calmar', ascending=False).head(20).to_string())

if __name__ == "__main__":
    main()
