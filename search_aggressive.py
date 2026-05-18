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

    # 嘗試更高 ROC 和更靈活的市場濾網
    rocs = [20, 25, 30, 35]
    rebs = [9, 14]
    atr_ps = [10, 20]
    atr_ms = [3.0, 4.0, 5.0, 6.0]
    mkt_ts = [0.35, 0.40, 0.42]
    mkt_ss = [5, 10, 14]

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
    df.to_csv('opt_aggressive.csv', index=False)

    # 篩選：2022正, 2020-2025 (除了2022) > 25% (或接近)
    cond = (df['Ret_2022'] > 0) & (df['Ret_2023'] > 0.20)
    print("Candidates (2022 > 0 and 2023 > 20%):")
    print(df[cond].sort_values('Full_Calmar', ascending=False).head(20).to_string())

if __name__ == "__main__":
    main()
