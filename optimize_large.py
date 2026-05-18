import pandas as pd
import numpy as np
from backtest_atr import clean_data, BacktesterATR, calculate_metrics
import itertools

def run_sim(prices, volumes, code_to_name, p):
    bt = BacktesterATR(prices, volumes, code_to_name)
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
    cagr, mdd, calmar, _ = calculate_metrics(equity)

    year_metrics = {}
    if not equity.empty:
        equity['Year'] = equity['日期'].dt.year
        for year, group in equity.groupby('Year'):
            y_cagr, y_mdd, y_calmar, y_ret = calculate_metrics(group)
            year_metrics[year] = (y_cagr, y_calmar)

    return {**p, 'full_cagr': cagr, 'full_mdd': mdd, 'full_calmar': calmar, 'years': year_metrics}

def optimize():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')

    # 網格搜尋範圍
    rocs = [12, 14, 16, 18]
    atr_ps = [10, 15, 20]
    atr_ms = [2.0, 2.5, 3.0, 3.5, 4.0]
    rebs = [5, 9, 14]
    mkt_ts = [0.42, 0.45]
    mkt_ss = [14, 20]

    results = []
    param_list = list(itertools.product(rocs, atr_ps, atr_ms, rebs, mkt_ts, mkt_ss))
    print(f"Total: {len(param_list)}")

    for i, (roc, ap, am, reb, mt, ms) in enumerate(param_list):
        if i % 50 == 0: print(f"Progress: {i}/{len(param_list)}")
        p = {'roc': roc, 'atr_p': ap, 'atr_m': am, 'reb': reb, 'mkt_t': mt, 'mkt_s': ms}
        res = run_sim(prices, volumes, code_to_name, p)
        results.append(res)

    flat_results = []
    for r in results:
        row = {
            'ROC': r['roc'], 'ATR_P': r['atr_p'], 'ATR_M': r['atr_m'], 'REB': r['reb'],
            'MKT_T': r['mkt_t'], 'MKT_S': r['mkt_s'],
            'Full_CAGR': r['full_cagr'], 'Full_MaxDD': r['full_mdd'], 'Full_Calmar': r['full_calmar']
        }
        for year, (cagr, calmar) in r['years'].items():
            row[f'{year}_CAGR'] = cagr
            row[f'{year}_Calmar'] = calmar
        flat_results.append(row)

    df = pd.DataFrame(flat_results)
    df.to_csv('optimization_results_large.csv', index=False)
    print("Saved to optimization_results_large.csv")

if __name__ == "__main__":
    optimize()
