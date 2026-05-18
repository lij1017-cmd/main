import pandas as pd
import numpy as np
from backtest_atr import clean_data, BacktesterATR, calculate_metrics
import itertools
from multiprocessing import Pool

def run_sim(params):
    prices, volumes, code_to_name, p = params
    bt = BacktesterATR(prices, volumes, code_to_name)
    equity, _, _, _, _ = bt.run(
        sma_period=p['sma'],
        roc_period=p['roc'],
        stop_loss_type='atr',
        atr_period=p['atr_p'],
        atr_multiplier=p['atr_m'],
        breadth_threshold=p['mkt_t'],
        mkt_sma_window=p['mkt_s'],
        breadth_window=p['mkt_w']
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

    # 擴展搜尋範圍
    # 保持 SMA 為 303 作為基本前提，但稍微變動 ROC
    rocs = [12, 14, 16]
    atr_ps = [10, 20]
    atr_ms = [2.0, 3.0, 4.0, 5.0]
    mkt_ts = [0.4, 0.45, 0.5] # 提高門檻看能否改善 2022
    mkt_ss = [10, 14, 20]

    param_list = []
    for roc, ap, am, mt, ms in itertools.product(rocs, atr_ps, atr_ms, mkt_ts, mkt_ss):
        param_list.append({
            'sma': 303, 'roc': roc, 'atr_p': ap, 'atr_m': am, 'mkt_t': mt, 'mkt_s': ms, 'mkt_w': 290
        })

    # 使用多進程加速
    print(f"Total combinations: {len(param_list)}")
    # 由於環境限制，我們手動執行一小部分或使用簡單循環，或者在 run_sim 中傳入價格
    results = []
    for i, p in enumerate(param_list):
        if i % 20 == 0: print(f"Progress: {i}/{len(param_list)}")
        res = run_sim((prices, volumes, code_to_name, p))
        results.append(res)

    flat_results = []
    for r in results:
        row = {
            'ROC': r['roc'], 'ATR_P': r['atr_p'], 'ATR_M': r['atr_m'],
            'MKT_T': r['mkt_t'], 'MKT_S': r['mkt_s'],
            'Full_CAGR': r['full_cagr'], 'Full_MaxDD': r['full_mdd'], 'Full_Calmar': r['full_calmar']
        }
        for year, (cagr, calmar) in r['years'].items():
            row[f'{year}_CAGR'] = cagr
            row[f'{year}_Calmar'] = calmar
        flat_results.append(row)

    df = pd.DataFrame(flat_results)
    df.to_csv('optimization_results_all.csv', index=False)
    print("Saved to optimization_results_all.csv")

if __name__ == "__main__":
    optimize()
