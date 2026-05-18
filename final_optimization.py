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
    if equity.empty:
        return None

    cagr, mdd, calmar, _ = calculate_metrics(equity)
    year_metrics = {}
    equity['Year'] = equity['日期'].dt.year
    for year, group in equity.groupby('Year'):
        y_cagr, y_mdd, y_calmar, y_ret = calculate_metrics(group)
        year_metrics[year] = {'cagr': y_cagr, 'calmar': y_calmar, 'ret': y_ret}

    return {
        **p,
        'Full_CAGR': cagr,
        'Full_MaxDD': mdd,
        'Full_Calmar': calmar,
        'years': year_metrics
    }

def main():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterATR(prices, volumes, code_to_name)

    # 針對性優化
    rocs = [10, 11, 12, 13, 14]
    atr_ps = [10, 15, 20]
    atr_ms = [2.5, 3.0, 3.5, 4.0]
    rebs = [5, 9, 13]
    mkt_ts = [0.4, 0.45, 0.5]
    mkt_ss = [10, 14, 20]

    # 由於組合較多，我們先做一個過濾
    param_list = list(itertools.product(rocs, atr_ps, atr_ms, rebs, mkt_ts, mkt_ss))
    print(f"Testing {len(param_list)} combinations...")

    results = []
    for i, (roc, ap, am, reb, mt, ms) in enumerate(param_list):
        if i % 100 == 0: print(f"Progress: {i}/{len(param_list)}")
        p = {'roc': roc, 'atr_p': ap, 'atr_m': am, 'reb': reb, 'mkt_t': mt, 'mkt_s': ms}
        res = run_sim(bt, p)
        if res:
            results.append(res)

    flat_results = []
    for r in results:
        row = {k: v for k, v in r.items() if k != 'years'}
        for y, metrics in r['years'].items():
            row[f'{y}_CAGR'] = metrics['cagr']
            row[f'{y}_Calmar'] = metrics['calmar']
            row[f'{y}_Ret'] = metrics['ret']
        flat_results.append(row)

    df = pd.DataFrame(flat_results)
    df.to_csv('final_optimization_results.csv', index=False)

    # 找出接近目標的
    # 目標: Full CAGR > 35%, Full Calmar > 2.9, 2022 Ret > 0, 其他年 CAGR > 25%
    # 放寬 2022 和 2023 的限制來尋找最佳解
    cond = (df['Full_CAGR'] > 0.35) & (df['2022_Ret'] > 0)
    top = df[cond].sort_values('Full_Calmar', ascending=False)

    print("\nTop Candidates (Full CAGR > 35% & 2022 Positive):")
    cols = ['roc', 'atr_p', 'atr_m', 'reb', 'mkt_t', 'mkt_s', 'Full_CAGR', 'Full_Calmar', '2022_Ret', '2023_CAGR', '2025_CAGR']
    print(top[cols].head(20).to_string())

if __name__ == "__main__":
    main()
