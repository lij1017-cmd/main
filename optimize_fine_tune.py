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
        breadth_window=p['mkt_w']
    )
    if equity.empty: return None
    cagr, mdd, calmar, _ = calculate_metrics(equity)
    y_metrics = {}
    equity['Year'] = equity['日期'].dt.year
    for year, group in equity.groupby('Year'):
        yc, ym, yca, yr = calculate_metrics(group)
        y_metrics[year] = {'ret': yr, 'calmar': yca, 'mdd': ym}
    return {**p, 'Full_CAGR': cagr, 'Full_Calmar': calmar, 'Full_MaxDD': mdd, 'years': y_metrics}

def main():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterATR(prices, volumes, code_to_name)

    # 微調核心優選參數
    smas = [300, 303, 305]
    rocs = [10, 11]
    rebs = [9, 10]
    atr_ps = [20, 25]
    atr_ms = [4.5, 4.8, 5.0]
    mkt_ts = [0.42, 0.45]
    mkt_ss = [14, 20]
    mkt_ws = [280, 290, 300]

    param_list = list(itertools.product(smas, rocs, rebs, atr_ps, atr_ms, mkt_ts, mkt_ss, mkt_ws))
    print(f"Total: {len(param_list)}")

    results = []
    for i, (sma, roc, reb, ap, am, mt, ms, mw) in enumerate(param_list):
        if i % 100 == 0: print(f"Progress: {i}/{len(param_list)}")
        p = {'sma': sma, 'roc': roc, 'reb': reb, 'atr_p': ap, 'atr_m': am, 'mkt_t': mt, 'mkt_s': ms, 'mkt_w': mw}
        res = run_sim(bt, p)
        if res:
            row = {**p, 'Full_CAGR': res['Full_CAGR'], 'Full_Calmar': res['Full_Calmar'], 'Full_MaxDD': res['Full_MaxDD']}
            for y, m in res['years'].items():
                row[f'Ret_{y}'] = m['ret']
                row[f'Cal_{y}'] = m['calmar']
            results.append(row)

    df = pd.DataFrame(results)
    df.to_csv('opt_fine_tune.csv', index=False)

    # 篩選
    cond = (df['Full_CAGR'] > 0.35) & (df['Full_Calmar'] > 2.8) & (df['Ret_2022'] > 0)
    print("\nTop Candidates (CAGR > 35% & Calmar > 2.8 & 2022 > 0):")
    print(df[cond].sort_values('Full_Calmar', ascending=False).head(20).to_string())

if __name__ == "__main__":
    main()
