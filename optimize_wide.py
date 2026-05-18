import pandas as pd
import numpy as np
from backtest_atr import clean_data, BacktesterATR, calculate_metrics
import itertools

def run_sim(bt, p):
    equity, _, _, _, _ = bt.run(
        sma_period=303,
        roc_period=p['roc'],
        stop_loss_type='atr',
        atr_period=20,
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
    y_calmars = {}
    equity['Year'] = equity['日期'].dt.year
    for year, group in equity.groupby('Year'):
        yc, ym, yca, yr = calculate_metrics(group)
        y_rets[year] = yr
        y_calmars[year] = yca
    return {**p, 'Full_CAGR': cagr, 'Full_Calmar': calmar, 'Full_MaxDD': mdd, 'years': y_rets, 'calmars': y_calmars}

def main():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterATR(prices, volumes, code_to_name)

    # 追求 2023/2025 的回報
    rocs = [10, 14, 18]
    rebs = [7, 9, 11]
    ams = [4.0, 4.5, 5.0]
    mts = [0.30, 0.35, 0.40, 0.45]
    mss = [5, 10, 14, 20]

    param_list = list(itertools.product(rocs, rebs, ams, mts, mss))
    print(f"Total: {len(param_list)}")

    results = []
    for i, (roc, reb, am, mt, ms) in enumerate(param_list):
        if i % 100 == 0: print(f"Progress: {i}/{len(param_list)}")
        p = {'roc': roc, 'reb': reb, 'atr_m': am, 'mkt_t': mt, 'mkt_s': ms}
        res = run_sim(bt, p)
        if res:
            row = {**p, 'Full_CAGR': res['Full_CAGR'], 'Full_Calmar': res['Full_Calmar']}
            for y in [2020, 2021, 2022, 2023, 2024, 2025]:
                row[f'Ret_{y}'] = res['years'].get(y, 0)
                row[f'Cal_{y}'] = res['calmars'].get(y, 0)
            results.append(row)

    df = pd.DataFrame(results)
    df.to_csv('opt_wide.csv', index=False)

    # 嚴格篩選
    cond = (df['Ret_2022'] > 0) & (df['Ret_2023'] > 0.20) & (df['Ret_2025'] > 0.20)
    print("Candidates:")
    print(df[cond].sort_values('Full_Calmar', ascending=False).head(20).to_string())

if __name__ == "__main__":
    main()
