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
        mkt_sma_window=14,
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

    rocs = [8, 9, 10, 11, 12]
    rebs = [7, 9, 11]
    ams = [4.0, 4.5, 5.0, 5.5, 6.0]
    mts = [0.38, 0.40, 0.42]

    param_list = list(itertools.product(rocs, rebs, ams, mts))
    print(f"Total: {len(param_list)}")

    results = []
    for i, (roc, reb, am, mt) in enumerate(param_list):
        if i % 50 == 0: print(f"Progress: {i}/{len(param_list)}")
        p = {'roc': roc, 'reb': reb, 'atr_m': am, 'mkt_t': mt}
        res = run_sim(bt, p)
        if res:
            row = {**p, 'Full_CAGR': res['Full_CAGR'], 'Full_Calmar': res['Full_Calmar'], 'Full_MaxDD': res['Full_MaxDD']}
            for y, r in res['years'].items():
                row[f'Ret_{y}'] = r
            results.append(row)

    df = pd.DataFrame(results)
    df.to_csv('opt_focused.csv', index=False)

    # 篩選
    cond = (df['Ret_2022'] > 0)
    print("Candidates (Top CAGR):")
    print(df[cond].sort_values('Full_CAGR', ascending=False).head(20).to_string())

if __name__ == "__main__":
    main()
