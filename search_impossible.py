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
    return {**p, 'Full_CAGR': cagr, 'Full_Calmar': calmar, 'Full_MaxDD': mdd, 'Ret_2022': y_rets.get(2022, 0), 'Ret_2023': y_rets.get(2023, 0)}

def main():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterATR(prices, volumes, code_to_name)

    # 專門尋找 2022/2023 的高點
    rocs = [5, 7, 10, 14, 20]
    rebs = [3, 5, 9]
    atr_ps = [10, 20]
    atr_ms = [2.0, 3.0, 4.0, 5.0]
    mkt_ts = [0.20, 0.30, 0.40]
    mkt_ss = [5, 14, 20]

    results = []
    param_list = list(itertools.product(rocs, rebs, atr_ps, atr_ms, mkt_ts, mkt_ss))
    print(f"Total: {len(param_list)}")

    for i, (roc, reb, ap, am, mt, ms) in enumerate(param_list):
        if i % 100 == 0: print(f"Progress: {i}/{len(param_list)}")
        p = {'roc': roc, 'reb': reb, 'atr_p': ap, 'atr_m': am, 'mkt_t': mt, 'mkt_s': ms}
        res = run_sim(bt, p)
        if res: results.append(res)

    df = pd.DataFrame(results)
    print("\nTop 2022 Returns:")
    print(df.sort_values('Ret_2022', ascending=False).head(10).to_string())
    print("\nTop 2023 Returns:")
    print(df.sort_values('Ret_2023', ascending=False).head(10).to_string())

if __name__ == "__main__":
    main()
