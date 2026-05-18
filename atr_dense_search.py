import pandas as pd
import numpy as np
from backtest_atr import clean_data, BacktesterATR, calculate_metrics
import itertools

def run_backtest(bt, p):
    equity, _, _, _, _ = bt.run(
        sma_period=303,
        roc_period=p['roc'],
        stop_loss_type='atr',
        atr_period=p['atr_p'],
        atr_multiplier=p['atr_m'],
        rebalance_interval=9,
        use_market_filter=True,
        breadth_threshold=p['mkt_t'],
        mkt_sma_window=p['mkt_s'],
        breadth_window=290
    )
    if equity.empty:
        return 0, 0, 0, {}

    cagr, mdd, calmar, _ = calculate_metrics(equity)
    year_metrics = {}
    equity['Year'] = equity['日期'].dt.year
    for year, group in equity.groupby('Year'):
        y_cagr, y_mdd, y_calmar, y_ret = calculate_metrics(group)
        year_metrics[year] = {'cagr': y_cagr, 'mdd': y_mdd, 'calmar': y_calmar}

    return cagr, mdd, calmar, year_metrics

def main():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterATR(prices, volumes, code_to_name)

    # 密集搜尋
    rocs = range(10, 21)
    atr_ps = [10, 15, 20]
    atr_ms = np.arange(1.5, 4.1, 0.5)

    # 保持市場濾網參數不變 (由 reproduce_equityV_filter.md 決定)
    mkt_t = 0.42
    mkt_s = 14

    results = []
    for roc, ap, am in itertools.product(rocs, atr_ps, atr_ms):
        p = {'roc': roc, 'atr_p': ap, 'atr_m': am, 'mkt_t': mkt_t, 'mkt_s': mkt_s}
        cagr, mdd, calmar, years = run_backtest(bt, p)

        res = {
            'ROC': roc, 'ATR_P': ap, 'ATR_M': am,
            'Full_CAGR': cagr, 'Full_Calmar': calmar
        }
        for y in range(2020, 2026):
            if y in years:
                res[f'{y}_CAGR'] = years[y]['cagr']
                res[f'{y}_Calmar'] = years[y]['calmar']
            else:
                res[f'{y}_CAGR'] = 0
                res[f'{y}_Calmar'] = 0
        results.append(res)

    df = pd.DataFrame(results)
    df.to_csv('atr_dense_search.csv', index=False)

    # 篩選符合條件的 (放寬一點 CAGR 限制來找方向)
    candidates = df[(df['Full_CAGR'] > 0.33) & (df['2022_CAGR'] > 0)]
    print("Top candidates by Full_Calmar:")
    print(candidates.sort_values('Full_Calmar', ascending=False).head(10).to_string())

if __name__ == "__main__":
    main()
