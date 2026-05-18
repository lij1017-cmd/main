import pandas as pd
import numpy as np
from backtest_atr import clean_data, BacktesterATR, calculate_metrics
import itertools

def optimize():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterATR(prices, volumes, code_to_name)

    # 網格搜尋參數
    atr_periods = [10, 15, 20, 25]
    atr_multipliers = np.arange(1.5, 6.5, 0.5)

    results = []

    for ap, am in itertools.product(atr_periods, atr_multipliers):
        print(f"Testing ATR Period: {ap}, Multiplier: {am}...")
        equity, _, _, _, _ = bt.run(
            sma_period=303,
            roc_period=14,
            stop_loss_type='atr',
            atr_period=ap,
            atr_multiplier=am,
            rebalance_interval=9,
            use_market_filter=True,
            breadth_threshold=0.42,
            mkt_sma_window=14,
            breadth_window=290
        )

        full_cagr, full_mdd, full_calmar, _ = calculate_metrics(equity)

        year_metrics = {}
        equity['Year'] = equity['日期'].dt.year
        for year, group in equity.groupby('Year'):
            y_cagr, y_mdd, y_calmar, y_ret = calculate_metrics(group)
            year_metrics[year] = (y_cagr, y_calmar)

        results.append({
            'atr_period': ap,
            'atr_multiplier': am,
            'full_cagr': full_cagr,
            'full_mdd': full_mdd,
            'full_calmar': full_calmar,
            'year_metrics': year_metrics
        })

    # 轉為 DataFrame 方便分析
    flat_results = []
    for r in results:
        row = {
            'Period': r['atr_period'],
            'Multiplier': r['atr_multiplier'],
            'Full_CAGR': r['full_cagr'],
            'Full_MaxDD': r['full_mdd'],
            'Full_Calmar': r['full_calmar']
        }
        for year, (cagr, calmar) in r['year_metrics'].items():
            row[f'{year}_CAGR'] = cagr
            row[f'{year}_Calmar'] = calmar
        flat_results.append(row)

    df = pd.DataFrame(flat_results)
    df.to_csv('optimization_results_atr.csv', index=False)
    print("Optimization finished. Results saved to optimization_results_atr.csv")

if __name__ == "__main__":
    optimize()
