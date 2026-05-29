import pandas as pd
import numpy as np
from backtest_vol import BacktesterVol, calculate_metrics_dual, clean_data
import os

def main():
    filepath = '樣本集-1.xlsx'
    prices, volumes, code_to_name = clean_data(filepath)
    TRADING_CAP = 30000000
    AUTH_CAP = 150000000
    bt = BacktesterVol(prices, volumes, code_to_name, trading_capital=TRADING_CAP, authorized_capital=AUTH_CAP)

    # 測試參數範圍
    sma_range = [250, 303, 350]
    roc_range = [10, 14, 20]
    vol_mult_range = [2.5, 2.7, 3.0, 3.5]
    breadth_range = [240, 290]

    results = []

    print("進行快速敏感度分析，聚焦於 2022 年 (空頭市場) 的表現...")

    for sma in sma_range:
        for roc in roc_range:
            for mult in vol_mult_range:
                for b_win in breadth_range:
                    # 測試 2022 年度
                    eq, _, _, _ = bt.run(
                        sma_period=sma,
                        roc_period=roc,
                        vol_multiplier=mult,
                        breadth_window=b_win,
                        start_date='2022-01-01',
                        end_date='2022-12-31'
                    )
                    metrics = calculate_metrics_dual(eq, TRADING_CAP, AUTH_CAP)
                    yearly_perf = metrics['Yearly Performance']
                    ret_2022 = yearly_perf.iloc[0]['年度報酬率'] if not yearly_perf.empty else -1

                    results.append({
                        'SMA': sma, 'ROC': roc, 'Multiplier': mult, 'Breadth': b_win,
                        'Return_2022': ret_2022,
                        'Std_MaxDD': metrics['Standard MaxDD']
                    })

    df = pd.DataFrame(results)
    df.to_csv('diagnostic_2022.csv', index=False)
    print("分析完成，結果儲存至 diagnostic_2022.csv")
    print(df.sort_values('Return_2022', ascending=False).head(10))

if __name__ == "__main__":
    main()
