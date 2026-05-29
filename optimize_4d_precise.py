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

    # 精細化 4D 搜索
    sma_range = [250, 303, 350]
    roc_range = [8, 10, 12, 14]
    mult_range = [2.3, 2.5, 2.7, 3.0]
    breadth_range = [240, 270, 290, 310]

    results = []

    total = len(sma_range) * len(roc_range) * len(mult_range) * len(breadth_range)
    count = 0

    print(f"開始精確 4D 搜索 (全期間執行以正確獲取 2022 年度表現)，總計 {total} 組...")

    for sma in sma_range:
        for roc in roc_range:
            for mult in mult_range:
                for b_win in breadth_range:
                    count += 1
                    if count % 20 == 0:
                        print(f"進度: {count}/{total}")

                    # 執行全期間
                    eq, _, _, _ = bt.run(
                        sma_period=sma, roc_period=roc,
                        vol_multiplier=mult, breadth_window=b_win
                    )
                    m = calculate_metrics_dual(eq, TRADING_CAP, AUTH_CAP)

                    # 獲取 2022 年度表現 (carry-over 模式)
                    yearly_perf = m['Yearly Performance']
                    ret_2022 = yearly_perf.loc[2022, '年度報酬率'] if 2022 in yearly_perf.index else -1

                    results.append({
                        'SMA': sma, 'ROC': roc, 'Multiplier': mult, 'Breadth': b_win,
                        'Full_Trading_CAGR': m['Trading CAGR'],
                        'Full_Std_MaxDD': m['Standard MaxDD'],
                        'Full_Calmar': m['Trading Calmar'],
                        'Return_2022': ret_2022
                    })

    df = pd.DataFrame(results)
    df.to_csv('optimization_results_4d_precise.csv', index=False)
    print("搜索完成，結果儲存至 optimization_results_4d_precise.csv")

if __name__ == "__main__":
    main()
