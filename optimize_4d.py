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

    # 4D 搜索範圍
    sma_range = [200, 250, 303, 350]
    roc_range = [10, 14, 20, 25]
    mult_range = [1.5, 2.0, 2.5, 3.0, 3.5]
    breadth_range = [200, 240, 290, 330]

    results = []

    total = len(sma_range) * len(roc_range) * len(mult_range) * len(breadth_range)
    count = 0

    print(f"開始 4D 全域搜索，總計 {total} 組組合...")

    for sma in sma_range:
        for roc in roc_range:
            for mult in mult_range:
                for b_win in breadth_range:
                    count += 1
                    if count % 20 == 0:
                        print(f"進度: {count}/{total}")

                    # 測試全期間
                    eq_full, _, _, _ = bt.run(
                        sma_period=sma, roc_period=roc,
                        vol_multiplier=mult, breadth_window=b_win
                    )
                    m_full = calculate_metrics_dual(eq_full, TRADING_CAP, AUTH_CAP)

                    # 測試 2022 年度
                    eq_2022, _, _, _ = bt.run(
                        sma_period=sma, roc_period=roc,
                        vol_multiplier=mult, breadth_window=b_win,
                        start_date='2022-01-01', end_date='2022-12-31'
                    )
                    m_2022 = calculate_metrics_dual(eq_2022, TRADING_CAP, AUTH_CAP)
                    ret_2022 = m_2022['Yearly Performance'].iloc[0]['年度報酬率'] if not m_2022['Yearly Performance'].empty else -1

                    results.append({
                        'SMA': sma, 'ROC': roc, 'Multiplier': mult, 'Breadth': b_win,
                        'Full_Trading_CAGR': m_full['Trading CAGR'],
                        'Full_Std_MaxDD': m_full['Standard MaxDD'],
                        'Full_Calmar': m_full['Trading Calmar'],
                        'Return_2022': ret_2022
                    })

    df = pd.DataFrame(results)
    df.to_csv('optimization_results_4d.csv', index=False)
    print("搜索完成，結果儲存至 optimization_results_4d.csv")

if __name__ == "__main__":
    main()
