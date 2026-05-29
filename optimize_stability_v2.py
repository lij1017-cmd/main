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

    # 4D 搜索範圍 (涵蓋使用者要求的數值)
    sma_range = [250, 303, 350]
    roc_range = [10, 12, 14, 16]
    mult_range = [2.3, 2.7, 3.0, 3.5]
    breadth_range = [200, 240, 290, 350] # 包含 200, 240, 290, 350

    wfa_periods = [
        ('2024-06-01', '2025-12-31'),
        ('2024-01-02', '2025-05-31'),
        ('2023-01-02', '2024-12-31'),
        ('2022-01-02', '2024-05-31'),
        ('2021-06-01', '2023-12-31'),
        ('2021-01-02', '2023-05-31'),
        ('2020-01-02', '2022-12-31'),
        ('2019-06-01', '2022-05-31'),
        ('2019-01-02', '2021-12-31'),
    ]

    results = []

    total = len(sma_range) * len(roc_range) * len(mult_range) * len(breadth_range)
    count = 0

    print(f"開始穩定性 4D 搜索，總計 {total} 組組合...")

    for sma in sma_range:
        for roc in roc_range:
            for mult in mult_range:
                for b_win in breadth_range:
                    count += 1
                    if count % 20 == 0:
                        print(f"進度: {count}/{total}")

                    # 1. 執行全期間獲取基本指標與 2022 報酬
                    eq_full, _, _, _ = bt.run(sma_period=sma, roc_period=roc, vol_multiplier=mult, breadth_window=b_win, use_breadth_weight=True)
                    m_full = calculate_metrics_dual(eq_full, TRADING_CAP, AUTH_CAP)
                    ret_2022 = m_full['Yearly Performance'].loc[2022, '年度報酬率'] if 2022 in m_full['Yearly Performance'].index else -1

                    # 2. 執行 WFA 期間計算穩定性
                    wfa_cagrs = []
                    for start, end in wfa_periods:
                        eq_wfa, _, _, _ = bt.run(sma, roc, vol_multiplier=mult, breadth_window=b_win, start_date=start, end_date=end, use_breadth_weight=True)
                        m_wfa = calculate_metrics_dual(eq_wfa, TRADING_CAP, AUTH_CAP)
                        wfa_cagrs.append(m_wfa['Trading CAGR'])

                    cagr_std = np.std(wfa_cagrs)
                    cagr_min = np.min(wfa_cagrs)

                    results.append({
                        'SMA': sma, 'ROC': roc, 'Multiplier': mult, 'Breadth': b_win,
                        'Full_Trading_CAGR': m_full['Trading CAGR'],
                        'Full_Std_MaxDD': m_full['Standard MaxDD'],
                        'Full_Calmar': m_full['Trading Calmar'],
                        'Return_2022': ret_2022,
                        'WFA_CAGR_Std': cagr_std,
                        'WFA_CAGR_Min': cagr_min
                    })

    df = pd.DataFrame(results)
    df.to_csv('optimization_stability_v2.csv', index=False)
    print("搜索完成，結果儲存至 optimization_stability_v2.csv")

if __name__ == "__main__":
    main()
