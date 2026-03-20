import pandas as pd
import numpy as np
from run_wfa import Backtester, clean_data, calculate_metrics

def main():
    prices, code_to_name = clean_data('個股合-1.xlsx')
    bt = Backtester(prices, code_to_name, 30000000)

    # 定義兩組參數
    param_sets = [
        {'id': '第一組', 'sma': 87, 'roc': 54, 'sl': 0.09},
        {'id': '第二組', 'sma': 93, 'roc': 88, 'sl': 0.08}
    ]

    # 定義兩個回測區間
    periods = [
        {'name': '2021-2025', 'start': '2021-01-02', 'end': '2025-12-31'},
        {'name': '2019-2025', 'start': '2019-01-02', 'end': '2025-12-31'}
    ]

    results = []

    for ps in param_sets:
        for period in periods:
            print(f"正在執行: {ps['id']} @ {period['name']}")
            # 統一使用 6 天再平衡
            eq_df, trades = bt.run(ps['sma'], ps['roc'], ps['sl'], 6, period['start'], period['end'])
            cagr, mdd, calmar = calculate_metrics(eq_df)

            results.append({
                '參數配置': ps['id'],
                '參數詳情': f"SMA{ps['sma']} / ROC{ps['roc']} / SL{ps['sl']*100:.0f}%",
                '回測區間': period['name'],
                '年化報酬 (CAGR)': f"{cagr:.2%}",
                '最大回撤 (MaxDD)': f"{mdd:.2%}",
                '卡瑪比率 (Calmar)': f"{calmar:.2f}",
                '交易筆數': trades
            })

    res_df = pd.DataFrame(results)
    print("\n=== 兩組參數於不同區間之績效對比 ===")
    print(res_df.to_string(index=False))

if __name__ == "__main__":
    main()
