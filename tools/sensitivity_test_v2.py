import pandas as pd
import numpy as np
from run_wfa import Backtester, clean_data, calculate_metrics

def run_trial(bt, sma, roc, sl, start_date, end_date):
    eq, trades = bt.run(int(sma), int(roc), float(sl), 6, start_date, end_date)
    cagr, mdd, calmar = calculate_metrics(eq)
    return cagr, mdd, calmar, trades

def main():
    prices, code_to_name = clean_data('個股合-1.xlsx')
    bt = Backtester(prices, code_to_name, 30000000)

    start_date = '2022-01-02'
    end_date = '2025-12-31'

    # 測試組別 (以新最佳參數 SMA 17, ROC 13, SL 10% 為基準)
    trials = [
        {'name': 'Optimized', 'sma': 17, 'roc': 13, 'sl': 0.10},
        {'name': 'Trial A', 'sma': 19, 'roc': 15, 'sl': 0.095},
        {'name': 'Trial B', 'sma': 15, 'roc': 11, 'sl': 0.10},
        {'name': 'Trial C', 'sma': 19, 'roc': 11, 'sl': 0.090},
    ]

    results = []
    for t in trials:
        print(f"Running {t['name']}: SMA={t['sma']}, ROC={t['roc']}, SL={t['sl']*100:.1f}%")
        cagr, mdd, calmar, trades = run_trial(bt, t['sma'], t['roc'], t['sl'], start_date, end_date)
        results.append({
            '組別': t['name'],
            'SMA': t['sma'],
            'ROC': t['roc'],
            'SL': f"{t['sl']*100:.1f}%",
            'CAGR': f"{cagr:.2%}",
            'MaxDD': f"{mdd:.2%}",
            'Calmar': f"{calmar:.2f}",
            '交易筆數': trades
        })

    res_df = pd.DataFrame(results)
    print("\n=== 參數微調敏感度測試 (2022-2025) ===")
    print(res_df.to_string(index=False))

if __name__ == "__main__":
    main()
