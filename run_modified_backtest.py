import pandas as pd
from backtest_modified import clean_data, BacktesterModified, calculate_metrics

def main():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')

    # 參數設定 (SMA=303, ROC=14, SL=9.99%, Reb=9)
    sma, roc_p, sl, reb = 303, 14, 0.0999, 9

    bt = BacktesterModified(prices, volumes, code_to_name)
    eq, trades, hold, trades2, daily = bt.run(sma, roc_p, sl, reb, 'peak', 10)

    # 篩選指定期間績效 (2019.01.01 – 2025.12.31)
    mask = (eq['日期'] >= '2019-01-01') & (eq['日期'] <= '2025-12-31')
    res_p = eq[mask]

    cagr, mdd, calmar, total_ret = calculate_metrics(res_p)
    print(f"年化報酬率 (CAGR): {cagr:.2%}")
    print(f"最大回撤 (MaxDD): {mdd:.2%}")
    print(f"Calmar Ratio: {calmar:.2f}")

    # 產出 Excel
    OUTPUT_EXCEL = 'trendstrategy_results_equityV-刪成交金額.xlsx'
    summary_df = pd.DataFrame([
        {'項目': '年化報酬率 (CAGR)', '數值': f"{cagr:.2%}"},
        {'項目': '最大回撤 (MaxDD)', '數值': f"{mdd:.2%}"},
        {'項目': 'Calmar Ratio', '數值': f"{calmar:.2f}"},
        {'項目': '總報酬率', '數值': f"{total_ret:.2%}"},
        {'項目': '版本', '數值': 'equityV-刪成交金額 (SMA303/ROC14/SL9.99%/Reb9)'}
    ])

    with pd.ExcelWriter(OUTPUT_EXCEL, engine='xlsxwriter') as writer:
        summary_df.to_excel(writer, sheet_name='SUMMARY', index=False)

    print(f"Excel file {OUTPUT_EXCEL} generated with 'SUMMARY' sheet.")

if __name__ == "__main__":
    main()
