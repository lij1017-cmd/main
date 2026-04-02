import pandas as pd
import numpy as np
import nbformat as nbf
from backtest_v2 import clean_data, BacktesterV2, calculate_metrics

def main():
    DATA_FILE = '資料-1.xlsx'
    SMA_PERIOD = 303
    ROC_PERIOD = 14
    STOP_LOSS_PCT = 0.0999
    REBALANCE = 9
    INITIAL_CAPITAL = 30000000

    print(f"Loading data from {DATA_FILE}...")
    prices, volumes, code_to_name = clean_data(DATA_FILE)

    print(f"Running backtest (SMA={SMA_PERIOD}, ROC={ROC_PERIOD}, SL={STOP_LOSS_PCT*100:.2f}%, Reb={REBALANCE})...")
    bt = BacktesterV2(prices, volumes, code_to_name, initial_capital=INITIAL_CAPITAL)

    # Run full period backtest
    eq_df, trades, hold, trades2, daily = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE, 'peak', 10)

    # Full period metrics
    mask_full = (eq_df['日期'] >= '2019-01-01') & (eq_df['日期'] <= '2026-03-31')
    res_full = eq_df[mask_full]
    cagr_full, mdd_full, calmar_full, total_ret_full = calculate_metrics(res_full)

    # 2026 metrics
    mask_2026 = (eq_df['日期'] >= '2026-01-01') & (eq_df['日期'] <= '2026-03-31')
    res_2026 = eq_df[mask_2026]
    cagr_2026, mdd_2026, calmar_2026, total_ret_2026 = calculate_metrics(res_2026)
    trades_2026 = trades[(trades['訊號日期'] >= '2026-01-01') & (trades['訊號日期'] <= '2026-03-31')]
    trade_count_2026 = len(trades_2026[trades_2026['狀態'].isin(['買進', '賣出'])])

    # WFA periods
    periods = [
        ('2019-01-02', '2021-12-31'),
        ('2019-06-01', '2022-05-31'),
        ('2020-01-02', '2022-12-31'),
        ('2020-06-01', '2023-05-31'),
        ('2021-01-02', '2023-12-31'),
        ('2021-06-01', '2024-05-31'),
        ('2022-01-02', '2024-12-31'),
        ('2022-06-01', '2025-05-31'),
        ('2023-01-02', '2025-12-31'),
    ]
    wfa_results = []
    for start, end in periods:
        mask_wfa = (eq_df['日期'] >= start) & (eq_df['日期'] <= end)
        res_wfa = eq_df[mask_wfa]
        cagr_wfa, mdd_wfa, calmar_wfa, total_ret_wfa = calculate_metrics(res_wfa)
        wfa_results.append({
            '區間': f"{start} ~ {end}",
            'CAGR': f"{cagr_wfa:.2%}",
            'MaxDD': f"{mdd_wfa:.2%}",
            'Calmar': f"{calmar_wfa:.2f}",
            '總報酬率': f"{total_ret_wfa:.2%}"
        })
    wfa_df = pd.DataFrame(wfa_results)

    # 1. Output Excel
    OUTPUT_EXCEL = 'trendstrategy_results_equityV1-1.xlsx'
    with pd.ExcelWriter(OUTPUT_EXCEL, engine='xlsxwriter') as writer:
        pd.DataFrame([
            {'項目': '年化報酬率 (CAGR)', '數值': f"{cagr_full:.2%}"},
            {'項目': '最大回撤 (MaxDD)', '數值': f"{mdd_full:.2%}"},
            {'項目': 'Calmar Ratio', '數值': f"{calmar_full:.2f}"},
            {'項目': '總報酬率', '數值': f"{total_ret_full:.2%}"},
            {'項目': '初始資金', '數值': f"{INITIAL_CAPITAL:,}"},
            {'項目': '期末淨值', '數值': f"{res_full['權益'].iloc[-1]:,.0f}"}
        ]).to_excel(writer, sheet_name='Summary', index=False)

        eq_df.to_excel(writer, sheet_name='Equity_Curve', index=False)
        hold.to_excel(writer, sheet_name='Equity_Hold', index=False)
        trades.to_excel(writer, sheet_name='Trades', index=False)
        trades2.to_excel(writer, sheet_name='Trades2', index=False)
        daily.to_excel(writer, sheet_name='Daily', index=False)

        pd.DataFrame([
            {'項目': '2026年化報酬率 (CAGR)', '數值': f"{cagr_2026:.2%}"},
            {'項目': '2026最大回撤 (MaxDD)', '數值': f"{mdd_2026:.2%}"},
            {'項目': '2026 Calmar Ratio', '數值': f"{calmar_2026:.2f}"},
            {'項目': '2026交易筆數', '數值': trade_count_2026}
        ]).to_excel(writer, sheet_name='2026', index=False)

        wfa_df.to_excel(writer, sheet_name='WFA', index=False)

        # Chart
        workbook = writer.book
        curves_sheet = writer.sheets['Equity_Curve']
        chart = workbook.add_chart({'type': 'line'})
        max_row = len(eq_df)
        chart.add_series({
            'name': 'Equity Curve',
            'categories': ['Equity_Curve', 1, 0, max_row, 0],
            'values': ['Equity_Curve', 1, 1, max_row, 1],
        })
        chart.set_title({'name': 'Equity Curve (2019-2026/03)'})
        curves_sheet.insert_chart('E2', chart)

    # 2. Output MD
    OUTPUT_MD = 'reproduce_equityV1-1.md'
    md_content = f"""# Asset Class Trend Following 策略回測報告 (equityV1-1)

## 策略說明
本報告彙整使用參數 **SMA={SMA_PERIOD}, ROC={ROC_PERIOD}, Stop Loss={STOP_LOSS_PCT*100:.2f}%, Rebalance={REBALANCE}天** 的回測結果。
此版本包含 2019.01.01 - 2026.03.31 的全期間回測與多區間 Walk-Forward Analysis (WFA)。

- **參數設定**：
  - SMA 週期：{SMA_PERIOD}
  - ROC 週期：{ROC_PERIOD}
  - 停損比例：{STOP_LOSS_PCT*100:.2f}% (最高價回落)
  - 再平衡週期：{REBALANCE} 天

---

## 全期間績效 (2019.01.01 – 2026.03.31)
- **年化報酬率 (CAGR)**：**{cagr_full:.2%}**
- **最大回撤 (MaxDD)**：**{mdd_full:.2%}**
- **Calmar Ratio**：**{calmar_full:.2f}**

---

## Walk-Forward Analysis (WFA)
{wfa_df.to_markdown(index=False)}

---

## 2026年 績效表現 (2026.01.01 - 2026.03.31)
- **CAGR**: {cagr_2026:.2%}
- **MaxDD**: {mdd_2026:.2%}
- **Calmar Ratio**: {calmar_2026:.2f}
- **交易筆數**: {trade_count_2026}

---

## 相關檔案
- `trendstrategy_results_equityV1-1.xlsx`：詳細回測數據與 WFA 成績表。
- `trendstrategy_equityV1-1.ipynb`：回測實作程式碼。
"""
    with open(OUTPUT_MD, 'w', encoding='utf-8') as f:
        f.write(md_content)

    print(f"Successfully generated deliverables for equityV1-1.")

if __name__ == "__main__":
    main()
