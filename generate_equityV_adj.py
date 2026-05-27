import pandas as pd
import numpy as np
import nbformat as nbf
import os
from backtest_equityV_adj import clean_data, BacktesterAdjusted, calculate_metrics_adj

def main():
    DATA_FILE = '樣本集-1.xlsx'
    AUTHORIZED_CAPITAL = 60000000
    TRADING_CAPITAL = 12000000

    SMA_PERIOD = 303
    ROC_PERIOD = 14
    STOP_LOSS_PCT = 0.0999
    REBALANCE = 9

    print(f"Loading data from {DATA_FILE}...")
    prices, volumes, code_to_name = clean_data(DATA_FILE)

    print(f"Running backtest (SMA={SMA_PERIOD}, ROC={ROC_PERIOD}, SL={STOP_LOSS_PCT*100:.2f}%, Reb={REBALANCE})...")
    bt = BacktesterAdjusted(prices, volumes, code_to_name,
                            authorized_capital=AUTHORIZED_CAPITAL,
                            trading_capital=TRADING_CAPITAL)

    # 執行全期間回測
    eq_df, trades, hold, trades2, daily = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE)

    # 1. 回測模式 (全期間)
    mask_full = (eq_df['日期'] >= '2019-01-01') & (eq_df['日期'] <= '2025-12-31')
    res_full = eq_df[mask_full]
    cagr_full, mdd_full, fmdd_full, calmar_full, total_ret_full = calculate_metrics_adj(res_full, AUTHORIZED_CAPITAL)

    print("\n--- 回測模式 (全期間 2019-2025) ---")
    print(f"CAGR: {cagr_full:.2%}")
    print(f"MaxDD (投資報告標準): {mdd_full:.2%}")
    print(f"Fixed Base MDD (實際資金風險控管): {fmdd_full:.2%}")
    print(f"Calmar Ratio: {calmar_full:.2f}")
    print(f"Total Return: {total_ret_full:.2%}")

    # 2. 實際交易模式 (年度 breakdown)
    yearly_results = []
    years = sorted(res_full['日期'].dt.year.unique())
    for year in years:
        mask_year = (res_full['日期'].dt.year == year)
        res_year = res_full[mask_year]
        # 年度模式要求：年初損益歸零，但持有部位延續
        # 這意味著計算該年度 return 時，分母應是該年初的權益值
        y_total_ret = (res_year['權益'].iloc[-1] / res_year['權益'].iloc[0]) - 1
        # 年度 MDD
        y_peak = res_year['權益'].cummax()
        y_mdd = ((res_year['權益'] - y_peak) / y_peak).min()
        y_fmdd = ((res_year['權益'] - y_peak) / AUTHORIZED_CAPITAL).min()

        yearly_results.append({
            '年度': year,
            '年度報酬率': f"{y_total_ret:.2%}",
            '年度 MaxDD': f"{y_mdd:.2%}",
            '年度固定基準 MDD': f"{y_fmdd:.2%}"
        })

    yearly_df = pd.DataFrame(yearly_results)
    print("\n--- 實際交易模式 (各年度績效) ---")
    print(yearly_df.to_string(index=False))

    # 輸出 Excel
    OUTPUT_EXCEL = 'trendstrategy_results_equityV(調).xlsx'
    with pd.ExcelWriter(OUTPUT_EXCEL, engine='xlsxwriter') as writer:
        # Summary Sheet
        summary_data = [
            {'模式': '回測模式 (2019-2025)', '項目': '年化報酬率 (CAGR)', '數值': f"{cagr_full:.2%}"},
            {'模式': '回測模式 (2019-2025)', '項目': '最大回撤 (MaxDD)', '數值': f"{mdd_full:.2%}"},
            {'模式': '回測模式 (2019-2025)', '項目': '固定基準 MDD', '數值': f"{fmdd_full:.2%}"},
            {'模式': '回測模式 (2019-2025)', '項目': 'Calmar Ratio', '數值': f"{calmar_full:.2f}"},
            {'模式': '回測模式 (2019-2025)', '項目': '總報酬率', '數值': f"{total_ret_full:.2%}"},
            {'模式': '回測模式 (2019-2025)', '項目': '授權資金', '數值': f"{AUTHORIZED_CAPITAL:,}"},
            {'模式': '回測模式 (2019-2025)', '項目': '使用資金', '數值': f"{TRADING_CAPITAL:,}"}
        ]
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)

        # 年度績效也放在 Summary 下方或新 Sheet
        yearly_df.to_excel(writer, sheet_name='Summary', index=False, startrow=len(summary_data) + 2)

        eq_df.to_excel(writer, sheet_name='Equity_Curve', index=False)
        hold.to_excel(writer, sheet_name='Equity_Hold', index=False)
        trades.to_excel(writer, sheet_name='Trades', index=False)
        trades2.to_excel(writer, sheet_name='Trades2', index=False)
        daily.to_excel(writer, sheet_name='Daily', index=False)

        # 繪製圖表
        workbook = writer.book
        curves_sheet = writer.sheets['Equity_Curve']
        chart = workbook.add_chart({'type': 'line'})
        max_row = len(eq_df)
        chart.add_series({
            'name': 'Equity Curve',
            'categories': ['Equity_Curve', 1, 0, max_row, 0],
            'values': ['Equity_Curve', 1, 1, max_row, 1],
        })
        chart.set_title({'name': 'Equity Curve (equityV(調))'})
        curves_sheet.insert_chart('F2', chart)

    # 輸出 MD
    OUTPUT_MD = 'reproduce_equityV(調).md'
    md_content = f"""# Asset Class Trend Following 策略回測報告 (equityV(調))

## 策略說明 (調整版)
本報告依據使用者需求調整資金配置與 MDD 計算邏輯後的回測結果。

- **資金設定**：
  - 授權資金：{AUTHORIZED_CAPITAL:,} TWD
  - 使用資金：{TRADING_CAPITAL:,} TWD
  - 持股配置：平均分配至 3 檔，每檔 {TRADING_CAPITAL/3:,.0f} TWD
- **參數設定**：
  - SMA 週期：{SMA_PERIOD}
  - ROC 週期：{ROC_PERIOD}
  - 停損比例：{STOP_LOSS_PCT*100:.2f}%
  - 再平衡週期：{REBALANCE} 天
- **特殊規則**：
  - 買入時若金額不足 1000 股仍買入並註記。
  - 再平衡續抱時維持原部位。
  - 固定基準 MDD 以 {AUTHORIZED_CAPITAL:,} 萬元為分母計算。

---

## 績效表現 - 回測模式 (2019 – 2025 全期間)
- **年化報酬率 (CAGR)**：**{cagr_full:.2%}**
- **最大回撤 (MaxDD - 投資報告標準)**：**{mdd_full:.2%}**
- **固定基準 MDD (實際資金風險控管)**：**{fmdd_full:.2%}**
- **Calmar Ratio**：**{calmar_full:.2f}**
- **總報酬率**：**{total_ret_full:.2%}**

---

## 績效表現 - 實際交易模式 (年度分解)
{yearly_df.to_markdown(index=False)}

---

## 相關檔案
- `{OUTPUT_EXCEL}`：詳細回測數據。
- `trendstrategy_equityV(調).ipynb`：回測實作程式碼。
"""
    with open(OUTPUT_MD, 'w', encoding='utf-8') as f:
        f.write(md_content)

    # 輸出 Jupyter Notebook
    nb = nbf.v4.new_notebook()
    nb.cells.append(nbf.v4.new_markdown_cell(f"# Asset Class Trend Following 策略回測 (equityV(調))\n\n本報告針對調整後的資金與 MDD 邏輯進行回測。"))

    code_block = f"""import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from backtest_equityV_adj import clean_data, BacktesterAdjusted, calculate_metrics_adj

# 1. 資料讀取
prices, volumes, code_to_name = clean_data('{DATA_FILE}')

# 2. 設定參數與執行回測
sma, roc, sl, reb = {SMA_PERIOD}, {ROC_PERIOD}, {STOP_LOSS_PCT}, {REBALANCE}
bt = BacktesterAdjusted(prices, volumes, code_to_name,
                        authorized_capital={AUTHORIZED_CAPITAL},
                        trading_capital={TRADING_CAPITAL})
eq, trades, hold, trades2, daily = bt.run(sma, roc, sl, reb)

# 3. 篩選指定期間績效
mask = (eq['日期'] >= '2019-01-01') & (eq['日期'] <= '2025-12-31')
res_p = eq[mask]

# 4. 繪製權益曲線
plt.figure(figsize=(12, 6))
plt.plot(res_p['日期'], res_p['權益'])
plt.title(f'Equity Curve (equityV(調))')
plt.grid(True)
plt.show()

# 5. 計算績效指標 (回測模式)
cagr, mdd, fmdd, calmar, total_ret = calculate_metrics_adj(res_p, {AUTHORIZED_CAPITAL})
print(f"年化報酬率 (CAGR): {{cagr:.2%}}")
print(f"最大回撤 (MaxDD - 投資報告標準): {{mdd:.2%}}")
print(f"固定基準 MDD (實際資金風險控管): {{fmdd:.2%}}")
print(f"Calmar Ratio: {{calmar:.2f}}")
print(f"總報酬率: {{total_ret:.2%}}")

# 6. 年度績效 (實際交易模式)
res_p['Year'] = res_p['日期'].dt.year
for year, group in res_p.groupby('Year'):
    y_ret = (group['權益'].iloc[-1] / group['權益'].iloc[0]) - 1
    y_peak = group['權益'].cummax()
    y_mdd = ((group['權益'] - y_peak) / y_peak).min()
    print(f"{{year}} 年度報酬: {{y_ret:.2%}}, 年度 MaxDD: {{y_mdd:.2%}}")
"""
    nb.cells.append(nbf.v4.new_code_cell(code_block))
    with open('trendstrategy_equityV(調).ipynb', 'w', encoding='utf-8') as f:
        nbf.write(nb, f)

    print(f"\nSuccessfully generated deliverables.")

if __name__ == "__main__":
    main()
