import pandas as pd
import numpy as np
import nbformat as nbf
from backtest_v2 import clean_data, BacktesterV2, calculate_metrics

def clean_and_save_data(input_file, output_file):
    print(f"Cleaning data from {input_file} and saving to {output_file}...")
    # Load data
    prices_df = pd.read_excel(input_file, sheet_name='還原收盤價', header=None)
    volumes_df = pd.read_excel(input_file, sheet_name='成交量', header=None)

    # Extract data part (starting from row 2, column 1)
    prices_data = prices_df.iloc[2:, 1:].astype(float)
    volumes_data = volumes_df.iloc[2:, 1:].astype(float)

    # Fill NaNs
    # Prices: ffill then bfill
    prices_data = prices_data.ffill().bfill()
    # Volumes: fill with 0
    volumes_data = volumes_data.fillna(0)

    # Put back into dataframes
    prices_df.iloc[2:, 1:] = prices_data
    volumes_df.iloc[2:, 1:] = volumes_data

    # Save to output_file
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        prices_df.to_excel(writer, sheet_name='還原收盤價', index=False, header=False)
        volumes_df.to_excel(writer, sheet_name='成交量', index=False, header=False)
    print(f"Successfully saved cleaned data to {output_file}")

def main():
    RAW_DATA = '資料.xlsx'
    CLEAN_DATA = '資料-1.xlsx'

    # Step 1: Clean data
    clean_and_save_data(RAW_DATA, CLEAN_DATA)

    # Step 2: Load cleaned data for backtest
    SMA_PERIOD = 303
    ROC_PERIOD = 14
    STOP_LOSS_PCT = 0.0999
    REBALANCE = 9
    INITIAL_CAPITAL = 30000000

    print(f"Loading data from {CLEAN_DATA}...")
    prices, volumes, code_to_name = clean_data(CLEAN_DATA)

    print(f"Running backtest (SMA={SMA_PERIOD}, ROC={ROC_PERIOD}, SL={STOP_LOSS_PCT*100:.2f}%, Reb={REBALANCE})...")
    bt = BacktesterV2(prices, volumes, code_to_name, initial_capital=INITIAL_CAPITAL)

    # Run full period backtest
    eq_df, trades, hold, trades2, daily = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE, 'peak', 10)

    # Performance for full period (2019-01-01 to 2026-03-31)
    mask_full = (eq_df['日期'] >= '2019-01-01') & (eq_df['日期'] <= '2026-03-31')
    res_full = eq_df[mask_full]
    cagr_full, mdd_full, calmar_full, total_ret_full = calculate_metrics(res_full)

    # Performance for 2026 (2026-01-01 to 2026-03-31)
    mask_2026 = (eq_df['日期'] >= '2026-01-01') & (eq_df['日期'] <= '2026-03-31')
    res_2026 = eq_df[mask_2026]
    cagr_2026, mdd_2026, calmar_2026, total_ret_2026 = calculate_metrics(res_2026)

    # Trade count for 2026
    trades_2026 = trades[(trades['訊號日期'] >= '2026-01-01') & (trades['訊號日期'] <= '2026-03-31')]
    trade_count_2026 = len(trades_2026[trades_2026['狀態'].isin(['買進', '賣出'])])

    # Output Excel
    OUTPUT_EXCEL = 'trendstrategy_results_equityV1.xlsx'
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

        # New 2026 sheet
        pd.DataFrame([
            {'項目': '2026年化報酬率 (CAGR)', '數值': f"{cagr_2026:.2%}"},
            {'項目': '2026最大回撤 (MaxDD)', '數值': f"{mdd_2026:.2%}"},
            {'項目': '2026 Calmar Ratio', '數值': f"{calmar_2026:.2f}"},
            {'項目': '2026交易筆數', '數值': trade_count_2026}
        ]).to_excel(writer, sheet_name='2026', index=False)

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

    # Output MD
    OUTPUT_MD = 'reproduce_equityV1.md'
    md_content = f"""# Asset Class Trend Following 策略回測報告 (equityV1)

## 策略說明
本報告彙整使用參數 **SMA={SMA_PERIOD}, ROC={ROC_PERIOD}, Stop Loss={STOP_LOSS_PCT*100:.2f}%, Rebalance={REBALANCE}天** 的回測結果。
此版本針對包含 2026.1.2 - 2026.3.31 的更新資料進行回測。

- **參數設定**：
  - SMA 週期：{SMA_PERIOD}
  - ROC 週期：{ROC_PERIOD}
  - 停損比例：{STOP_LOSS_PCT*100:.2f}% (最高價回落)
  - 再平衡週期：{REBALANCE} 天
- **核心濾網**：
  - 成交金額 > 3,000 萬 (30M)
  - 價格高於 5, 10, 20 日均線

---

## 績效表現 (2019.01.01 – 2026.03.31)
- **年化報酬率 (CAGR)**：**{cagr_full:.2%}**
- **最大回撤 (MaxDD)**：**{mdd_full:.2%}**
- **Calmar Ratio**：**{calmar_full:.2f}**
- **總報酬率**：**{total_ret_full:.2%}**

---

## 2026年 績效表現 (2026.01.01 - 2026.03.31)
- **CAGR**: {cagr_2026:.2%}
- **MaxDD**: {mdd_2026:.2%}
- **Calmar Ratio**: {calmar_2026:.2f}
- **交易筆數**: {trade_count_2026}

---

## 相關檔案
- `trendstrategy_results_equityV1.xlsx`：詳細回測數據。
- `trendstrategy_equityV1.ipynb`：回測實作程式碼。
- `資料-1.xlsx`：清洗後的數據檔。
"""
    with open(OUTPUT_MD, 'w', encoding='utf-8') as f:
        f.write(md_content)

    # Output Jupyter Notebook
    nb = nbf.v4.new_notebook()
    nb.cells.append(nbf.v4.new_markdown_cell(f"# Asset Class Trend Following 策略回測 (equityV1)\n\n本報告針對 **2019-01-01 - 2026-03-31** 期間進行回測。"))

    code_block = f"""import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from backtest_v2 import clean_data, BacktesterV2, calculate_metrics

# 1. 資料讀取與清洗
# 註：此處假設 '{CLEAN_DATA}' 已存在，若需重新清洗可執行 generate_equityV1.py
prices, volumes, code_to_name = clean_data('{CLEAN_DATA}')

# 2. 設定參數與執行回測
sma, roc, sl, reb = {SMA_PERIOD}, {ROC_PERIOD}, {STOP_LOSS_PCT}, {REBALANCE}
bt = BacktesterV2(prices, volumes, code_to_name)
eq, trades, hold, trades2, daily = bt.run(sma, roc, sl, reb, 'peak', 10)

# 3. 篩選指定期間績效
mask = (eq['日期'] >= '2019-01-01') & (eq['日期'] <= '2026-03-31')
res_p = eq[mask]

# 4. 繪製權益曲線
plt.figure(figsize=(12, 6))
plt.plot(res_p['日期'], res_p['權益'])
plt.title(f'Equity Curve (equityV1)')
plt.grid(True)
plt.show()

# 5. 計算績效指標
cagr, mdd, calmar, total_ret = calculate_metrics(res_p)
print(f"年化報酬率 (CAGR): {{cagr:.2%}}")
print(f"最大回撤 (MaxDD): {{mdd:.2%}}")
print(f"Calmar Ratio: {{calmar:.2f}}")

# 6. 2026 績效
mask_2026 = (eq['日期'] >= '2026-01-01') & (eq['日期'] <= '2026-03-31')
res_2026 = eq[mask_2026]
cagr_2026, mdd_2026, calmar_2026, total_ret_2026 = calculate_metrics(res_2026)
print(\"\\n2026 績效 (2026.01.01 - 2026.03.31):\")
print(f\"CAGR: {{cagr_2026:.2%}}\")
print(f\"MaxDD: {{mdd_2026:.2%}}\")
print(f\"Calmar Ratio: {{calmar_2026:.2f}}\")
"""
    nb.cells.append(nbf.v4.new_code_cell(code_block))
    with open('trendstrategy_equityV1.ipynb', 'w', encoding='utf-8') as f:
        nbf.write(nb, f)

    print(f"Successfully generated deliverables.")

if __name__ == "__main__":
    main()
