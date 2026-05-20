import pandas as pd
import numpy as np
import nbformat as nbf
from backtest_v2 import clean_data, BacktesterV2, calculate_metrics

def main():
    DATA_FILE = '樣本集-1.xlsx'
    SMA_PERIOD = 303
    ROC_PERIOD = 14
    STOP_LOSS_PCT = 0.0999
    REBALANCE = 9
    INITIAL_CAPITAL = 30000000
    FIXED_BASE_CAPITAL = 150000000

    print(f"Loading data from {DATA_FILE}...")
    prices, volumes, code_to_name = clean_data(DATA_FILE)

    print(f"Running backtest (SMA={SMA_PERIOD}, ROC={ROC_PERIOD}, SL={STOP_LOSS_PCT*100:.2f}%, Reb={REBALANCE})...")
    bt = BacktesterV2(prices, volumes, code_to_name, initial_capital=INITIAL_CAPITAL)

    # Run full period backtest
    eq_df, trades, hold, trades2, daily = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE, 'peak', 10)

    # Filter period 2019.01.01 – 2025.12.31
    mask = (eq_df['日期'] >= '2019-01-01') & (eq_df['日期'] <= '2025-12-31')
    res = eq_df[mask].copy()

    # Calculate Standard Metrics
    cagr, mdd_std, calmar_std, total_ret = calculate_metrics(res)

    # Calculate Fixed Base MDD
    # Fixed Base MDD = (Current Equity - Peak Equity) / 150,000,000
    # We need rolling peak equity
    res['Peak_Equity'] = res['權益'].cummax()
    res['Fixed_Base_Drawdown'] = (res['權益'] - res['Peak_Equity']) / FIXED_BASE_CAPITAL
    mdd_fixed = res['Fixed_Base_Drawdown'].min()

    # Recalculate Calmar for Fixed Base MDD
    calmar_fixed = cagr / abs(mdd_fixed) if mdd_fixed != 0 else 0

    print("\n--- Backtest Results (2019-2025) ---")
    print(f"Total Return: {total_ret:.2%}")
    print(f"CAGR: {cagr:.2%}")
    print(f"Standard MaxDD: {mdd_std:.2%}")
    print(f"Fixed Base MaxDD (Base 150M): {mdd_fixed:.2%}")
    print(f"Calmar Ratio (Standard): {calmar_std:.2f}")
    print(f"Calmar Ratio (Fixed Base): {calmar_fixed:.2f}")

    # Output Excel
    OUTPUT_EXCEL = 'trendstrategy_results_equityV(固).xlsx'
    with pd.ExcelWriter(OUTPUT_EXCEL, engine='xlsxwriter') as writer:
        pd.DataFrame([
            {'項目': '年化報酬率 (CAGR)', '數值': f"{cagr:.2%}"},
            {'項目': '標準最大回撤 (Standard MaxDD)', '數值': f"{mdd_std:.2%}"},
            {'項目': '固定基準最大回撤 (Fixed Base MaxDD)', '數值': f"{mdd_fixed:.2%}"},
            {'項目': 'Calmar Ratio (標準)', '數值': f"{calmar_std:.2f}"},
            {'項目': 'Calmar Ratio (固定基準)', '數值': f"{calmar_fixed:.2f}"},
            {'項目': '總報酬率', '數值': f"{total_ret:.2%}"},
            {'項目': '初始資金', '數值': f"{INITIAL_CAPITAL:,}"},
            {'項目': '授權基準資金', '數值': f"{FIXED_BASE_CAPITAL:,}"},
            {'項目': '期末淨值', '數值': f"{res['權益'].iloc[-1]:,.0f}"}
        ]).to_excel(writer, sheet_name='Summary', index=False)

        # Update res for Excel output
        res_output = res[['日期', '權益', '回撤(Drawdown)', 'Fixed_Base_Drawdown']].copy()
        res_output.columns = ['日期', '權益', '標準回撤(Drawdown)', '固定基準回撤']
        res_output.to_excel(writer, sheet_name='Equity_Curve', index=False)

        hold.to_excel(writer, sheet_name='Equity_Hold', index=False)
        trades.to_excel(writer, sheet_name='Trades', index=False)
        trades2.to_excel(writer, sheet_name='Trades2', index=False)
        daily.to_excel(writer, sheet_name='Daily', index=False)

        workbook = writer.book
        curves_sheet = writer.sheets['Equity_Curve']
        chart = workbook.add_chart({'type': 'line'})
        max_row = len(res_output)
        chart.add_series({
            'name': 'Equity Curve',
            'categories': ['Equity_Curve', 1, 0, max_row, 0],
            'values': ['Equity_Curve', 1, 1, max_row, 1],
        })
        chart.set_title({'name': 'Equity Curve (equityV(固))'})
        curves_sheet.insert_chart('F2', chart)

    # Output MD
    OUTPUT_MD = 'reproduce_equityV(固).md'
    md_content = f"""# Asset Class Trend Following 策略回測報告 (equityV(固))

## 策略說明
本報告彙整使用參數 **SMA={SMA_PERIOD}, ROC={ROC_PERIOD}, Stop Loss={STOP_LOSS_PCT*100:.2f}%, Rebalance={REBALANCE}天** 的回測結果。
此版本修正了 MDD 計算邏輯，新增「固定基準 MDD」，以授權資金 {FIXED_BASE_CAPITAL:,} 元為基準。

- **參數設定**：
  - SMA 週期：{SMA_PERIOD}
  - ROC 週期：{ROC_PERIOD}
  - 停損比例：{STOP_LOSS_PCT*100:.2f}% (最高價回落)
  - 再平衡週期：{REBALANCE} 天
- **核心濾網**：
  - 成交金額 > 3,000 萬 (30M)
  - 價格高於 5, 10, 20 日均線

---

## 績效表現 (2019.01.01 – 2025.12.31)
- **年化報酬率 (CAGR)**：**{cagr:.2%}**
- **總報酬率**：**{total_ret:.2%}**
- **最大回撤 (Standard MaxDD)**：**{mdd_std:.2%}**
- **固定基準最大回撤 (Fixed Base MaxDD)**：**{mdd_fixed:.2%}**
- **Calmar Ratio (標準)**：**{calmar_std:.2f}**
- **Calmar Ratio (固定基準)**：**{calmar_fixed:.2f}**

---

## 結果分析
本次回測特別引入了「固定基準 MDD」。標準 MaxDD 為 **{mdd_std:.2%}**，而基於 {FIXED_BASE_CAPITAL/1e8:.1f} 億授權資金計算的固定基準 MaxDD 為 **{mdd_fixed:.2%}**。
這種計算方式有助於比較「投資報告標準」與「實際資金風險控管」，在資產規模較小時，固定基準 MDD 會顯著低於標準 MDD，反映了相對於總授權額度的實際風險暴露。

---

## 相關檔案
- `{OUTPUT_EXCEL}`：詳細回測數據。
- `trendstrategy_equityV(固).ipynb`：回測實作程式碼。
"""
    with open(OUTPUT_MD, 'w', encoding='utf-8') as f:
        f.write(md_content)

    # Output Jupyter Notebook
    nb = nbf.v4.new_notebook()
    nb.cells.append(nbf.v4.new_markdown_cell(f"# Asset Class Trend Following 策略回測 (equityV(固))\n\n本報告針對 **2019-01-01 - 2025-12-31** 期間進行回測，並比較兩種 MDD 計算邏輯。"))

    code_block = f"""import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from backtest_v2 import clean_data, BacktesterV2, calculate_metrics

# 1. 資料讀取與清洗
prices, volumes, code_to_name = clean_data('{DATA_FILE}')

# 2. 設定參數與執行回測
sma, roc, sl, reb = {SMA_PERIOD}, {ROC_PERIOD}, {STOP_LOSS_PCT}, {REBALANCE}
initial_capital = {INITIAL_CAPITAL}
fixed_base_capital = {FIXED_BASE_CAPITAL}

bt = BacktesterV2(prices, volumes, code_to_name, initial_capital=initial_capital)
eq, trades, hold, trades2, daily = bt.run(sma, roc, sl, reb, 'peak', 10)

# 3. 篩選指定期間績效
mask = (eq['日期'] >= '2019-01-01') & (eq['日期'] <= '2025-12-31')
res = eq[mask].copy()

# 4. 計算固定基準 MDD
res['Peak_Equity'] = res['權益'].cummax()
res['Fixed_Base_Drawdown'] = (res['權益'] - res['Peak_Equity']) / fixed_base_capital
mdd_fixed = res['Fixed_Base_Drawdown'].min()

# 5. 繪製權益曲線
plt.figure(figsize=(12, 6))
plt.plot(res['日期'], res['權益'])
plt.title(f'Equity Curve (equityV(固))')
plt.grid(True)
plt.show()

# 6. 計算與顯示績效指標
cagr, mdd_std, calmar_std, total_ret = calculate_metrics(res)
calmar_fixed = cagr / abs(mdd_fixed) if mdd_fixed != 0 else 0

print(f"年化報酬率 (CAGR): {{cagr:.2%}}")
print(f"總報酬率: {{total_ret:.2%}}")
print(f"標準最大回撤 (Standard MaxDD): {{mdd_std:.2%}}")
print(f"固定基準最大回撤 (Fixed Base MaxDD): {{mdd_fixed:.2%}}")
print(f"Calmar Ratio (標準): {{calmar_std:.2f}}")
print(f"Calmar Ratio (固定基準): {{calmar_fixed:.2f}}")
"""
    nb.cells.append(nbf.v4.new_code_cell(code_block))
    with open('trendstrategy_equityV(固).ipynb', 'w', encoding='utf-8') as f:
        nbf.write(nb, f)

    print(f"Successfully generated deliverables.")

if __name__ == "__main__":
    main()
