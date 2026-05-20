import pandas as pd
import numpy as np
import nbformat as nbf
from backtest_atr_fixed import clean_data, BacktesterATR, calculate_metrics
import os

# ==========================================
# 1. 核心設定與回測
# ==========================================
def main():
    filepath = '樣本集-1.xlsx'
    prices, volumes, code_to_name = clean_data(filepath)
    # 初始資金改為 150,000,000
    bt = BacktesterATR(prices, volumes, code_to_name, initial_capital=150000000)

    # 參數設定
    sma_period = 303
    roc_period = 10
    atr_period = 15
    atr_multiplier = 4.3
    rebalance_interval = 9
    mkt_t = 0.42
    mkt_s = 14
    mkt_w = 290

    print("正在執行方案 B (固定基準 MDD 版) 回測...")
    eq, trades, hold, trades2, daily = bt.run(
        sma_period=sma_period,
        roc_period=roc_period,
        stop_loss_type='atr',
        atr_period=atr_period,
        atr_multiplier=atr_multiplier,
        rebalance_interval=rebalance_interval,
        use_market_filter=True,
        breadth_threshold=mkt_t,
        mkt_sma_window=mkt_s,
        breadth_window=mkt_w
    )

    cagr, mdd, mdd_fixed, calmar, total_ret = calculate_metrics(eq)

    # 計算年度數據
    eq['Year'] = eq['日期'].dt.year
    annual_data = []
    for year, group in eq.groupby('Year'):
        y_cagr, y_mdd, y_mdd_fixed, y_calmar, y_ret = calculate_metrics(group)
        annual_data.append({
            '年份': year, 'CAGR': f"{y_cagr:.2%}",
            'Standard MaxDD': f"{y_mdd:.2%}",
            'Fixed Base MaxDD': f"{y_mdd_fixed:.2%}",
            'Calmar': f"{y_calmar:.2f}", '年度報酬': f"{y_ret:.2%}"
        })
    df_annual = pd.DataFrame(annual_data)

    summary_df = pd.DataFrame([
        {'項目': '年化報酬率 (CAGR)', '數值': f"{cagr:.2%}"},
        {'項目': '標準最大回撤 (Standard MaxDD)', '數值': f"{mdd:.2%}"},
        {'項目': '固定基準最大回撤 (Fixed Base MaxDD)', '數值': f"{mdd_fixed:.2%}"},
        {'項目': '卡瑪比率 (Calmar Ratio)', '數值': f"{calmar:.2f}"},
        {'項目': '總報酬率 (Total Return)', '數值': f"{total_ret:.2%}"},
        {'項目': '初始授權資金', '數值': "150,000,000"},
        {'項目': '核心參數', '數值': f"SMA {sma_period}, ROC {roc_period}, Reb {rebalance_interval}"},
        {'項目': '停損機制', '數值': f"ATR (P={atr_period}, M={atr_multiplier})"},
        {'項目': '市場濾網', '數值': f"Threshold {mkt_t}, SMA {mkt_s}, Window {mkt_w}"}
    ])

    # 產出檔案
    suffix = "equityV-filter-atr(固)"
    generate_xlsx(f"{suffix}.xlsx", eq, trades, hold, trades2, daily, summary_df, df_annual)
    generate_md(f"reproduce_{suffix}.md", summary_df, df_annual, bt)
    generate_ipynb(f"trendstrategy_{suffix}.ipynb", suffix)

    print(f"完成！已產出 {suffix}.xlsx, reproduce_{suffix}.md, trendstrategy_{suffix}.ipynb")

# ==========================================
# 2. Excel 產出工具
# ==========================================
def generate_xlsx(filename, eq_df, trades, hold, trades2, daily, summary_df, annual_df):
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        annual_df.to_excel(writer, sheet_name='Summary', index=False, startrow=12)
        eq_df.to_excel(writer, sheet_name='Equity_Curve', index=False)
        hold.to_excel(writer, sheet_name='Equity_Hold', index=False)
        trades.to_excel(writer, sheet_name='Trades', index=False)
        trades2.to_excel(writer, sheet_name='Trades2', index=False)
        daily.to_excel(writer, sheet_name='Daily', index=False)

        workbook = writer.book
        curves_sheet = writer.sheets['Equity_Curve']
        chart = workbook.add_chart({'type': 'line'})
        max_row = len(eq_df)
        chart.add_series({
            'name': 'Equity Curve',
            'categories': ['Equity_Curve', 1, 0, max_row, 0],
            'values': ['Equity_Curve', 1, 1, max_row, 1],
        })
        chart.set_title({'name': 'Equity Curve (Scenario B - Fixed MDD)'})
        curves_sheet.insert_chart('H2', chart)

# ==========================================
# 3. Markdown 產出工具
# ==========================================
def generate_md(filename, summary_df, annual_df, bt):
    content = f"""# Asset Class Trend Following 策略優化報告 (equityV-filter-atr(固))

## 1. 策略說明與優化目標
本報告為 `equityV` 策略之固定基準 MDD 版本。
- **核心邏輯**：SMA 303 與 ROC 10 判定趨勢與動能。
- **停損機制**：ATR 動態停損（Period=15, Multiplier=4.3）。
- **市場濾網**：雙重確認濾網（寬度 0.42, 均線 14）。
- **MDD 計算修正**：以初始授權資金 150,000,000 元為基準。

## 2. 績效總結 (2019-2025)
{summary_df.to_markdown(index=False)}

### 年度績效明細
{annual_df.to_markdown(index=False)}

## 3. 相關檔案
- `equityV-filter-atr(固).xlsx`：詳細交易日誌與權益曲線。
- `trendstrategy_equityV-filter-atr(固).ipynb`：完整回測程式碼。
"""
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(content)

# ==========================================
# 4. Notebook 產出工具
# ==========================================
def generate_ipynb(filename, title):
    nb = nbf.v4.new_notebook()

    with open('backtest_atr_fixed.py', 'r', encoding='utf-8') as f:
        code = f.read()

    nb.cells = [
        nbf.v4.new_markdown_cell(f"# {title} 策略回測\n包含 ATR 動態停損與市場濾網邏輯，並修正 MDD 計算邏輯。"),
        nbf.v4.new_code_cell("import pandas as pd\nimport numpy as np\nimport matplotlib.pyplot as plt"),
        nbf.v4.new_code_cell(code),
        nbf.v4.new_markdown_cell("## 執行回測 (固定基準 MDD 版)"),
        nbf.v4.new_code_cell("""prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
bt = BacktesterATR(prices, volumes, code_to_name, initial_capital=150000000)
eq, trades, hold, trades2, daily = bt.run(
    sma_period=303, roc_period=10, stop_loss_type='atr',
    atr_period=15, atr_multiplier=4.3, rebalance_interval=9,
    use_market_filter=True, breadth_threshold=0.42, mkt_sma_window=14, breadth_window=290
)
metrics = calculate_metrics(eq)
print(f"CAGR: {metrics[0]:.2%}")
print(f"Standard MaxDD: {metrics[1]:.2%}")
print(f"Fixed Base MaxDD: {metrics[2]:.2%}")
print(f"Calmar Ratio: {metrics[3]:.2f}")
print(f"Total Return: {metrics[4]:.2%}")"""),
        nbf.v4.new_code_cell("""plt.figure(figsize=(12, 6))
plt.plot(eq['日期'], eq['權益'])
plt.title('Equity Curve (Fixed MDD version)')
plt.grid(True)
plt.show()""")
    ]

    with open(filename, 'w', encoding='utf-8') as f:
        nbf.write(nb, f)

if __name__ == "__main__":
    main()
