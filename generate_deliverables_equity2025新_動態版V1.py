import pandas as pd
import numpy as np
import pickle
import nbformat as nbf
from run_backtest_equity2025新_動態版V1 import Backtester, clean_data, calculate_metrics

# ==========================================
# 1. 參數設定區塊 (Global Parameters)
# ==========================================
DATA_FILE = '個股合-1.xlsx'
SMA_PERIOD = 87
ROC_PERIOD = 54
STOP_LOSS_PCT = 0.09
REBALANCE = 6
INITIAL_CAPITAL = 30000000

# ==========================================
# 2. 執行回測 (Run Backtest)
# ==========================================
prices, code_to_name = clean_data(DATA_FILE)
bt = Backtester(prices, code_to_name, INITIAL_CAPITAL)
eq_df, trades, hold, trades2, daily = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE)
cagr, mdd, calmar, ret = calculate_metrics(eq_df)

# ==========================================
# 3. 產出 Excel 報表 (Generate Excel Report)
# ==========================================
OUTPUT_EXCEL = 'trendstrategy_results_equity2025新-動態版V1.xlsx'
with pd.ExcelWriter(OUTPUT_EXCEL, engine='xlsxwriter') as writer:
    # Summary
    pd.DataFrame([
        {'項目': '年化報酬率 (CAGR)', '數值': f"{cagr:.2%}"},
        {'項目': '最大回撤 (MaxDD)', '數值': f"{mdd:.2%}"},
        {'項目': 'Calmar Ratio', '數值': f"{calmar:.2f}"},
        {'項目': '總報酬率', '數值': f"{ret:.2%}"},
        {'項目': '版本', '數值': '動態版 V1 (增強報表)'}
    ]).to_excel(writer, sheet_name='Summary', index=False)

    # Equity_Curve (含回撤)
    eq_df.to_excel(writer, sheet_name='Equity_Curve', index=False)

    # Equity_Hold (含現金、市值)
    hold.to_excel(writer, sheet_name='Equity_Hold', index=False)

    # Trades (原始紀錄)
    cols = ['訊號日期', '股票代號', '狀態', '價格', '股數', '動能值', '標的名稱', '原因', '買入手續費', '賣出手續費', '賣出交易稅', '說明']
    trades[cols].to_excel(writer, sheet_name='Trades', index=False)

    # Trades2 (成對交易)
    trades2.to_excel(writer, sheet_name='Trades2', index=False)

    # Daily (每日持股明細)
    daily.to_excel(writer, sheet_name='Daily', index=False)

# ==========================================
# 4. 產出 Markdown 報告 (Generate MD Report)
# ==========================================
OUTPUT_MD = 'reproduce_report_equity2025新-動態版V1.md'
md_content = f"""# Asset Class Trend Following 策略回測報告 (equity2025新-動態版V1)

## 需求清單 (Communication Template)
根據最新需求，本版本已完成以下報表優化：
1. **新增 Trades2 工作表**：紀錄成對交易（買入至賣出），包含進出場日期、價格、損益、報酬率及詳細進出場原因（如股價>SMA、排名外等）。
2. **Equity_Curve 優化**：新增「回撤(Drawdown)」欄位，紀錄每日權益回落幅度。
3. **Equity_Hold 優化**：新增「現金」與「股票市值」欄位，結構調整為 Date, Holdings, Count, 現金, 股票市值, 總資產, 補充說明。
4. **新增 Daily 工作表**：每日持股快照，包含股數、收盤價與市值。
5. **程式註解與命名**：全繁體中文註解，檔名結尾統一使用 `equity2025新-動態版V1`。

## 績效表現 (2019-2025)
- **年化報酬率 (CAGR)**: {cagr:.2%}
- **最大回撤 (MaxDD)**: {mdd:.2%}
- **Calmar Ratio**: {calmar:.2f}

## 策略核心規則
- **動態分配**: 再平衡賣出所得立即投入新標的 (上限 1000 萬)。
- **單位限制**: 交易單位必須為 1000 股之整數倍。
- **指標**: SMA{SMA_PERIOD} 結合 ROC{ROC_PERIOD}。
"""
with open(OUTPUT_MD, 'w', encoding='utf-8') as f:
    f.write(md_content)

# ==========================================
# 5. 產出 Jupyter Notebook (Generate Notebook)
# ==========================================
nb = nbf.v4.new_notebook()
with open('run_backtest_equity2025新_動態版V1.py', 'r', encoding='utf-8') as f:
    backtest_code = f.read()
    if 'if __name__ == "__main__":' in backtest_code:
        backtest_code = backtest_code[:backtest_code.find('if __name__ == "__main__":')]

nb.cells.append(nbf.v4.new_markdown_cell("# Asset Class Trend Following 策略回測 (equity2025新-動態版V1)"))
nb.cells.append(nbf.v4.new_code_cell(backtest_code))
nb.cells.append(nbf.v4.new_code_cell(
    f"prices, code_to_name = clean_data('{DATA_FILE}')\n"
    "bt = Backtester(prices, code_to_name, 30000000)\n"
    f"eq_df, trades, hold, trades2, daily = bt.run({SMA_PERIOD}, {ROC_PERIOD}, {STOP_LOSS_PCT}, {REBALANCE})\n"
    "import matplotlib.pyplot as plt\n"
    "plt.figure(figsize=(12, 6))\n"
    "plt.plot(eq_df['日期'], eq_df['權益'])\n"
    "plt.title('Dynamic Allocation V1 Equity Curve')\n"
    "plt.grid(True)\n"
    "plt.show()"
))

with open('trendstrategy_equity2025新-動態版V1.ipynb', 'w', encoding='utf-8') as f:
    nbf.write(nb, f)

print("所有交付成果檔案 (Excel, MD, IPYNB) 已成功產出。")
