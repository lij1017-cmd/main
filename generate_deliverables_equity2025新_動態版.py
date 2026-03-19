import pandas as pd
import numpy as np
import pickle
import nbformat as nbf
from run_backtest_equity2025新_動態版 import Backtester, clean_data, calculate_metrics

# ==========================================
# 1. 參數設定區塊 (Global Parameters)
# ==========================================
DATA_FILE = '個股合-1.xlsx'
SMA_BEST = 87
ROC_BEST = 54
SL_BEST = 0.09
REBALANCE = 6
INITIAL_CAPITAL = 30000000

# ==========================================
# 2. 執行回測 (Run Backtest)
# ==========================================
prices, code_to_name = clean_data(DATA_FILE)
bt = Backtester(prices, code_to_name, INITIAL_CAPITAL)
eq, trades, hold = bt.run(SMA_BEST, ROC_BEST, SL_BEST, REBALANCE)
cagr, mdd, calmar, ret = calculate_metrics(eq)

# ==========================================
# 3. 產出 Excel 報表 (Generate Excel Report)
# ==========================================
OUTPUT_EXCEL = 'trendstrategy_results_equity2025新-動態版.xlsx'
with pd.ExcelWriter(OUTPUT_EXCEL, engine='xlsxwriter') as writer:
    # Summary
    pd.DataFrame([
        {'項目': '年化報酬率 (CAGR)', '數值': f"{cagr:.2%}"},
        {'項目': '最大回撤 (MaxDD)', '數值': f"{mdd:.2%}"},
        {'項目': 'Calmar Ratio', '數值': f"{calmar:.2f}"},
        {'項目': '總報酬率', '數值': f"{ret:.2%}"},
        {'項目': '分配模式', '數值': '動態部位預算 (Proceeds Reinvestment)'},
        {'項目': '上限預算', '數值': '10,000,000 TWD/部位'}
    ]).to_excel(writer, sheet_name='Summary', index=False)

    # Equity Curve
    eq.reset_index().rename(columns={'index': '日期', 0: '權益'}).to_excel(writer, sheet_name='Equity_Curve', index=False)

    # Equity Hold
    hold.to_excel(writer, sheet_name='Equity_Hold', index=False)

    # Trades
    trades['分配模式'] = "動態版"
    cols = ['訊號日期', '股票代號', '狀態', '價格', '股數', '動能值', '標的名稱', '分配模式', '原因', '買入手續費', '賣出手續費', '賣出交易稅', '說明']
    trades[cols].to_excel(writer, sheet_name='Trades', index=False)

# ==========================================
# 4. 產出 Markdown 報告 (Generate MD Report)
# ==========================================
OUTPUT_MD = 'reproduce_report_equity2025新-動態版.md'
md_content = f"""# Asset Class Trend Following 策略回測報告 (equity2025新-動態版)

## 策略摘要
本策略使用「動態分配邏輯」：再平衡日若有標的更換，賣出所得資金（上限 1000 萬）將立即用於建立新部位。
此機制確保了資金在市場波動導致帳戶價值變動時，仍能維持穩定的部位配置。

## 動態分配規則
1. **槽位機制**: 投資組合分為 3 個邏輯槽位，每槽位上限 1000 萬。
2. **所得再投**: 再平衡賣出股票後，其所得金額 (Proceeds) 優先作為該槽位新訊號的買入預算。
3. **超額歸池**: 若賣出所得超過 1000 萬，多餘部分與買入餘額回到中央資金池。
4. **填補空缺**: 因停損空出的槽位，於再平衡日從資金池撥款買入新訊號。

## 績效表現 (2019-2025)
- **年化報酬率 (CAGR)**: {cagr:.2%}
- **最大回撤 (MaxDD)**: {mdd:.2%}
- **Calmar Ratio**: {calmar:.2f}

## 核心規則
- **選股**: 價格 > SMA 且 ROC > 0，取前 3 名。
- **單位**: 交易單位必須為 1000 股之整數倍。
- **停損**: 9.0% 追蹤停損。
"""
with open(OUTPUT_MD, 'w', encoding='utf-8') as f:
    f.write(md_content)

# ==========================================
# 5. 產出 Jupyter Notebook (Generate Notebook)
# ==========================================
nb = nbf.v4.new_notebook()
with open('run_backtest_equity2025新_動態版.py', 'r', encoding='utf-8') as f:
    backtest_code = f.read()
    if 'if __name__ == "__main__":' in backtest_code:
        backtest_code = backtest_code[:backtest_code.find('if __name__ == "__main__":')]

nb.cells.append(nbf.v4.new_markdown_cell("# Asset Class Trend Following 策略回測 (equity2025新-動態版)"))
nb.cells.append(nbf.v4.new_code_cell(backtest_code))
nb.cells.append(nbf.v4.new_code_cell(
    "prices, code_to_name = clean_data('個股合-1.xlsx')\n"
    "bt = Backtester(prices, code_to_name, 30000000)\n"
    "eq, trades, hold = bt.run(87, 54, 0.09, 6)\n"
    "import matplotlib.pyplot as plt\n"
    "plt.figure(figsize=(12, 6))\n"
    "plt.plot(eq)\n"
    "plt.title('Dynamic Allocation Equity Curve')\n"
    "plt.grid(True)\n"
    "plt.show()"
))

with open('trendstrategy_equity2025新-動態版.ipynb', 'w', encoding='utf-8') as f:
    nbf.write(nb, f)

print("所有交付成果檔案已成功產出。")
