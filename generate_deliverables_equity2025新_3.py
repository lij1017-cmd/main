import pandas as pd
import numpy as np
import pickle
import nbformat as nbf
from run_backtest_equity2025新_3 import Backtester, clean_data, calculate_metrics, calculate_win_rate

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
win_rate = calculate_win_rate(trades)

# ==========================================
# 3. 參數敏感度分析 (Plateau/Sensitivity Analysis)
# ==========================================
plateau_data = []
for s in [SMA_BEST-4, SMA_BEST-2, SMA_BEST, SMA_BEST+2, SMA_BEST+4]:
    if s < 5: continue
    eq_s, _, _ = bt.run(s, ROC_BEST, SL_BEST, REBALANCE)
    c_s, m_s, cl_s, _ = calculate_metrics(eq_s)
    plateau_data.append({'SMA': s, 'CAGR': f"{c_s:.2%}", 'MaxDD': f"{m_s:.2%}", 'Calmar': f"{cl_s:.2f}"})
plateau_df = pd.DataFrame(plateau_data)

# ==========================================
# 4. 產出 Excel 報表 (Generate Excel Report)
# ==========================================
OUTPUT_EXCEL = 'trendstrategy_results_equity2025新-3.xlsx'
with pd.ExcelWriter(OUTPUT_EXCEL, engine='xlsxwriter') as writer:
    summary_df = pd.DataFrame([
        {'項目': '年化報酬率 (CAGR)', '數值': f"{cagr:.2%}"},
        {'項目': '最大回撤 (MaxDD)', '數值': f"{mdd:.2%}"},
        {'項目': 'Calmar Ratio', '數值': f"{calmar:.2f}"},
        {'項目': '總報酬率', '數值': f"{ret:.2%}"},
        {'項目': '最佳 SMA', '數值': SMA_BEST},
        {'項目': '最佳 ROC', '數值': ROC_BEST},
        {'項目': '最佳 停損%', '數值': f"{SL_BEST*100:.1f}%"},
        {'項目': '再平衡週期', '數值': f"{REBALANCE}日"}
    ])
    summary_df.to_excel(writer, sheet_name='Summary', index=False)
    eq.reset_index().rename(columns={'index': '日期', 0: '權益'}).to_excel(writer, sheet_name='Equity_Curve', index=False)
    hold.to_excel(writer, sheet_name='Equity_Hold', index=False)

    trades['最佳參數'] = f"SMA={SMA_BEST}, ROC={ROC_BEST}, SL={SL_BEST}"
    cols = ['訊號日期', '股票代號', '狀態', '價格', '股數', '動能值', '標的名稱', '最佳參數', '原因', '買入手續費', '賣出手續費', '賣出交易稅', '說明']
    trades = trades[cols]
    trades.to_excel(writer, sheet_name='Trades', index=False)

# ==========================================
# 5. 產出 Markdown 報告 (Generate MD Report)
# ==========================================
OUTPUT_MD = 'reproduce_report_equity2025新-3.md'
md_content = f"""# Asset Class Trend Following 策略回測報告 (equity2025新-3)

## 策略摘要
本策略使用「個股合-1.xlsx」進行 2019-2025 全週期回測。採用 6 日再平衡機制，結合 SMA 與 ROC 進行選股，並落實 9.0% 追蹤停損。
本次更新 (v3)：
1. **精準補充說明**：
   - 停損顯示具體股票名稱。
   - 再平衡不足時說明排除原因 (如 ROC < 0 或 價格 < SMA)。
   - 紀錄因資金或整張交易限制導致的不足。
2. **維持先前優化**：所有交易以 1000 股為單位，訊號日期明確標註。

## 核心參數
- **SMA 週期**: {SMA_BEST}
- **ROC 週期**: {ROC_BEST}
- **停損比例 (StopLoss%)**: {SL_BEST*100:.1f}%
- **再平衡週期**: {REBALANCE} 個交易日

## 績效表現 (2019-2025)
- **年化報酬率 (CAGR)**: {cagr:.2%}
- **最大回撤 (MaxDD)**: {mdd:.2%}
- **Calmar Ratio**: {calmar:.2f}

## 參數高原表 (SMA 敏感度)
{plateau_df.to_markdown(index=False)}

## 策略規則
1. **選股**: 價格 > SMA 且 ROC > 0，取前 3 名。
2. **再平衡**: 每 {REBALANCE} 日執行一次。
3. **執行**: T 產生訊號，T+1 收盤價成交。
4. **成本**: 買入手續費 0.1425%，賣出手續費 0.1425%，賣出交易稅 0.3%。
5. **停損**: 持有期間最高價回落 {SL_BEST*100:.1f}% 即於次日收盤出清。
6. **單位限制**: 交易單位必須為 1000 股之整數倍。
"""
with open(OUTPUT_MD, 'w', encoding='utf-8') as f:
    f.write(md_content)

# ==========================================
# 6. 產出 Jupyter Notebook (Generate Notebook)
# ==========================================
nb = nbf.v4.new_notebook()
with open('run_backtest_equity2025新_3.py', 'r', encoding='utf-8') as f:
    backtest_code = f.read()
    if 'if __name__ == "__main__":' in backtest_code:
        backtest_code = backtest_code[:backtest_code.find('if __name__ == "__main__":')]

nb.cells.append(nbf.v4.new_markdown_cell("# Asset Class Trend Following 策略回測 (equity2025新-3)"))
nb.cells.append(nbf.v4.new_markdown_cell("## 1. 全域參數設定區塊\n在此設定回測所需的核心參數。"))
nb.cells.append(nbf.v4.new_code_cell(
    f"# 全域參數設定區塊\n"
    f"SMA_PERIOD = {SMA_BEST}\n"
    f"ROC_PERIOD = {ROC_BEST}\n"
    f"STOP_LOSS_PCT = {SL_BEST}\n"
    f"REBALANCE_INTERVAL = {REBALANCE}\n"
    f"INITIAL_CAPITAL = {INITIAL_CAPITAL}\n"
    f"DATA_FILE = '{DATA_FILE}'"
))
nb.cells.append(nbf.v4.new_markdown_cell("## 2. 策略核心代碼\n定義回測引擎、資料清洗與指標計算邏輯。"))
nb.cells.append(nbf.v4.new_code_cell(backtest_code))
nb.cells.append(nbf.v4.new_markdown_cell("## 3. 執行回測與結果視覺化\n啟動引擎進行回測並繪製權益曲線。"))
nb.cells.append(nbf.v4.new_code_cell(
    "import matplotlib.pyplot as plt\n"
    "prices, code_to_name = clean_data(DATA_FILE)\n"
    "bt = Backtester(prices, code_to_name, INITIAL_CAPITAL)\n"
    "eq, trades, hold = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE_INTERVAL)\n"
    "\n"
    "plt.figure(figsize=(12, 6))\n"
    "plt.plot(eq)\n"
    "plt.title('Equity Curve (2019-2025)')\n"
    "plt.xlabel('Date')\n"
    "plt.ylabel('Total Equity')\n"
    "plt.grid(True)\n"
    "plt.show()"
))

with open('trendstrategy_equity2025新-3.ipynb', 'w', encoding='utf-8') as f:
    nbf.write(nb, f)

print("所有交付成果檔案 (Excel, MD, IPYNB) 已成功產出。")
