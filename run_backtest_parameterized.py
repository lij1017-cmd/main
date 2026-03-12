import pandas as pd
import numpy as np
import pickle
import os
import nbformat as nbf
from backtest import Backtester, calculate_metrics, calculate_win_rate

def clean_data(filepath):
    df_raw = pd.read_excel(filepath, header=None)
    stock_codes = df_raw.iloc[0, 2:].values
    stock_names = df_raw.iloc[1, 2:].values
    dates = pd.to_datetime(df_raw.iloc[2:, 1])
    prices = df_raw.iloc[2:, 2:].astype(float)
    prices.index = dates
    prices.columns = stock_codes
    code_to_name = dict(zip(stock_codes, stock_names))
    prices = prices.ffill().bfill()
    prices = prices.dropna(axis=1, how='all')
    return prices, code_to_name

# Parameters
SMA_TARGET = 35
ROC_TARGET = 49
SL_TARGET = 0.09
DATA_FILE = '個股1.xlsx'
INITIAL_CAPITAL = 30000000

# 1. Load and clean data
prices, code_to_name = clean_data(DATA_FILE)

# 2. Run backtest with target params
bt = Backtester(prices, code_to_name, INITIAL_CAPITAL)
eq, trades, hold = bt.run(SMA_TARGET, ROC_TARGET, SL_TARGET)
cagr, mdd, calmar, ret = calculate_metrics(eq)
win_rate = calculate_win_rate(trades)

print(f"Backtest Results: CAGR={cagr:.2%}, MaxDD={mdd:.2%}, Calmar={calmar:.2f}, WinRate={win_rate:.2%}")

# 3. Generate Plateau Table (SMA variation)
plateau_data = []
for s in [SMA_TARGET-4, SMA_TARGET-2, SMA_TARGET, SMA_TARGET+2, SMA_TARGET+4]:
    eq_s, _, _ = bt.run(s, ROC_TARGET, SL_TARGET)
    c_s, m_s, cl_s, _ = calculate_metrics(eq_s)
    plateau_data.append({'SMA': s, 'CAGR': f"{c_s:.2%}", 'MaxDD': f"{m_s:.2%}", 'Calmar': f"{cl_s:.2f}"})
plateau_df = pd.DataFrame(plateau_data)

# 4. Save Excel
OUTPUT_EXCEL = 'trendstrategy_results_equity2025成-1.xlsx'
with pd.ExcelWriter(OUTPUT_EXCEL, engine='xlsxwriter') as writer:
    summary_df = pd.DataFrame([
        {'項目': '年化報酬率 (CAGR)', '數值': f"{cagr:.2%}"},
        {'項目': '最大回撤 (MaxDD)', '數值': f"{mdd:.2%}"},
        {'項目': 'Calmar Ratio', '數值': f"{calmar:.2f}"},
        {'項目': '勝率 (Win Rate)', '數值': f"{win_rate:.2%}"},
        {'項目': '總報酬率', '數值': f"{ret:.2%}"},
        {'項目': '最佳 SMA', '數值': SMA_TARGET},
        {'項目': '最佳 ROC', '數值': ROC_TARGET},
        {'項目': '最佳 停損%', '數值': f"{SL_TARGET*100:.1f}%"}
    ])
    summary_df.to_excel(writer, sheet_name='Summary', index=False)
    eq.reset_index().rename(columns={'index': '日期', 0: '權益'}).to_excel(writer, sheet_name='Equity_Curve', index=False)
    hold.to_excel(writer, sheet_name='Equity_Hold', index=False)
    trades['最佳參數'] = f"SMA={SMA_TARGET}, ROC={ROC_TARGET}, SL={SL_TARGET}"
    # Ensure '報酬率' is in the trades log for the Excel output if it exists
    cols = ['日期', '股票代號', '狀態', '價格', '股數', '動能值', '標的名稱', '最佳參數', '原因', '說明']
    if '報酬率' in trades.columns:
        cols.append('報酬率')
    trades = trades[cols]
    trades.to_excel(writer, sheet_name='Trades', index=False)

# 5. Generate Markdown Report
OUTPUT_MD = 'reproduce_report2025成-1.md'
md_content = f"""# Asset Class Trend Following 策略重現報告 (2025成-1)

## 策略摘要
本策略採用資產類別趨勢追隨 (Asset Class Trend Following) 邏輯，並加入最高價回落停損機制。

## 核心參數
- **SMA 週期**: {SMA_TARGET}
- **ROC 週期**: {ROC_TARGET}
- **停損比例 (StopLoss%)**: {SL_TARGET*100:.1f}%

## 績效表現
- **年化報酬率 (CAGR)**: {cagr:.2%}
- **最大回撤 (MaxDD)**: {mdd:.2%}
- **Calmar Ratio**: {calmar:.2f}
- **勝率 (Win Rate)**: {win_rate:.2%}

## 參數高原表 (SMA 敏感度分析)
{plateau_df.to_markdown(index=False)}

## 策略邏輯說明
1. **資料處理**: 對「{DATA_FILE}」進行清洗，缺失值以前期值填補，較早上市者以前期首價填補。
2. **進場條件**: 價格高於 SMA 且 ROC > 0。
3. **投資組合**: 每 5 日再平衡，選取 ROC 前 3 名，等權重分配資金。
4. **交易執行**: T 日產生訊號，T+1 日收盤價執行。包含交易成本 (買進 0.1425%, 賣出 0.1425% + 0.3% 稅)。
5. **停損機制**: 採用最高價回落停損 (Peak-to-Trough Stop)，當價格低於持有期間最高價之 {SL_TARGET*100:.1f}% 時觸發。
"""
with open(OUTPUT_MD, 'w', encoding='utf-8') as f:
    f.write(md_content)

# 6. Generate Notebook
with open('backtest.py', 'r') as f:
    lines = f.readlines()
    try:
        main_idx = next(i for i, line in enumerate(lines) if 'if __name__ == "__main__":' in line)
        backtest_code = "".join(lines[:main_idx])
    except StopIteration:
        backtest_code = "".join(lines)

nb = nbf.v4.new_notebook()
nb.cells.append(nbf.v4.new_markdown_cell("# Asset Class Trend Following 策略回測 (2025成-1)"))
nb.cells.append(nbf.v4.new_code_cell(
    f"# --- 全域參數設定區塊 ---\n"
    f"SMA_PERIOD = {SMA_TARGET}\n"
    f"ROC_PERIOD = {ROC_TARGET}\n"
    f"STOP_LOSS_PCT = {SL_TARGET}\n"
    f"INITIAL_CAPITAL = {INITIAL_CAPITAL}\n"
    f"DATA_FILE = '{DATA_FILE}'\n"
))
nb.cells.append(nbf.v4.new_code_cell("import pandas as pd\nimport numpy as np\nimport matplotlib.pyplot as plt\nimport os"))
nb.cells.append(nbf.v4.new_code_cell(
    "def clean_data(filepath):\n"
    "    df_raw = pd.read_excel(filepath, header=None)\n"
    "    stock_codes = df_raw.iloc[0, 2:].values\n"
    "    stock_names = df_raw.iloc[1, 2:].values\n"
    "    dates = pd.to_datetime(df_raw.iloc[2:, 1])\n"
    "    prices = df_raw.iloc[2:, 2:].astype(float)\n"
    "    prices.index = dates\n"
    "    prices.columns = stock_codes\n"
    "    code_to_name = dict(zip(stock_codes, stock_names))\n"
    "    prices = prices.ffill().bfill()\n"
    "    prices = prices.dropna(axis=1, how='all')\n"
    "    return prices, code_to_name\n"
))
nb.cells.append(nbf.v4.new_code_cell(backtest_code))
nb.cells.append(nbf.v4.new_code_cell(
    "prices, code_to_name = clean_data(DATA_FILE)\n"
    "bt = Backtester(prices, code_to_name, INITIAL_CAPITAL)\n"
    "eq, trades, hold = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT)\n"
    "\n"
    "plt.figure(figsize=(12, 6))\n"
    "plt.plot(eq)\n"
    "plt.title('Equity Curve')\n"
    "plt.grid(True)\n"
    "plt.show()\n"
))

OUTPUT_NB = 'trendstrategy_equity2025成-1.ipynb'
with open(OUTPUT_NB, 'w', encoding='utf-8') as f:
    nbf.write(nb, f)

print("Deliverables generated successfully.")
