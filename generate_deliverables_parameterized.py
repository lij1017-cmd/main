import pandas as pd
import numpy as np
import pickle
from backtest import Backtester, calculate_metrics
from backtest import calculate_win_rate
import nbformat as nbf
import argparse
import os

def main(sma, roc, sl, suffix):
    # 1. Load data
    with open('cleaned_data.pkl', 'rb') as f:
        prices, code_to_name = pickle.load(f)

    print(f"Using parameters: SMA={sma}, ROC={roc}, SL={sl:.3f}")

    # 2. Run final backtest
    bt = Backtester(prices, code_to_name)
    eq, trades, hold = bt.run(sma, roc, sl)
    cagr, mdd, calmar, ret = calculate_metrics(eq)
    win_rate = calculate_win_rate(trades)

    print(f"Final Backtest: CAGR={cagr:.2%}, MaxDD={mdd:.2%}, Calmar={calmar:.2f}, WinRate={win_rate:.2%}")

    # 3. Generate Plateau Table (SMA variation)
    plateau_data = []
    for s in [sma-4, sma-2, sma, sma+2, sma+4]:
        if s < 5: continue
        eq_s, _, _ = bt.run(s, roc, sl)
        c_s, m_s, cl_s, _ = calculate_metrics(eq_s)
        plateau_data.append({'SMA': s, 'CAGR': f"{c_s:.2%}", 'MaxDD': f"{m_s:.2%}", 'Calmar': f"{cl_s:.2f}"})
    plateau_df = pd.DataFrame(plateau_data)

    # 4. Save Excel
    OUTPUT_EXCEL = f'trendstrategy_results_equity2025成{suffix}.xlsx'
    with pd.ExcelWriter(OUTPUT_EXCEL, engine='xlsxwriter') as writer:
        # Summary
        summary_df = pd.DataFrame([
            {'項目': '年化報酬率 (CAGR)', '數值': f"{cagr:.2%}"},
            {'項目': '最大回撤 (MaxDD)', '數值': f"{mdd:.2%}"},
            {'項目': 'Calmar Ratio', '數值': f"{calmar:.2f}"},
            {'項目': '勝率 (Win Rate)', '數值': f"{win_rate:.2%}"},
            {'項目': '總報酬率', '數值': f"{ret:.2%}"},
            {'項目': '最佳 SMA', '數值': sma},
            {'項目': '最佳 ROC', '數值': roc},
            {'項目': '最佳 停損%', '數值': f"{sl*100:.1f}%"}
        ])
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

        # Equity_Curve
        eq_curve_df = eq.reset_index()
        eq_curve_df.columns = ['日期', '權益']
        eq_curve_df.to_excel(writer, sheet_name='Equity_Curve', index=False)

        # Equity_Hold
        hold.to_excel(writer, sheet_name='Equity_Hold', index=False)

        # Trades
        trades['最佳參數'] = f"SMA={sma}, ROC={roc}, SL={sl}"
        trades = trades[['日期', '股票代號', '狀態', '價格', '股數', '動能值', '標的名稱', '最佳參數', '原因', '說明']]
        trades.to_excel(writer, sheet_name='Trades', index=False)

    print(f"Excel saved to {OUTPUT_EXCEL}")

    # 5. Generate Markdown Report
    OUTPUT_MD = f'reproduce_report2025成{suffix}.md'
    md_content = f"""# Asset Class Trend Following 策略重現報告 (2025成{suffix})

## 策略摘要
本策略採用資產類別趨勢追隨 (Asset Class Trend Following) 邏輯，並加入最高價回落停損機制。透過指定的參數進行回測。

## 核心參數
- **SMA 週期**: {sma}
- **ROC 週期**: {roc}
- **停損比例 (StopLoss%)**: {sl*100:.1f}%

## 績效表現
- **年化報酬率 (CAGR)**: {cagr:.2%}
- **最大回撤 (MaxDD)**: {mdd:.2%}
- **Calmar Ratio**: {calmar:.2f}
- **勝率 (Win Rate)**: {win_rate:.2%}

## 參數高原表 (SMA 敏感度分析)
{plateau_df.to_markdown(index=False)}

## 策略邏輯說明
1. **資料處理**: 對「個股1.xlsx」進行清洗，缺失值以前期值填補，較早上市者以前期首價填補。
2. **進場條件**: 價格高於 SMA 且 ROC > 0。
3. **投資組合**: 每 5 日再平衡，選取 ROC 前 3 名，等權重分配資金。
4. **交易執行**: T 日產生訊號，T+1 日收盤價執行。包含交易成本 (買進 0.1425%, 賣出 0.1425% + 0.3% 稅)。
5. **停損機制**: 採用最高價回落停損 (Peak-to-Trough Stop)，當價格低於持有期間最高價之 {sl*100:.1f}% 時觸發。

"""
    with open(OUTPUT_MD, 'w', encoding='utf-8') as f:
        f.write(md_content)
    print(f"Markdown report saved to {OUTPUT_MD}")

    # 6. Generate Notebook
    nb = nbf.v4.new_notebook()

    code_cells = [
        ("# --- 全域參數設定區塊 ---\n"
         f"SMA_PERIOD = {sma}\n"
         f"ROC_PERIOD = {roc}\n"
         f"STOP_LOSS_PCT = {sl}\n"
         "INITIAL_CAPITAL = 30000000\n"
         "DATA_FILE = '個股1.xlsx'\n"),

        ("import pandas as pd\n"
         "import numpy as np\n"
         "import matplotlib.pyplot as plt\n"
         "import os"),

        ("def clean_data(filepath):\n"
         "    df_raw = pd.read_excel(filepath, header=None)\n"
         "    stock_codes = df_raw.iloc[0, 2:].values\n"
         "    stock_names = df_raw.iloc[1, 2:].values\n"
         "    dates = pd.to_datetime(df_raw.iloc[2:, 1])\n"
         "    prices = df_raw.iloc[2:, 2:].astype(float)\n"
         "    prices.index = dates\n"
         "    prices.columns = stock_codes\n"
         "    code_to_name = dict(zip(stock_codes, stock_names))\n"
         "    prices = prices.ffill().bfill()\n"
         "    return prices, code_to_name\n"),
    ]

    with open('backtest.py', 'r') as f:
        lines = f.readlines()
        try:
            main_idx = next(i for i, line in enumerate(lines) if 'if __name__ == "__main__":' in line)
            backtest_code = "".join(lines[:main_idx])
        except StopIteration:
            backtest_code = "".join(lines)

    nb.cells.append(nbf.v4.new_markdown_cell(f"# Asset Class Trend Following 策略回測 (2025成{suffix})"))
    nb.cells.append(nbf.v4.new_code_cell(code_cells[0]))
    nb.cells.append(nbf.v4.new_code_cell(code_cells[1]))
    nb.cells.append(nbf.v4.new_code_cell(code_cells[2]))
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

    OUTPUT_NB = f'trendstrategy_equity2025成{suffix}.ipynb'
    with open(OUTPUT_NB, 'w', encoding='utf-8') as f:
        nbf.write(nb, f)
    print(f"Notebook saved to {OUTPUT_NB}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--sma", type=int, default=33)
    parser.add_argument("--roc", type=int, default=58)
    parser.add_argument("--sl", type=float, default=0.095)
    parser.add_argument("--suffix", type=str, default="")
    args = parser.parse_args()
    main(args.sma, args.roc, args.sl, args.suffix)
