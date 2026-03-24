import pandas as pd
import numpy as np
import nbformat as nbf
import xlsxwriter
from backtest_engine import Backtester, clean_data, calculate_metrics

def main():
    DATA_FILE = '個股合-1.xlsx'
    # 參數設定
    SMA_PERIOD = 54
    ROC_PERIOD = 52
    STOP_LOSS_PCT = 0.09
    REBALANCE = 9
    INITIAL_CAPITAL = 30000000

    prices, code_to_name = clean_data(DATA_FILE)
    bt = Backtester(prices, code_to_name, INITIAL_CAPITAL)

    start_date = '2019-01-02'
    end_date = '2024-12-31'

    print(f"正在執行指定期間交付成果產出 (2019-2024) (SMA={SMA_PERIOD}, ROC={ROC_PERIOD}, SL={STOP_LOSS_PCT*100:.1f}%, Reb={REBALANCE})...")
    eq_df, trades, hold, trades2, daily = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE, start_date, end_date)
    cagr, mdd, calmar = calculate_metrics(eq_df)

    # 1. 產出 Excel
    OUTPUT_EXCEL = 'trendstrategy_results_equity-new.xlsx'
    with pd.ExcelWriter(OUTPUT_EXCEL, engine='xlsxwriter') as writer:
        # Summary
        pd.DataFrame([
            {'項目': '年化報酬率 (CAGR)', '數值': f"{cagr:.2%}"},
            {'項目': '最大回撤 (MaxDD)', '數值': f"{mdd:.2%}"},
            {'項目': 'Calmar Ratio', '數值': f"{calmar:.2f}"},
            {'項目': '初始資金', '數值': f"{INITIAL_CAPITAL:,}"},
            {'項目': '期末淨值', '數值': f"{eq_df['權益'].iloc[-1]:,.0f}"},
            {'項目': '版本', '數值': 'Equity-New (SMA54/ROC52/SL9/Reb9)'}
        ]).to_excel(writer, sheet_name='Summary', index=False)

        eq_df.to_excel(writer, sheet_name='Equity_Curve', index=False)
        hold.to_excel(writer, sheet_name='Equity_Hold', index=False)
        trades.to_excel(writer, sheet_name='Trades', index=False)
        trades2.to_excel(writer, sheet_name='Trades2', index=False)
        daily.to_excel(writer, sheet_name='Daily', index=False)

        # 插入 Equity Curve 圖表
        workbook = writer.book
        curves_sheet = writer.sheets['Equity_Curve']
        chart = workbook.add_chart({'type': 'line'})
        max_row = len(eq_df)
        chart.add_series({
            'name': 'Equity Curve',
            'categories': ['Equity_Curve', 1, 0, max_row, 0],
            'values': ['Equity_Curve', 1, 1, max_row, 1],
        })
        chart.set_title({'name': 'Full Period Equity Curve (SMA54/ROC52)'})
        curves_sheet.insert_chart('D2', chart)

    # 2. 產出 Markdown
    OUTPUT_MD = 'reproduce_report_equity-new.md'
    md_content = f"""# Asset Class Trend Following 策略回測報告 (Equity-New)

## 策略參數
- **SMA 週期**: {SMA_PERIOD}
- **ROC 週期**: {ROC_PERIOD}
- **追蹤停損**: {STOP_LOSS_PCT*100:.1f}%
- **再平衡頻率**: {REBALANCE} 交易日

## 回測績效 (2019-2024)
- **年化報酬率 (CAGR)**: {cagr:.2%}
- **最大回撤 (MaxDD)**: {mdd:.2%}
- **卡瑪比率 (Calmar Ratio)**: {calmar:.2f}
- **期末淨值**: ${eq_df['權益'].iloc[-1]:,.0f}

## 核心邏輯 (Dynamic Allocation V1)
- **資金上限**: 每次再平衡固定投入 3000 萬，每檔平均上限 1000 萬。
- **賣出再投入**: 再平衡日賣出後，資金立即投入新標的 (按 1000 股整張)。
- **餘額管理**: 賣出金額超過 1000 萬部分保留在現金池。
- **停損出場**: 觸發停損隔日出場，資金回流現金池，下次再平衡重新分配。
"""
    with open(OUTPUT_MD, 'w', encoding='utf-8') as f:
        f.write(md_content)

    # 3. 產出 Jupyter Notebook
    nb = nbf.v4.new_notebook()
    nb.cells.append(nbf.v4.new_markdown_cell("# Asset Class Trend Following 策略回測 (Equity-New)"))
    with open('backtest_engine.py', 'r', encoding='utf-8') as f:
        engine_code = f.read()
    nb.cells.append(nbf.v4.new_code_cell(engine_code))
    nb.cells.append(nbf.v4.new_code_cell(
        f"prices, code_to_name = clean_data('{DATA_FILE}')\n"
        f"bt = Backtester(prices, code_to_name, 30000000)\n"
        f"eq_df, trades, hold, trades2, daily = bt.run({SMA_PERIOD}, {ROC_PERIOD}, {STOP_LOSS_PCT}, {REBALANCE})\n"
        "from backtest_engine import calculate_metrics\n"
        "cagr, mdd, calmar = calculate_metrics(eq_df)\n"
        "print(f'CAGR: {cagr:.2%}, MaxDD: {mdd:.2%}, Calmar: {calmar:.2f}')"
    ))
    with open('trendstrategy_equity-new.ipynb', 'w', encoding='utf-8') as f:
        nbf.write(nb, f)

    print(f"所有交付成果檔案 (Excel, MD, IPYNB) 已成功產出。")
    print(f"CAGR: {cagr:.2%}, MDD: {mdd:.2%}, Calmar: {calmar:.2f}")

if __name__ == "__main__":
    main()
