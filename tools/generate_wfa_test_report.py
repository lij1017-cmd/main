import pandas as pd
import numpy as np
import xlsxwriter
from run_wfa import Backtester, clean_data, calculate_metrics

def main():
    # 最佳化結果
    SMA_PERIOD = 54
    ROC_PERIOD = 52
    STOP_LOSS_PCT = 0.09
    REBALANCE = 9
    INITIAL_CAPITAL = 30000000
    DATA_FILE = '個股合-1.xlsx'

    prices, code_to_name = clean_data(DATA_FILE)
    bt = Backtester(prices, code_to_name, INITIAL_CAPITAL)

    # 目標區間
    start_date = '2020-01-02'
    end_date = '2024-05-31'

    print(f"正在產出驗證報表: SMA={SMA_PERIOD}, ROC={ROC_PERIOD}, SL={STOP_LOSS_PCT*100:.1f}%, Reb={REBALANCE}")

    eq_df, t_log, h_log, t2_log, d_log = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE, start_date, end_date)
    cagr, mdd, calmar = calculate_metrics(eq_df)
    trades = len(t_log)

    # 產出 Excel
    OUTPUT_FILE = 'walk-forward-test.xlsx'
    with pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter') as writer:
        # Summary
        pd.DataFrame([{
            '回測期間': f"{start_date} - {end_date}",
            '年化報酬 (CAGR)': cagr,
            '最大回撤 (MaxDD)': mdd,
            '卡瑪比率 (Calmar)': calmar,
            '交易次數': trades,
            'SMA': SMA_PERIOD,
            'ROC': ROC_PERIOD,
            'Stop Loss': f"{STOP_LOSS_PCT*100:.1f}%",
            'Rebalance': REBALANCE
        }]).to_excel(writer, sheet_name='Summary', index=False)

        workbook = writer.book
        summary_sheet = writer.sheets['Summary']
        pct_fmt = workbook.add_format({'num_format': '0.00%'})
        num_fmt = workbook.add_format({'num_format': '0.00'})
        summary_sheet.set_column('B:C', 15, pct_fmt)
        summary_sheet.set_column('D:D', 15, num_fmt)

        # Equity_Curve
        eq_df.to_excel(writer, sheet_name='Equity_Curve', index=False)
        curves_sheet = writer.sheets['Equity_Curve']

        chart = workbook.add_chart({'type': 'line'})
        max_row = len(eq_df)
        chart.add_series({
            'name': 'Equity Curve',
            'categories': ['Equity_Curve', 1, 0, max_row, 0],
            'values': ['Equity_Curve', 1, 1, max_row, 1],
        })
        chart.set_title({'name': f'Equity Curve (Test: {start_date}-{end_date})'})
        curves_sheet.insert_chart('D2', chart)

    print(f"已成功產出測試報表: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
