import pandas as pd
import numpy as np
import xlsxwriter
from backtest_vol import BacktesterVol, calculate_metrics_dual, clean_data
import os

def main():
    DATA_FILE = '樣本集-1.xlsx'
    OUTPUT_FILE = 'equityV-adj1.xlsx'
    TRADING_CAP = 30000000
    AUTH_CAP = 150000000

    # 參數設定
    SMA_PERIOD = 303
    ROC_PERIOD = 14
    REBALANCE = 9
    VOL_PERIOD = 15
    VOL_MULTIPLIER = 2.7
    BREADTH_WINDOW = 290
    BREADTH_THRESHOLD = 0.42
    MKT_SMA = 14

    print("讀取資料中...")
    prices, volumes, code_to_name = clean_data(DATA_FILE)
    bt = BacktesterVol(prices, volumes, code_to_name, trading_capital=TRADING_CAP, authorized_capital=AUTH_CAP)

    print("執行回測中...")
    eq_df, trades_df, trades2_df, daily_details_df = bt.run(
        sma_period=SMA_PERIOD,
        roc_period=ROC_PERIOD,
        vol_period=VOL_PERIOD,
        vol_multiplier=VOL_MULTIPLIER,
        rebalance_interval=REBALANCE,
        breadth_threshold=BREADTH_THRESHOLD,
        mkt_sma_window=MKT_SMA,
        breadth_window=BREADTH_WINDOW
    )

    metrics = calculate_metrics_dual(eq_df, TRADING_CAP, AUTH_CAP)

    print(f"產出 Excel 檔案: {OUTPUT_FILE}...")
    writer = pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter')

    # 1. Trades
    trades_df.to_excel(writer, sheet_name='Trades', index=False)

    # 2. Trades2
    trades2_df.to_excel(writer, sheet_name='Trades2', index=False)

    # 3. Equity_Curve
    eq_df.to_excel(writer, sheet_name='Equity_Curve', index=False)

    # 4. Equity_Hold (持股明細紀錄)
    daily_details_df.to_excel(writer, sheet_name='Equity_Hold', index=False)

    # 5. Daily (每日數據摘要)
    eq_df.to_excel(writer, sheet_name='Daily', index=False)

    # 6. Summary
    summary_data = [
        ['策略名稱', 'equityV-adj1'],
        ['最初投入資金 (Trading Capital)', TRADING_CAP],
        ['初始授權金額 (Authorized Capital)', AUTH_CAP],
        ['最初投入資金 CAGR', metrics['Trading CAGR']],
        ['初始授權金額 CAGR', metrics['Authorized CAGR']],
        ['標準 MDD (對峰值)', metrics['Standard MaxDD']],
        ['固定基準 MDD (對 1.5 億)', metrics['Fixed Base MaxDD']],
        ['Trading Calmar Ratio', metrics['Trading Calmar']],
        ['', ''],
        ['--- 年度績效 (Actual Trading Mode) ---', '']
    ]

    yearly_df = metrics['Yearly Performance']
    for year, row in yearly_df.iterrows():
        summary_data.append([f'{int(year)} 年度報酬率', row['年度報酬率']])
        summary_data.append([f'{int(year)} 年度損益', row['年度損益']])

    summary_df = pd.DataFrame(summary_data, columns=['指標', '數值'])
    summary_df.to_excel(writer, sheet_name='Summary', index=False)

    # 美化與圖表
    workbook = writer.book

    # 格式設定
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
    percent_fmt = workbook.add_format({'num_format': '0.00%'})
    num_fmt = workbook.add_format({'num_format': '#,##0'})

    # Summary 頁面格式
    summary_sheet = writer.sheets['Summary']
    summary_sheet.set_column('A:A', 35)
    summary_sheet.set_column('B:B', 20)

    # 針對 Summary 的特定行應用百分比格式
    for i in range(3, 8):
        summary_sheet.write(i+1, 1, summary_data[i][1], percent_fmt)
    for i in range(10, len(summary_data), 2):
        summary_sheet.write(i+1, 1, summary_data[i][1], percent_fmt)
        summary_sheet.write(i+2, 1, summary_data[i+1][1], num_fmt)

    # Equity_Curve 圖表
    curve_sheet = writer.sheets['Equity_Curve']
    chart = workbook.add_chart({'type': 'line'})
    max_row = len(eq_df)
    chart.add_series({
        'name': '權益曲線 (Equity Curve)',
        'categories': ['Equity_Curve', 1, 0, max_row, 0],
        'values':     ['Equity_Curve', 1, 1, max_row, 1],
    })
    chart.set_title({'name': '權益增長趨勢 (30M Base)'})
    chart.set_x_axis({'name': '日期'})
    chart.set_y_axis({'name': '總權益'})
    curve_sheet.insert_chart('H2', chart)

    writer.close()
    print(f"完成！檔案已儲存為 {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
