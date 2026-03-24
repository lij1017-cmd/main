import pandas as pd
import numpy as np
import xlsxwriter

def clean_data(filepath):
    """
    清洗並預處理輸入的 Excel 資料檔。
    """
    df_raw = pd.read_excel(filepath, header=None)
    stock_codes = df_raw.iloc[0, 2:].values
    stock_names = df_raw.iloc[1, 2:].values
    dates = pd.to_datetime(df_raw.iloc[2:, 1])
    prices = df_raw.iloc[2:, 2:].astype(float)
    prices.index = dates
    prices.columns = stock_codes
    code_to_name = dict(zip(stock_codes, stock_names))
    prices = prices.ffill().bfill()
    return prices, code_to_name

from backtest_engine import Backtester, calculate_metrics

def calculate_metrics(eq_df):
    if eq_df is None or eq_df.empty: return 0, 0, 0
    equity = eq_df['權益']
    total_return = (equity.iloc[-1] / equity.iloc[0]) - 1
    days = (eq_df['日期'].iloc[-1] - eq_df['日期'].iloc[0]).days
    years = days / 365.25
    cagr = (1 + total_return) ** (1 / years) - 1 if years > 0 else 0
    rolling_max = equity.cummax()
    drawdowns = (equity - rolling_max) / rolling_max
    max_dd = drawdowns.min()
    calmar = cagr / abs(max_dd) if max_dd != 0 else 0
    return cagr, max_dd, calmar

def main():
    DATA_FILE = '個股合-1.xlsx'
    # 採用 SMA87 基準參數 (對應 walk-forward-3 需求)
    SMA_PERIOD = 87
    ROC_PERIOD = 54
    STOP_LOSS_PCT = 0.09
    REBALANCE = 6
    INITIAL_CAPITAL = 30000000

    # 最新要求的 WFA 區間
    periods = [
        ('2024-06-01', '2025-12-31'),
        ('2024-01-02', '2025-05-31'),
        ('2023-01-02', '2024-12-31'),
        ('2022-01-02', '2024-05-31'),
        ('2021-06-01', '2023-12-31'),
        ('2021-01-02', '2023-05-31'),
        ('2020-01-02', '2022-12-31'),
        ('2019-06-01', '2022-05-31'),
        ('2019-01-02', '2021-12-31'),
    ]

    prices, code_to_name = clean_data(DATA_FILE)
    bt = Backtester(prices, code_to_name, INITIAL_CAPITAL)

    summary_results = []
    all_equity_curves = []

    for start_str, end_str in periods:
        print(f"Executing WFA: {start_str} to {end_str}")
        eq_df, t_log, h_log, t2_log, d_log = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE, start_str, end_str)
        cagr, mdd, calmar = calculate_metrics(eq_df)
        # 僅計算實際的 '買進' 與 '賣出'，排除 '保持'
        trades = len(t_log[t_log['狀態'].isin(['買進', '賣出'])])

        period_label = f"{start_str} - {end_str}"
        summary_results.append({
            '回測期間': period_label,
            'CAGR': cagr,
            'MaxDD': mdd,
            'Calmar Ratio': calmar,
            '交易次數': trades
        })

        temp_eq = eq_df[['日期', '權益']].copy()
        temp_eq.columns = [f'日期_{period_label}', f'權益_{period_label}']
        all_equity_curves.append(temp_eq)

    summary_df = pd.DataFrame(summary_results)
    OUTPUT_FILE = 'walk-forward-3.xlsx'
    writer = pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter')
    summary_df.to_excel(writer, sheet_name='Summary', index=False)
    workbook = writer.book
    summary_sheet = writer.sheets['Summary']
    percent_fmt = workbook.add_format({'num_format': '0.00%'})
    num_fmt = workbook.add_format({'num_format': '0.00'})
    summary_sheet.set_column('B:C', 15, percent_fmt)
    summary_sheet.set_column('D:D', 15, num_fmt)
    summary_sheet.set_column('A:A', 30)

    curves_df = pd.concat(all_equity_curves, axis=1)
    curves_df.to_excel(writer, sheet_name='Equity_Curves', index=False)
    curves_sheet = writer.sheets['Equity_Curves']

    for idx, (start_str, end_str) in enumerate(periods):
        period_label = f"{start_str} - {end_str}"
        chart = workbook.add_chart({'type': 'line'})
        date_col = 2 * idx
        val_col = 2 * idx + 1
        max_row = len(all_equity_curves[idx])
        chart.add_series({
            'name':       f'Equity {period_label}',
            'categories': ['Equity_Curves', 1, date_col, max_row, date_col],
            'values':     ['Equity_Curves', 1, val_col, max_row, val_col],
        })
        chart.set_title({'name': f'Equity Curve ({period_label})'})
        chart.set_legend({'position': 'none'})
        row_pos = idx * 15
        curves_sheet.insert_chart(row_pos, 2 * len(periods) + 2, chart)

    writer.close()
    print(f"WFA Done: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
