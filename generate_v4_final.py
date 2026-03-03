import pandas as pd
import xlsxwriter
import numpy as np

# Load raw data to get dimensions and headers
# Using all stocks (approx 131)
df_raw = pd.read_excel('個股合-1.xlsx', header=None)
num_rows = df_raw.shape[0]
num_cols = df_raw.shape[1]
stock_cols = range(2, num_cols)
last_col_letter = xlsxwriter.utility.xl_col_to_name(num_cols - 1)

output_filename = 'trendstrategy_formulas_equity25-3.xlsx'
workbook = xlsxwriter.Workbook(output_filename, {'nan_inf_to_errors': True})

# Formats
header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1, 'align': 'center'})
date_format = workbook.add_format({'num_format': 'yyyy-mm-dd', 'align': 'center'})
pct_format = workbook.add_format({'num_format': '0.00%', 'align': 'center'})
price_format = workbook.add_format({'num_format': '#,##0.00', 'align': 'center'})
num_format = workbook.add_format({'num_format': '#,##0', 'align': 'center'})
text_format = workbook.add_format({'align': 'center'})

def write_basic_structure(sheet):
    for c in range(num_cols):
        val0 = df_raw.iloc[0, c] if pd.notna(df_raw.iloc[0, c]) else ""
        val1 = df_raw.iloc[1, c] if pd.notna(df_raw.iloc[1, c]) else ""
        sheet.write(0, c, val0, header_format)
        sheet.write(1, c, val1, header_format)
    for r in range(2, num_rows):
        sheet.write(r, 0, df_raw.iloc[r, 0] if pd.notna(df_raw.iloc[r, 0]) else "", text_format)
        sheet.write(r, 1, df_raw.iloc[r, 1], date_format)

# 1. Prices Sheet
ws_prices = workbook.add_worksheet('Prices')
write_basic_structure(ws_prices)
for r in range(2, num_rows):
    for c in stock_cols:
        val = df_raw.iloc[r, c]
        if pd.notna(val):
            ws_prices.write(r, c, float(val), price_format)
        else:
            ws_prices.write(r, c, "")

# Strategy Parameters
SMA_P = 69
ROC_P = 23
SL_PCT = 0.09
INITIAL_CAPITAL = 30000000
SLOT_CAPITAL = 10000000
NUM_SLOTS = 3

# Logic Sheets
ws_sma = workbook.add_worksheet('SMA')
write_basic_structure(ws_sma)
ws_roc = workbook.add_worksheet('ROC')
write_basic_structure(ws_roc)
ws_elig = workbook.add_worksheet('Eligible')
write_basic_structure(ws_elig)
ws_rank = workbook.add_worksheet('Rank')
write_basic_structure(ws_rank)
ws_rebal = workbook.add_worksheet('RebalanceDay')
write_basic_structure(ws_rebal)
ws_status = workbook.add_worksheet('Status')
write_basic_structure(ws_status)
ws_shares = workbook.add_worksheet('Shares')
write_basic_structure(ws_shares)
ws_maxp = workbook.add_worksheet('MaxPrice')
write_basic_structure(ws_maxp)
ws_slt = workbook.add_worksheet('SL_Trigger')
write_basic_structure(ws_slt)
ws_decide = workbook.add_worksheet('Decision')
write_basic_structure(ws_decide)
ws_entry_row = workbook.add_worksheet('EntryRow')
write_basic_structure(ws_entry_row)
ws_trade_idx = workbook.add_worksheet('TradeIndex')
write_basic_structure(ws_trade_idx)

start_data_row = max(SMA_P, ROC_P) + 2

for r in range(2, num_rows):
    excel_row = r + 1
    val_rebal = f"=IF({excel_row}>={start_data_row}, MOD({excel_row}-{start_data_row}, 5)=0, FALSE)"
    ws_rebal.write_formula(r, 2, val_rebal)

    for c in stock_cols:
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        if excel_row >= SMA_P + 2:
            ws_sma.write_formula(r, c, f"=AVERAGE(Prices!{col_letter}{excel_row-SMA_P+1}:Prices!{col_letter}{excel_row})", price_format)
        if excel_row >= ROC_P + 2:
            ws_roc.write_formula(r, c, f"=IF(Prices!{col_letter}{excel_row-ROC_P}<>0, (Prices!{col_letter}{excel_row}-Prices!{col_letter}{excel_row-ROC_P})/Prices!{col_letter}{excel_row-ROC_P}, \"\")", pct_format)
        ws_elig.write_formula(r, c, f"=IF(AND(Prices!{col_letter}{excel_row}>SMA!{col_letter}{excel_row}, ISNUMBER(ROC!{col_letter}{excel_row}), ROC!{col_letter}{excel_row}>0), ROC!{col_letter}{excel_row}, -1)", pct_format)
        ws_rank.write_formula(r, c, f"=IF(Eligible!{col_letter}{excel_row}>0, RANK(Eligible!{col_letter}{excel_row}, Eligible!$C{excel_row}:${last_col_letter}{excel_row}), \"\")")

        if excel_row < start_data_row:
            ws_status.write(r, c, "Cash", text_format)
            ws_shares.write(r, c, 0, num_format)
            ws_maxp.write(r, c, 0, price_format)
            ws_slt.write(r, c, False)
            ws_decide.write(r, c, "")
            ws_entry_row.write(r, c, 0)
            ws_trade_idx.write(r, c, "")
        else:
            ws_status.write_formula(r, c, f'=IF(Decision!{col_letter}{excel_row-1}="BUY", "Holding", IF(Decision!{col_letter}{excel_row-1}="SELL", "Cash", Status!{col_letter}{excel_row-1}))', text_format)
            ws_shares.write_formula(r, c, f'=IF(Decision!{col_letter}{excel_row-1}="BUY", INT({SLOT_CAPITAL}/Prices!{col_letter}{excel_row}), IF(Decision!{col_letter}{excel_row-1}="SELL", 0, Shares!{col_letter}{excel_row-1}))', num_format)
            ws_maxp.write_formula(r, c, f'=IF(Status!{col_letter}{excel_row}="Holding", IF(Decision!{col_letter}{excel_row-1}="BUY", Prices!{col_letter}{excel_row}, MAX(Prices!{col_letter}{excel_row}, MaxPrice!{col_letter}{excel_row-1})), 0)', price_format)
            ws_slt.write_formula(r, c, f'=AND(Status!{col_letter}{excel_row}="Holding", Prices!{col_letter}{excel_row}<MaxPrice!{col_letter}{excel_row}*(1-{SL_PCT}))')
            ws_decide.write_formula(r, c, f'=IF(Status!{col_letter}{excel_row}="Cash", IF(AND(RebalanceDay!$C{excel_row}, Rank!{col_letter}{excel_row}<>"", Rank!{col_letter}{excel_row}<={NUM_SLOTS}), "BUY", ""), IF(OR(AND(RebalanceDay!$C{excel_row}, OR(Rank!{col_letter}{excel_row}="", Rank!{col_letter}{excel_row}>{NUM_SLOTS})), SL_Trigger!{col_letter}{excel_row}), "SELL", "Hold"))', text_format)
            ws_entry_row.write_formula(r, c, f'=IF(Decision!{col_letter}{excel_row-1}="BUY", {excel_row}, IF(Status!{col_letter}{excel_row}="Holding", EntryRow!{col_letter}{excel_row-1}, 0))')
            ws_trade_idx.write_formula(r, c, f'=IF(Decision!{col_letter}{excel_row}="SELL", {excel_row}*10000+{c}, "")')

# Equity Calc
ws_equity_calc = workbook.add_worksheet('EquityCalc')
ws_equity_calc.write(0, 0, "日期", header_format)
ws_equity_calc.write(0, 1, "現金", header_format)
ws_equity_calc.write(0, 2, "市值", header_format)
ws_equity_calc.write(0, 3, "總資產", header_format)
ws_equity_calc.write(0, 4, "最高資產", header_format)
ws_equity_calc.write(0, 5, "回撤", header_format)

for r in range(2, num_rows):
    excel_row = r + 1
    ws_equity_calc.write(r, 0, df_raw.iloc[r, 1], date_format)
    if excel_row == start_data_row:
        cash_f = f"={INITIAL_CAPITAL} + SUMPRODUCT((Decision!$C{excel_row-1}:${last_col_letter}{excel_row-1}=\"SELL\")*Shares!$C{excel_row-1}:${last_col_letter}{excel_row-1}*Prices!$C{excel_row}:${last_col_letter}{excel_row}) - SUMPRODUCT((Decision!$C{excel_row-1}:${last_col_letter}{excel_row-1}=\"BUY\")*{SLOT_CAPITAL})"
        ws_equity_calc.write_formula(r, 1, cash_f, num_format)
    elif excel_row > start_data_row:
        cash_f = f"=B{excel_row-1} + SUMPRODUCT((Decision!$C{excel_row-1}:${last_col_letter}{excel_row-1}=\"SELL\")*Shares!$C{excel_row-1}:${last_col_letter}{excel_row-1}*Prices!$C{excel_row}:${last_col_letter}{excel_row}) - SUMPRODUCT((Decision!$C{excel_row-1}:${last_col_letter}{excel_row-1}=\"BUY\")*{SLOT_CAPITAL})"
        ws_equity_calc.write_formula(r, 1, cash_f, num_format)
    else:
        ws_equity_calc.write(r, 1, INITIAL_CAPITAL, num_format)

    ws_equity_calc.write_formula(r, 2, f"=SUMPRODUCT(Shares!$C{excel_row}:${last_col_letter}{excel_row}*Prices!$C{excel_row}:${last_col_letter}{excel_row})", num_format)
    ws_equity_calc.write_formula(r, 3, f"=B{excel_row}+C{excel_row}", num_format)
    if excel_row == 3:
        ws_equity_calc.write_formula(r, 4, f"=D{excel_row}", num_format)
    else:
        ws_equity_calc.write_formula(r, 4, f"=MAX(D$3:D{excel_row})", num_format)
    ws_equity_calc.write_formula(r, 5, f"=IF(E{excel_row}<>0, (D{excel_row}-E{excel_row})/E{excel_row}, 0)", pct_format)

# Performance Sheet - Support up to 5000 trades to ensure "unlimited" feel
ws_perf = workbook.add_worksheet('Performance')
perf_headers = ["股票代號", "進場日期", "出場日期", "進場價格", "出場價格", "持有天數", "報酬率 (%)"]
for i, h in enumerate(perf_headers):
    ws_perf.write(0, i, h, header_format)

trade_range = f"TradeIndex!$C$3:${last_col_letter}${num_rows}"
for i in range(1, 5001):
    excel_row = i + 1
    ws_perf.write_formula(i, 7, f"=IFERROR(SMALL({trade_range}, {i}), \"\")")
    idx_cell = f"H{excel_row}"
    ws_perf.write_formula(i, 8, f"=IF({idx_cell}<>\"\", INT({idx_cell}/10000), \"\")")
    ws_perf.write_formula(i, 9, f"=IF({idx_cell}<>\"\", MOD({idx_cell}, 10000), \"\")")
    entry_row_f = f'=IF(I{excel_row}<>"", INDEX(EntryRow!$A$1:${last_col_letter}${num_rows}, I{excel_row}, J{excel_row}+1), "")'
    ws_perf.write_formula(i, 10, f"=IF({idx_cell}<>\"\", {entry_row_f}, \"\")")

    ws_perf.write_formula(i, 0, f'=IF(J{excel_row}<>"", INDEX(Prices!$1:$1, J{excel_row}+1), "")', text_format)
    ws_perf.write_formula(i, 1, f'=IF(K{excel_row}<>"", INDEX(Prices!$B$1:$B${num_rows}, K{excel_row}), "")', date_format)
    ws_perf.write_formula(i, 2, f'=IF(I{excel_row}<>"", INDEX(Prices!$B$1:$B${num_rows}, I{excel_row}), "")', date_format)
    ws_perf.write_formula(i, 3, f'=IF(K{excel_row}<>"", INDEX(Prices!$1:${last_col_letter}${num_rows}, K{excel_row}, J{excel_row}+1), "")', price_format)
    ws_perf.write_formula(i, 4, f'=IF(I{excel_row}<>"", INDEX(Prices!$1:${last_col_letter}${num_rows}, I{excel_row}, J{excel_row}+1), "")', price_format)
    ws_perf.write_formula(i, 5, f'=IF(C{excel_row}<>"", C{excel_row}-B{excel_row}, "")', num_format)
    ws_perf.write_formula(i, 6, f'=IF(D{excel_row}<>0, (E{excel_row}-D{excel_row})/D{excel_row}, "")', pct_format)

# Stats
ws_stats = workbook.add_worksheet('總績效統計')
stats = [
    ("總交易次數", "=COUNT(Performance!C:C)"),
    ("勝率", "=IF(COUNT(Performance!G:G)>0, COUNTIF(Performance!G:G, \">0\")/COUNT(Performance!G:G), 0)"),
    ("平均報酬率", "=IFERROR(AVERAGE(Performance!G:G), 0)"),
    ("最大回撤", "=MIN(EquityCalc!F:F)"),
    ("最終資產", f"=INDEX(EquityCalc!D:D, {num_rows})"),
    ("年化報酬率", f"=(INDEX(EquityCalc!D:D, {num_rows})/{INITIAL_CAPITAL})^(252/COUNT(EquityCalc!A:A)) - 1")
]
for i, (label, formula) in enumerate(stats):
    ws_stats.write(i, 0, label, header_format)
    ws_stats.write_formula(i, 1, formula, pct_format if "%" in label or "率" in label or "回撤" in label else num_format)

# Equity Curve
ws_curve = workbook.add_worksheet('Equity Curve')
ws_curve.write(0, 0, "日期", header_format)
ws_curve.write(0, 1, "總資產", header_format)
for r in range(2, num_rows):
    excel_row = r + 1
    ws_curve.write_formula(r, 0, f"=EquityCalc!A{excel_row}", date_format)
    ws_curve.write_formula(r, 1, f"=EquityCalc!D{excel_row}", num_format)
chart = workbook.add_chart({'type': 'line'})
chart.add_series({
    'name': 'Total Assets',
    'categories': f"='Equity Curve'!$A$3:$A${num_rows}",
    'values':     f"='Equity Curve'!$B$3:$B${num_rows}",
})
chart.set_title({'name': 'Equity Curve'})
ws_curve.insert_chart('D2', chart, {'x_scale': 2, 'y_scale': 2})

# Dashboard
ws_dash = workbook.add_worksheet('Dashboard')
ws_dash.write(0, 0, "策略績效摘要", header_format)
for i, (label, _) in enumerate(stats):
    ws_dash.write(i+1, 0, label, text_format)
    ws_dash.write_formula(i+1, 1, f"=總績效統計!B{i+1}", pct_format if "%" in label or "率" in label or "回撤" in label else num_format)

ws_dash.write(0, 3, "當前持股與建議", header_format)
dash_headers = ["日期", "股票代號", "標的名稱", "狀態", "建議動作", "股數", "動能(ROC)"]
for i, h in enumerate(dash_headers):
    ws_dash.write(1, i+3, h, header_format)

last_row_idx = num_rows
for i, c in enumerate(stock_cols):
    row_offset = 2 + i
    col_letter = xlsxwriter.utility.xl_col_to_name(c)
    ws_dash.write_formula(row_offset, 3, f"=Prices!B{last_row_idx}", date_format)
    ws_dash.write_formula(row_offset, 4, f"=Prices!{col_letter}1", text_format)
    ws_dash.write_formula(row_offset, 5, f"=Prices!{col_letter}2", text_format)
    ws_dash.write_formula(row_offset, 6, f"=Status!{col_letter}{last_row_idx}", text_format)
    ws_dash.write_formula(row_offset, 7, f"=Decision!{col_letter}{last_row_idx}", text_format)
    ws_dash.write_formula(row_offset, 8, f"=Shares!{col_letter}{last_row_idx}", num_format)
    ws_dash.write_formula(row_offset, 9, f"=ROC!{col_letter}{last_row_idx}", pct_format)

ws_dash.set_column('A:B', 15)
ws_dash.set_column('D:K', 15)
ws_perf.set_column('A:G', 15)
ws_curve.set_column('A:B', 15)

workbook.close()
print(f"{output_filename} generated successfully.")
