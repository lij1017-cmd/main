import pandas as pd
import xlsxwriter
import numpy as np

# Load raw data to get dimensions and headers
df_raw = pd.read_excel('個股合-1.xlsx', header=None)
df_raw.iloc[0:2, :] = df_raw.iloc[0:2, :].fillna("")

num_rows = df_raw.shape[0]
num_cols = df_raw.shape[1]
stock_cols = range(2, num_cols)

workbook = xlsxwriter.Workbook('trendstrategy_formulas_equity25-1.xlsx', {'nan_inf_to_errors': True})

# Formats
header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})

def write_basic_structure(sheet):
    for c in range(num_cols):
        sheet.write(0, c, df_raw.iloc[0, c], header_format)
        sheet.write(1, c, df_raw.iloc[1, c], header_format)
    for r in range(2, num_rows):
        sheet.write(r, 0, df_raw.iloc[r, 0] if pd.notna(df_raw.iloc[r,0]) else "")
        sheet.write(r, 1, df_raw.iloc[r, 1], date_format)

# Sheets: Prices, SMA64, ROC23, Eligibility, Rank
ws_prices = workbook.add_worksheet('Prices')
write_basic_structure(ws_prices)
for r in range(2, num_rows):
    for c in stock_cols:
        val = df_raw.iloc[r, c]
        if pd.notna(val): ws_prices.write(r, c, float(val))
        else: ws_prices.write(r, c, "")

ws_sma = workbook.add_worksheet('SMA64')
write_basic_structure(ws_sma)
for r in range(2, num_rows):
    excel_row = r + 1
    for c in stock_cols:
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        if excel_row >= 66: ws_sma.write_formula(r, c, f"=AVERAGE(Prices!{col_letter}{excel_row-63}:Prices!{col_letter}{excel_row})")

ws_roc = workbook.add_worksheet('ROC23')
write_basic_structure(ws_roc)
for r in range(2, num_rows):
    excel_row = r + 1
    for c in stock_cols:
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        if excel_row >= 26: ws_roc.write_formula(r, c, f"=IF(Prices!{col_letter}{excel_row-23}<>0, (Prices!{col_letter}{excel_row}-Prices!{col_letter}{excel_row-23})/Prices!{col_letter}{excel_row-23}, \"\")")

ws_elig = workbook.add_worksheet('Eligibility')
write_basic_structure(ws_elig)
for r in range(2, num_rows):
    excel_row = r + 1
    for c in stock_cols:
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        ws_elig.write_formula(r, c, f"=IF(AND(Prices!{col_letter}{excel_row}>SMA64!{col_letter}{excel_row}, ISNUMBER(ROC23!{col_letter}{excel_row}), ROC23!{col_letter}{excel_row}>0), ROC23!{col_letter}{excel_row}, -1000000)")

ws_rank = workbook.add_worksheet('Rank')
write_basic_structure(ws_rank)
last_col_letter = xlsxwriter.utility.xl_col_to_name(num_cols - 1)
for r in range(2, num_rows):
    excel_row = r + 1
    for c in stock_cols:
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        ws_rank.write_formula(r, c, f"=IF(Eligibility!{col_letter}{excel_row}>-999999, RANK(Eligibility!{col_letter}{excel_row}, Eligibility!$C{excel_row}:${last_col_letter}{excel_row}), \"\")")

# Coresheets
ws_status = workbook.add_worksheet('Status')
write_basic_structure(ws_status)
ws_maxp = workbook.add_worksheet('MaxPriceTracker')
write_basic_structure(ws_maxp)
ws_signals = workbook.add_worksheet('Signals')
write_basic_structure(ws_signals)
ws_shares = workbook.add_worksheet('Shares')
write_basic_structure(ws_shares)

start_row = 67
for r in range(2, num_rows):
    excel_row = r + 1
    for c in stock_cols:
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        if excel_row < start_row:
            ws_status.write(r, c, 0)
            ws_maxp.write(r, c, 0)
            ws_signals.write(r, c, "")
            ws_shares.write(r, c, 0)
        else:
            if excel_row < start_row + 2:
                ws_status.write(r, c, 0)
            else:
                ws_status.write_formula(r, c, f"=IF(Signals!{col_letter}{excel_row-2}=\"BUY\", 1, IF(OR(Signals!{col_letter}{excel_row-2}=\"SELL\", Signals!{col_letter}{excel_row-2}=\"STOP\"), 0, Status!{col_letter}{excel_row-1}))")

            ws_maxp.write_formula(r, c, f"=IF(Status!{col_letter}{excel_row}=0, 0, IF(Status!{col_letter}{excel_row-1}=0, Prices!{col_letter}{excel_row}, MAX(Prices!{col_letter}{excel_row}, MaxPriceTracker!{col_letter}{excel_row-1})))")

            is_reb = f"MOD({excel_row}-{start_row}, 5)=0"
            ws_signals.write_formula(r, c,
                f"=IF(Status!{col_letter}{excel_row}=1, "
                f"IF(Prices!{col_letter}{excel_row}<MaxPriceTracker!{col_letter}{excel_row}*0.91, \"STOP\", IF(AND({is_reb}, Rank!{col_letter}{excel_row}>3), \"SELL\", \"HOLD\")), "
                f"IF(AND({is_reb}, Rank!{col_letter}{excel_row}<>\"\", Rank!{col_letter}{excel_row}<=3), \"BUY\", \"\")"
                f")")

            if excel_row < num_rows:
                ws_shares.write_formula(r, c, f"=IF(Signals!{col_letter}{excel_row}=\"BUY\", 10000000/Prices!{col_letter}{excel_row+1}, 0)")
            else:
                ws_shares.write(r, c, 0)

workbook.close()
