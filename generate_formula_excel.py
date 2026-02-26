import pandas as pd
import xlsxwriter
import numpy as np

# Load raw data to get dimensions and headers
df_raw = pd.read_excel('個股合-1.xlsx', header=None)
# Fill NaNs in headers with empty string for writing
df_raw.iloc[0:2, :] = df_raw.iloc[0:2, :].fillna("")

num_rows = df_raw.shape[0]
num_cols = df_raw.shape[1]
stock_cols = range(2, num_cols)

workbook = xlsxwriter.Workbook('trendstrategy_formulas_equity25.xlsx', {'nan_inf_to_errors': True})

# Formats
header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})

def write_basic_structure(sheet):
    # Write Headers (Stock Codes and Names)
    for c in range(num_cols):
        sheet.write(0, c, df_raw.iloc[0, c], header_format)
        sheet.write(1, c, df_raw.iloc[1, c], header_format)
    # Write Dates
    for r in range(2, num_rows):
        sheet.write(r, 0, df_raw.iloc[r, 0] if pd.notna(df_raw.iloc[r,0]) else "")
        sheet.write(r, 1, df_raw.iloc[r, 1], date_format)

# 1. Prices Sheet
ws_prices = workbook.add_worksheet('Prices')
write_basic_structure(ws_prices)
for r in range(2, num_rows):
    for c in stock_cols:
        val = df_raw.iloc[r, c]
        if pd.notna(val):
            ws_prices.write(r, c, float(val))
        else:
            ws_prices.write(r, c, "")

# 2. SMA64 Sheet
ws_sma = workbook.add_worksheet('SMA64')
write_basic_structure(ws_sma)
for r in range(2, num_rows):
    excel_row = r + 1
    for c in stock_cols:
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        if excel_row >= 66:
            formula = f"=AVERAGE(Prices!{col_letter}{excel_row-63}:Prices!{col_letter}{excel_row})"
            ws_sma.write_formula(r, c, formula)

# 3. ROC23 Sheet
ws_roc = workbook.add_worksheet('ROC23')
write_basic_structure(ws_roc)
for r in range(2, num_rows):
    excel_row = r + 1
    for c in stock_cols:
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        if excel_row >= 26:
            formula = f"=IF(Prices!{col_letter}{excel_row-23}<>0, (Prices!{col_letter}{excel_row}-Prices!{col_letter}{excel_row-23})/Prices!{col_letter}{excel_row-23}, \"\")"
            ws_roc.write_formula(r, c, formula)

# 4. Eligibility Sheet
ws_elig = workbook.add_worksheet('Eligibility')
write_basic_structure(ws_elig)
for r in range(2, num_rows):
    excel_row = r + 1
    for c in stock_cols:
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        formula = f"=IF(AND(Prices!{col_letter}{excel_row}>SMA64!{col_letter}{excel_row}, ISNUMBER(ROC23!{col_letter}{excel_row}), ROC23!{col_letter}{excel_row}>0), ROC23!{col_letter}{excel_row}, -1000000)"
        ws_elig.write_formula(r, c, formula)

# 5. Rank Sheet
ws_rank = workbook.add_worksheet('Rank')
write_basic_structure(ws_rank)
last_col_letter = xlsxwriter.utility.xl_col_to_name(num_cols - 1)
for r in range(2, num_rows):
    excel_row = r + 1
    for c in stock_cols:
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        formula = f"=IF(Eligibility!{col_letter}{excel_row}>-999999, RANK(Eligibility!{col_letter}{excel_row}, Eligibility!$C{excel_row}:${last_col_letter}{excel_row}), \"\")"
        ws_rank.write_formula(r, c, formula)

# 6. Signals Sheet
ws_signals = workbook.add_worksheet('Signals')
write_basic_structure(ws_signals)
for r in range(2, num_rows):
    excel_row = r + 1
    for c in stock_cols:
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        formula = f"=IF(AND(Rank!{col_letter}{excel_row}<>\"\", Rank!{col_letter}{excel_row}<=3), \"Buy\", \"\")"
        ws_signals.write_formula(r, c, formula)

# 7. Shares Sheet
ws_shares = workbook.add_worksheet('Shares')
write_basic_structure(ws_shares)
for r in range(2, num_rows):
    excel_row = r + 1
    for c in stock_cols:
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        formula = f"=IF(Signals!{col_letter}{excel_row}=\"Buy\", 10000000/Prices!{col_letter}{excel_row}, 0)"
        ws_shares.write_formula(r, c, formula)

# 8. StopLoss Tracker
ws_sl_max = workbook.add_worksheet('StopLoss_Max')
write_basic_structure(ws_sl_max)
for r in range(2, num_rows):
    excel_row = r + 1
    for c in stock_cols:
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        if r == 2:
            formula = f"=IF(Signals!{col_letter}{excel_row}=\"Buy\", Prices!{col_letter}{excel_row}, 0)"
        else:
            formula = f"=IF(Signals!{col_letter}{excel_row}=\"Buy\", MAX(Prices!{col_letter}{excel_row}, {col_letter}{excel_row-1}), 0)"
        ws_sl_max.write_formula(r, c, formula)

# 9. StopLoss Signal
ws_sl_sig = workbook.add_worksheet('StopLoss_Signal')
write_basic_structure(ws_sl_sig)
for r in range(2, num_rows):
    excel_row = r + 1
    for c in stock_cols:
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        formula = f"=IF(AND(StopLoss_Max!{col_letter}{excel_row}>0, Prices!{col_letter}{excel_row}<StopLoss_Max!{col_letter}{excel_row}*0.91), \"STOP\", \"OK\")"
        ws_sl_sig.write_formula(r, c, formula)

workbook.close()
print("trendstrategy_formulas_equity25.xlsx generated.")
