import pandas as pd
import xlsxwriter
import numpy as np

# Load raw data to get dimensions and headers
df_raw = pd.read_excel('個股合-1.xlsx', header=None)
num_rows = df_raw.shape[0]
num_cols = df_raw.shape[1]
stock_cols = range(2, num_cols)
last_col_letter = xlsxwriter.utility.xl_col_to_name(num_cols - 1)

workbook = xlsxwriter.Workbook('trendstrategy_formulas_equity25-1.xlsx', {'nan_inf_to_errors': True})

# Formats
header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
pct_format = workbook.add_format({'num_format': '0.00%'})
price_format = workbook.add_format({'num_format': '#,##0.00'})

def write_basic_structure(sheet):
    for c in range(num_cols):
        val0 = df_raw.iloc[0, c] if pd.notna(df_raw.iloc[0, c]) else ""
        val1 = df_raw.iloc[1, c] if pd.notna(df_raw.iloc[1, c]) else ""
        sheet.write(0, c, val0, header_format)
        sheet.write(1, c, val1, header_format)
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
        if excel_row >= 66: # 2 header rows + 64 data points
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
            ws_roc.write_formula(r, c, formula, pct_format)

# 4. Eligible Sheet
ws_elig = workbook.add_worksheet('Eligible')
write_basic_structure(ws_elig)
for r in range(2, num_rows):
    excel_row = r + 1
    for c in stock_cols:
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        # Price > SMA and ROC > 0
        formula = f"=IF(AND(Prices!{col_letter}{excel_row}>SMA64!{col_letter}{excel_row}, ISNUMBER(ROC23!{col_letter}{excel_row}), ROC23!{col_letter}{excel_row}>0), ROC23!{col_letter}{excel_row}, -1)"
        ws_elig.write_formula(r, c, formula, pct_format)

# 5. Rank Sheet
ws_rank = workbook.add_worksheet('Rank')
write_basic_structure(ws_rank)
for r in range(2, num_rows):
    excel_row = r + 1
    for c in stock_cols:
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        # Rank among all stocks for that day
        formula = f"=IF(Eligible!{col_letter}{excel_row}>0, RANK(Eligible!{col_letter}{excel_row}, Eligible!$C{excel_row}:${last_col_letter}{excel_row}), \"\")"
        ws_rank.write_formula(r, c, formula)

# 6. StopLoss_Max Sheet (Tracks peak price since holding)
# To handle T signal -> T+1 execution, we need to know if we were holding on day T.
ws_sl_max = workbook.add_worksheet('StopLoss_Max')
write_basic_structure(ws_sl_max)
# We will define Status/Portfolio first or use them here? Let's use Portfolio sheet later.
# For simplicity, let's put it all in Portfolio sheet.

# 7. Portfolio Sheet
# This will be the main logic sheet.
# Columns per stock: [Status, Shares, MaxPrice, StopLossTrigger, BuySignal, SellSignal]
# This might make the sheet too wide.
# Better: Separate sheets for Status, Shares, MaxPrice, SL_Trigger, Decision.

# 7a. Rebalance Day Check
ws_rebal = workbook.add_worksheet('RebalanceDay')
write_basic_structure(ws_rebal)
start_data_row = 66 # Start of strategy
for r in range(2, num_rows):
    excel_row = r + 1
    val = f"=IF({excel_row}>={start_data_row}, MOD({excel_row}-{start_data_row}, 5)=0, FALSE)"
    ws_rebal.write_formula(r, 2, val) # Just use one column for simplicity

# 7b. MaxPrice (Peak since buy)
ws_maxp = workbook.add_worksheet('MaxPrice')
write_basic_structure(ws_maxp)

# 7c. SL_Trigger (Price < MaxPrice * 0.91)
ws_slt = workbook.add_worksheet('SL_Trigger')
write_basic_structure(ws_slt)

# 7d. Status (Holding or Cash)
ws_status = workbook.add_worksheet('Status')
write_basic_structure(ws_status)

# 7e. Shares
ws_shares = workbook.add_worksheet('Shares')
write_basic_structure(ws_shares)

# 7f. Decisions (Made at T close)
ws_decide = workbook.add_worksheet('Decision')
write_basic_structure(ws_decide)

for r in range(2, num_rows):
    excel_row = r + 1
    for c in stock_cols:
        col_letter = xlsxwriter.utility.xl_col_to_name(c)

        # Decision at T close (excel_row)
        # Buy if RebalanceDay and Rank <= 3 and Status was Cash
        # Sell if Holding and ( (RebalanceDay and Rank > 3) or StopLossTrigger )

        # We need a circular reference or a specific order.
        # Status[T] depends on Decision[T-1].

        if excel_row < start_data_row:
            ws_status.write(r, c, "Cash")
            ws_shares.write(r, c, 0)
            ws_maxp.write(r, c, 0)
            ws_slt.write(r, c, False)
            ws_decide.write(r, c, "")
        else:
            # Status at T (excel_row) depends on Decision at T-1 (excel_row-1)
            status_formula = f'=IF(Decision!{col_letter}{excel_row-1}="BUY", "Holding", IF(Decision!{col_letter}{excel_row-1}="SELL", "Cash", Status!{col_letter}{excel_row-1}))'
            ws_status.write_formula(r, c, status_formula)

            # Shares at T depends on Decision at T-1
            # If Decision[T-1] was BUY, Shares = 10,000,000 / Prices!{col_letter}{excel_row} (T execution)
            shares_formula = f'=IF(Decision!{col_letter}{excel_row-1}="BUY", 10000000/Prices!{col_letter}{excel_row}, IF(Decision!{col_letter}{excel_row-1}="SELL", 0, Shares!{col_letter}{excel_row-1}))'
            ws_shares.write_formula(r, c, shares_formula)

            # MaxPrice at T
            # If Status is Holding, MAX(Prices!{col_letter}{excel_row}, MaxPrice!{col_letter}{excel_row-1})
            # If Status just became Holding (Decision[T-1]=="BUY"), start with Prices!{col_letter}{excel_row}
            maxp_formula = f'=IF(Status!{col_letter}{excel_row}="Holding", IF(Decision!{col_letter}{excel_row-1}="BUY", Prices!{col_letter}{excel_row}, MAX(Prices!{col_letter}{excel_row}, MaxPrice!{col_letter}{excel_row-1})), 0)'
            ws_maxp.write_formula(r, c, maxp_formula)

            # SL Trigger at T
            slt_formula = f'=AND(Status!{col_letter}{excel_row}="Holding", Prices!{col_letter}{excel_row}<MaxPrice!{col_letter}{excel_row}*0.91)'
            ws_slt.write_formula(r, c, slt_formula)

            # Decision at T close
            decide_formula = f'=IF(Status!{col_letter}{excel_row}="Cash", IF(AND(RebalanceDay!$C{excel_row}, Rank!{col_letter}{excel_row}<>"", Rank!{col_letter}{excel_row}<=3), "BUY", ""), IF(OR(AND(RebalanceDay!$C{excel_row}, OR(Rank!{col_letter}{excel_row}="", Rank!{col_letter}{excel_row}>3)), SL_Trigger!{col_letter}{excel_row}), "SELL", "Hold"))'
            ws_decide.write_formula(r, c, decide_formula)

# 8. Summary / Dashboard Sheet
ws_dash = workbook.add_worksheet('Dashboard')
ws_dash.write(0, 0, "當前建議 (基於最後一列數據)", header_format)
ws_dash.write(1, 0, "日期", header_format)
ws_dash.write(1, 1, "股票代號", header_format)
ws_dash.write(1, 2, "標的名稱", header_format)
ws_dash.write(1, 3, "狀態", header_format)
ws_dash.write(1, 4, "建議動作", header_format)
ws_dash.write(1, 5, "股數", header_format)
ws_dash.write(1, 6, "動能(ROC)", header_format)

# Use Excel formulas to pull the last row's data
last_row_idx = num_rows # excel row
for i, c in enumerate(stock_cols):
    row_offset = 2 + i
    col_letter = xlsxwriter.utility.xl_col_to_name(c)
    ws_dash.write_formula(row_offset, 0, f"=Prices!B{last_row_idx}")
    ws_dash.write_formula(row_offset, 1, f"=Prices!{col_letter}1")
    ws_dash.write_formula(row_offset, 2, f"=Prices!{col_letter}2")
    ws_dash.write_formula(row_offset, 3, f"=Status!{col_letter}{last_row_idx}")
    ws_dash.write_formula(row_offset, 4, f"=Decision!{col_letter}{last_row_idx}")
    ws_dash.write_formula(row_offset, 5, f"=Shares!{col_letter}{last_row_idx}")
    ws_dash.write_formula(row_offset, 6, f"=ROC23!{col_letter}{last_row_idx}", pct_format)

workbook.close()
print("trendstrategy_formulas_equity25-1.xlsx generated.")
