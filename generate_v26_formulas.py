import pandas as pd
import xlsxwriter
import numpy as np

# Load raw data to get dimensions and headers
df_raw = pd.read_excel('個股合-1.xlsx', header=None)
num_rows = df_raw.shape[0]
num_cols = df_raw.shape[1]
stock_cols = range(2, num_cols)
last_col_letter = xlsxwriter.utility.xl_col_to_name(num_cols - 1)

output_file = 'trendstrategy_formulas_equity26.xlsx'
workbook = xlsxwriter.Workbook(output_file, {'nan_inf_to_errors': True})

# Formats
header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1, 'align': 'center'})
date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
pct_format = workbook.add_format({'num_format': '0.00%'})
price_format = workbook.add_format({'num_format': '#,##0.00'})
num_format = workbook.add_format({'num_format': '#,##0'})

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
        if excel_row >= 66: # 64 days including current: 3 to 66 is 64 rows
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
        formula = f"=IF(AND(Prices!{col_letter}{excel_row}>SMA64!{col_letter}{excel_row}, ISNUMBER(ROC23!{col_letter}{excel_row}), ROC23!{col_letter}{excel_row}>0), ROC23!{col_letter}{excel_row}, -1)"
        ws_elig.write_formula(r, c, formula, pct_format)

# 5. Rank Sheet
ws_rank = workbook.add_worksheet('Rank')
write_basic_structure(ws_rank)
for r in range(2, num_rows):
    excel_row = r + 1
    for c in stock_cols:
        col_letter = xlsxwriter.utility.xl_col_to_name(c)
        # Rank within row among all stocks
        formula = f"=IF(Eligible!{col_letter}{excel_row}>0, RANK(Eligible!{col_letter}{excel_row}, Eligible!$C{excel_row}:${last_col_letter}{excel_row}), \"\")"
        ws_rank.write_formula(r, c, formula)

# 6. RebalanceDay Sheet
ws_rebal = workbook.add_worksheet('RebalanceDay')
write_basic_structure(ws_rebal)
start_data_row = 66
for r in range(2, num_rows):
    excel_row = r + 1
    # Every 5 days starting from start_data_row
    val = f"=IF({excel_row}>={start_data_row}, MOD({excel_row}-{start_data_row}, 5)=0, FALSE)"
    ws_rebal.write_formula(r, 2, val)

# Logic Sheets
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

# New Helper Sheets for Performance
ws_entry_row = workbook.add_worksheet('EntryRow')
write_basic_structure(ws_entry_row)

ws_trade_idx = workbook.add_worksheet('TradeIndex')
write_basic_structure(ws_trade_idx)

for r in range(2, num_rows):
    excel_row = r + 1
    for c in stock_cols:
        col_letter = xlsxwriter.utility.xl_col_to_name(c)

        if excel_row < start_data_row:
            ws_status.write(r, c, "Cash")
            ws_shares.write(r, c, 0)
            ws_maxp.write(r, c, 0)
            ws_slt.write(r, c, False)
            ws_decide.write(r, c, "")
            ws_entry_row.write(r, c, 0)
            ws_trade_idx.write(r, c, "")
        else:
            # Status: Depends on previous day's decision
            # Status[T] = status DURING day T (before EOD decision)
            # Actually, let's say Status[T] is status after T-1 execution.
            status_formula = f'=IF(Decision!{col_letter}{excel_row-1}="BUY", "Holding", IF(Decision!{col_letter}{excel_row-1}="SELL", "Cash", Status!{col_letter}{excel_row-1}))'
            ws_status.write_formula(r, c, status_formula)

            # Shares
            # If we bought yesterday EOD (at T-1 signal, executed at T close), we have shares at end of T.
            # But "T signal, T+1 execute".
            # Decision[T-1] -> Execute at T close.
            # So Shares[T] reflects the result of Decision[T-1].
            shares_formula = f'=IF(Decision!{col_letter}{excel_row-1}="BUY", 10000000/Prices!{col_letter}{excel_row}, IF(Decision!{col_letter}{excel_row-1}="SELL", 0, Shares!{col_letter}{excel_row-1}))'
            ws_shares.write_formula(r, c, shares_formula)

            # MaxPrice
            # MaxPrice during holding period.
            maxp_formula = f'=IF(Status!{col_letter}{excel_row}="Holding", IF(Decision!{col_letter}{excel_row-1}="BUY", Prices!{col_letter}{excel_row}, MAX(Prices!{col_letter}{excel_row}, MaxPrice!{col_letter}{excel_row-1})), 0)'
            ws_maxp.write_formula(r, c, maxp_formula)

            # SL Trigger (Check at T close)
            # Stop loss logic: If Price[T] < MaxPrice[T] * 0.91
            slt_formula = f'=AND(Status!{col_letter}{excel_row}="Holding", Prices!{col_letter}{excel_row}<MaxPrice!{col_letter}{excel_row}*0.91)'
            ws_slt.write_formula(r, c, slt_formula)

            # Decision at T close
            # Signal at T, execute at T+1.
            # If Cash: if rebalance day and top 3 -> BUY.
            # If Holding: if rebalance day and not top 3 -> SELL.
            # If Holding: if SL_Trigger -> SELL.
            decide_formula = f'=IF(Status!{col_letter}{excel_row}="Cash", IF(AND(RebalanceDay!$C{excel_row}, Rank!{col_letter}{excel_row}<>"", Rank!{col_letter}{excel_row}<=3), "BUY", ""), IF(OR(AND(RebalanceDay!$C{excel_row}, OR(Rank!{col_letter}{excel_row}="", Rank!{col_letter}{excel_row}>3)), SL_Trigger!{col_letter}{excel_row}), "SELL", "Hold"))'
            ws_decide.write_formula(r, c, decide_formula)

            # EntryRow tracking (execution row)
            entry_row_formula = f'=IF(Decision!{col_letter}{excel_row-1}="BUY", {excel_row}, IF(Status!{col_letter}{excel_row}="Holding", EntryRow!{col_letter}{excel_row-1}, 0))'
            ws_entry_row.write_formula(r, c, entry_row_formula)

            # TradeIndex (for SELL events)
            trade_idx_formula = f'=IF(Decision!{col_letter}{excel_row}="SELL", {excel_row}*10000+{c}, "")'
            ws_trade_idx.write_formula(r, c, trade_idx_formula)

# Equity Sheet
ws_equity = workbook.add_worksheet('Equity')
ws_equity.write(0, 0, "日期", header_format)
ws_equity.write(0, 1, "現金", header_format)
ws_equity.write(0, 2, "市值", header_format)
ws_equity.write(0, 3, "總資產", header_format)
ws_equity.write(0, 4, "最高資產", header_format)
ws_equity.write(0, 5, "回撤", header_format)

initial_cash = 30000000
for r in range(2, num_rows):
    excel_row = r + 1
    ws_equity.write(r, 0, df_raw.iloc[r, 1], date_format)

    if excel_row == 3:
        ws_equity.write(r, 1, initial_cash, num_format)
    else:
        # Cash[T] = Cash[T-1] + SellProceeds[T] - BuyCost[T]
        # Actions at T are determined by Decision[T-1]
        # SellProceeds[T] = SUM(Shares[T-1] * Price[T]) where Decision[T-1]=="SELL"
        # BuyCost[T] = SUM(10,000,000) where Decision[T-1]=="BUY"
        cash_f = f"=B{excel_row-1} + SUMPRODUCT((Decision!$C{excel_row-1}:${last_col_letter}{excel_row-1}=\"SELL\")*Shares!$C{excel_row-1}:${last_col_letter}{excel_row-1}*Prices!$C{excel_row}:${last_col_letter}{excel_row}) - SUMPRODUCT((Decision!$C{excel_row-1}:${last_col_letter}{excel_row-1}=\"BUY\")*10000000)"
        ws_equity.write_formula(r, 1, cash_f, num_format)

    # Market Value = SUMPRODUCT(Shares[T], Prices[T])
    mv_f = f"=SUMPRODUCT(Shares!$C{excel_row}:${last_col_letter}{excel_row}*Prices!$C{excel_row}:${last_col_letter}{excel_row})"
    ws_equity.write_formula(r, 2, mv_f, num_format)
    ws_equity.write_formula(r, 3, f"=B{excel_row}+C{excel_row}", num_format)

    if excel_row == 3:
        ws_equity.write_formula(r, 4, f"=D{excel_row}", num_format)
    else:
        ws_equity.write_formula(r, 4, f"=MAX(D$3:D{excel_row})", num_format)

    ws_equity.write_formula(r, 5, f"=IF(E{excel_row}<>0, (D{excel_row}-E{excel_row})/E{excel_row}, 0)", pct_format)

# Performance Sheet
ws_perf = workbook.add_worksheet('Performance')
perf_headers = ["進場日期", "出場日期", "進場價格", "出場價格", "持有天數", "報酬率 (%)", "累積資金曲線"]
for i, h in enumerate(perf_headers):
    ws_perf.write(0, i, h, header_format)

# Max number of trades to extract
max_trades = 1000
for i in range(1, max_trades + 1):
    excel_row = i + 1
    trade_idx_range = f"TradeIndex!$C$3:${last_col_letter}${num_rows}"
    # Col H: Hidden Trade Index
    ws_perf.write_formula(i, 7, f"=IFERROR(SMALL({trade_idx_range}, {i}), \"\")")

    idx_cell = f"H{excel_row}"
    # Col I: Exit Row = INT(idx/10000)
    ws_perf.write_formula(i, 8, f"=IF({idx_cell}<>\"\", INT({idx_cell}/10000), \"\")")
    # Col J: Col Index = MOD(idx, 10000)
    ws_perf.write_formula(i, 9, f"=IF({idx_cell}<>\"\", MOD({idx_cell}, 10000), \"\")")
    # Col K: Entry Row
    entry_row_f = f'=IF(I{excel_row}<>"", INDEX(EntryRow!$A$1:${last_col_letter}${num_rows}, I{excel_row}, J{excel_row}+1), "")'
    ws_perf.write_formula(i, 10, f"=IF({idx_cell}<>\"\", {entry_row_f}, \"\")")

    # Visible Columns
    # Entry Date
    ws_perf.write_formula(i, 0, f'=IF(K{excel_row}<>"", INDEX(Prices!$B$1:$B${num_rows}, K{excel_row}), "")', date_format)
    # Exit Date
    ws_perf.write_formula(i, 1, f'=IF(I{excel_row}<>"", INDEX(Prices!$B$1:$B${num_rows}, I{excel_row}), "")', date_format)
    # Entry Price
    ws_perf.write_formula(i, 2, f'=IF(K{excel_row}<>"", INDEX(Prices!$A$1:${last_col_letter}${num_rows}, K{excel_row}, J{excel_row}+1), "")', price_format)
    # Exit Price
    ws_perf.write_formula(i, 3, f'=IF(I{excel_row}<>"", INDEX(Prices!$A$1:${last_col_letter}${num_rows}, I{excel_row}, J{excel_row}+1), "")', price_format)
    # Holding Days
    ws_perf.write_formula(i, 4, f'=IF(B{excel_row}<>"", B{excel_row}-A{excel_row}, "")')
    # Return %
    ws_perf.write_formula(i, 5, f'=IF(C{excel_row}<>"", (D{excel_row}-C{excel_row})/C{excel_row}, "")', pct_format)
    # Accumulated Equity
    ws_perf.write_formula(i, 6, f'=IF(I{excel_row}<>"", INDEX(Equity!$D$1:$D${num_rows}, I{excel_row}), "")', num_format)

# Summary Stats Sheet
ws_stats = workbook.add_worksheet('總績效統計')
# Use full column references for dynamic range
stats = [
    ("總交易次數", "=COUNT(Performance!B:B)"),
    ("勝率", "=IF(COUNT(Performance!F:F)>0, COUNTIF(Performance!F:F, \">0\")/COUNT(Performance!F:F), 0)"),
    ("平均報酬率", "=IFERROR(AVERAGE(Performance!F:F), 0)"),
    ("最大回撤", "=MIN(Equity!F:F)"),
    ("年化報酬率", f"=IF(COUNT(Equity!A:A)>1, (INDEX(Equity!D:D, MATCH(9.99999999999999E+307, Equity!D:D))/{initial_cash})^(252/COUNT(Equity!A:A)) - 1, 0)")
]
for i, (label, formula) in enumerate(stats):
    ws_stats.write(i, 0, label, header_format)
    if "%" in label or "率" in label or "回撤" in label:
        ws_stats.write_formula(i, 1, formula, pct_format)
    else:
        ws_stats.write_formula(i, 1, formula)

workbook.close()
print(f"{output_file} generated successfully.")
