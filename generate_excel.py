import pandas as pd
from backtest_equity2MA import clean_data, Backtester, calculate_metrics

prices, volumes, names = clean_data('樣本集-1.xlsx')
bt = Backtester(prices, volumes, names)

# s1, s2, r, sv, rb, st = 170, 290, 14, 50, 10, 'ma'
s1, s2, r, sv, rb, st = 150, 345, 17, 15, 9, 'ma'

eq, trades, hold, trades2, daily = bt.run(s1, s2, r, sv, rb, st)

# Summary sheet
cagr, mdd, calmar, ret = calculate_metrics(eq)
win_rate = (trades2['損益'] > 0).mean() if not trades2.empty else 0
summary_df = pd.DataFrame([{
    '策略名稱': 'Asset Class Trend Following (equity-2MA)',
    '參數': f'SMA:({s1},{s2}), ROC:{r}, SL:{st}({sv}), Rebalance:{rb}',
    'CAGR': f'{cagr:.2%}',
    'MaxDD': f'{mdd:.2%}',
    'Calmar Ratio': f'{calmar:.2f}',
    '勝率': f'{win_rate:.2%}',
    '總報酬率': f'{ret:.2%}'
}])

with pd.ExcelWriter('trendstrategy_results_equity2MA.xlsx', engine='xlsxwriter') as writer:
    trades.to_excel(writer, sheet_name='Trades', index=False)
    trades2.to_excel(writer, sheet_name='Trades2', index=False)
    eq.to_excel(writer, sheet_name='Equity_Curve', index=False)
    hold.to_excel(writer, sheet_name='Equity_Hold', index=False)
    daily.to_excel(writer, sheet_name='Daily', index=False)
    summary_df.to_excel(writer, sheet_name='Summary', index=False)

    # Add charts
    workbook = writer.book
    worksheet = writer.sheets['Equity_Curve']
    chart = workbook.add_chart({'type': 'line'})
    chart.add_series({
        'name': 'Equity Curve',
        'categories': ['Equity_Curve', 1, 0, len(eq), 0],
        'values': ['Equity_Curve', 1, 1, len(eq), 1],
    })
    chart.set_title({'name': 'Equity Curve'})
    worksheet.insert_chart('E2', chart)

print("Excel report generated.")
