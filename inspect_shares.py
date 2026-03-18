import pandas as pd

# Load the trades and equity hold data
excel_file = 'trendstrategy_results_equity2025新-1.xlsx'
try:
    trades_df = pd.read_excel(excel_file, sheet_name='Trades')
    hold_df = pd.read_excel(excel_file, sheet_name='Equity_Hold')

    # Filter for the specific period
    start_date = '2019-06-03'
    end_date = '2019-06-10'

    # Get daily status from hold_df
    mask = (hold_df['Date'] >= start_date) & (hold_df['Date'] <= end_date)
    period_holdings = hold_df[mask]

    # Get trades in/around this period to determine quantity changes
    # Trades are logged on the signal date. Execution is T+1.
    # We need to know the shares at each date.

    print("### Daily Details (2019/06/03 - 2019/06/10) ###")
    for idx, row in period_holdings.iterrows():
        date_str = row['Date'].strftime('%Y/%m/%d')
        holdings = row['Holdings']

        # Find trades on this specific date (Execution Date)
        # Note: In our current log, '日期' is the Signal Date.
        # So execution at T+1 means if Signal is T, execution price and status change happens at T+1 close.

        # Let's find trades where T+1 == current_date
        # We need the full trades list for this.

        print(f"日期: {date_str}")
        print(f"持股內容: {holdings}")

        # Look for trades executed on this date
        # Signal T -> Execution T+1.
        # We need to find signals from T-1.
        prev_date = trades_df['日期'].unique()
        # This is getting complex. Let's just list the trades and match them manually or via script.

    # Let's get the specific trade details for these dates
    # We want trades that changed the portfolio on these dates.
    # 3227 was sold on 6/5 (Signal 6/4).
    # 2633 was bought on 6/6 (Signal 6/5).

    # Let's get the shares from the trades log
    relevant_stocks = ['2207', '3227', '2633']
    for s in relevant_stocks:
        s_trades = trades_df[trades_df['股票代號'].astype(str) == s]
        print(f"\nTrades for {s}:")
        print(s_trades[['日期', '狀態', '價格', '股數', '原因']])

except Exception as e:
    print(f"Error: {e}")
