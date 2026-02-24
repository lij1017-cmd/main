import pandas as pd
import numpy as np
import pickle
from backtester import Backtester, calculate_metrics
from data_prep import clean_data

def generate():
    prices, code_to_name = clean_data('個股1.xlsx')
    best_params = (69, 23, 0.09)
    bt = Backtester(prices)
    eq, trades, holdings, action_log = bt.run(*best_params)
    cagr, mdd, calmar, win_rate = calculate_metrics(eq, trades)

    trades_df = action_log.copy()
    trades_df['標的名稱'] = trades_df['股票代號'].map(code_to_name)
    trades_df['最佳參數'] = f"SMA={best_params[0]}, ROC={best_params[1]}, SL={best_params[2]}"
    trades_df['說明'] = trades_df.apply(lambda row: f"資產：{row['標的名稱']} ({row['股票代號']})，動能值(ROC)：{row['動能值']:.2%}", axis=1)

    equity_df = pd.DataFrame({'Date': eq.index, 'Equity': eq.values, 'Drawdown': ((eq - eq.cummax()) / eq.cummax()).values})
    holdings_df = holdings.copy()
    holdings_df['Holdings_Names'] = holdings_df['Holdings'].apply(lambda x: [code_to_name[a] for a in x])
    summary_df = pd.DataFrame({'Metric': ['CAGR', 'MaxDD', 'Calmar Ratio', 'Win Rate', 'SMA', 'ROC', 'StopLoss%'],
                               'Value': [f"{cagr:.2%}", f"{mdd:.2%}", f"{calmar:.2f}", f"{win_rate:.2%}", best_params[0], best_params[1], best_params[2]]})

    with pd.ExcelWriter('trendstrategy_results_equity2024.xlsx') as writer:
        trades_df.to_excel(writer, sheet_name='Trades', index=False)
        equity_df.to_excel(writer, sheet_name='Equity_Curve', index=False)
        holdings_df.to_excel(writer, sheet_name='Equity_Hold', index=False)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
    print("Excel generated.")

if __name__ == "__main__":
    generate()
