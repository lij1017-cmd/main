import pandas as pd
import numpy as np
import pickle
from backtester import Backtester, calculate_metrics

def generate():
    prices = pd.read_pickle('prices_cleaned.pkl')
    with open('code_to_name.pkl', 'rb') as f:
        code_to_name = pickle.load(f)
    with open('best_params.pkl', 'rb') as f:
        best_params = pickle.load(f)

    sma_p, roc_p, sl_p = best_params
    bt = Backtester(prices)
    eq, trades, holdings, rebalance_log = bt.run(sma_p, roc_p, sl_p)
    cagr, mdd, calmar, win_rate = calculate_metrics(eq, trades)

    trades_df = rebalance_log.copy()
    trades_df['標的名稱'] = trades_df['股票代號'].map(code_to_name)
    trades_df['最佳參數'] = f"SMA={sma_p}, ROC={roc_p}, SL={sl_p}"
    trades_df['說明'] = trades_df.apply(lambda row: f"選取資產：{row['標的名稱']} ({row['股票代號']})，相對動能值(ROC)：{row['動能值']:.2%}", axis=1)

    rolling_max = eq.cummax()
    drawdown = (eq - rolling_max) / rolling_max
    equity_df = pd.DataFrame({
        'Date': eq.index,
        'Equity': eq.values,
        'Drawdown': drawdown.values
    })

    holdings_df = holdings.copy()
    holdings_df['Holdings_Names'] = holdings_df['Holdings'].apply(lambda x: [code_to_name[a] for a in x])

    summary_df = pd.DataFrame({
        'Metric': ['CAGR', 'MaxDD', 'Calmar Ratio', 'Win Rate', 'SMA', 'ROC', 'StopLoss%'],
        'Value': [f"{cagr:.2%}", f"{mdd:.2%}", f"{calmar:.2f}", f"{win_rate:.2%}", sma_p, roc_p, sl_p]
    })

    with pd.ExcelWriter('trendstrategy_results_equity2024.xlsx') as writer:
        trades_df.to_excel(writer, sheet_name='Trades', index=False)
        equity_df.to_excel(writer, sheet_name='Equity_Curve', index=False)
        holdings_df.to_excel(writer, sheet_name='Equity_Hold', index=False)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

    print("Excel file generated.")

if __name__ == "__main__":
    generate()
