import pandas as pd
import numpy as np
from backtest_engine import Backtester, clean_data, calculate_metrics

def main():
    # 參數設定
    SMA_PERIOD = 54
    ROC_PERIOD = 52
    STOP_LOSS_PCT = 0.09
    REBALANCE = 9
    INITIAL_CAPITAL = 30000000
    DATA_FILE = '個股合-1.xlsx'

    prices, code_to_name = clean_data(DATA_FILE)
    bt = Backtester(prices, code_to_name, INITIAL_CAPITAL)

    # 回測區間 2024.11.01 - 2025.12.31
    start_date = '2024-11-01'
    end_date = '2025-12-31'

    print(f"正在執行測試 (2024 Nov - 2025): {start_date} 至 {end_date}")
    print(f"參數: SMA={SMA_PERIOD}, ROC={ROC_PERIOD}, SL={STOP_LOSS_PCT*100:.1f}%, Reb={REBALANCE}")

    eq_df, t_log, h_log, t2_log, d_log = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE, start_date, end_date)
    cagr, mdd, calmar = calculate_metrics(eq_df)

    print(f"\n=== 動態版 V1 測試結果 (2024.11-2025.12) ===")
    print(f"年化報酬率 (CAGR): {cagr:.2%}")
    print(f"最大回撤 (MaxDD): {mdd:.2%}")
    print(f"卡瑪比率 (Calmar Ratio): {calmar:.2f}")
    print(f"總交易筆數: {len(t_log)}")
    print(f"期末淨值: {eq_df['權益'].iloc[-1]:,.0f}")

if __name__ == "__main__":
    main()
