import pandas as pd
import numpy as np
from run_wfa import Backtester, clean_data, calculate_metrics

def main():
    # 參數設定 (ACO 最佳化結果)
    SMA_PERIOD = 93
    ROC_PERIOD = 88
    STOP_LOSS_PCT = 0.08
    REBALANCE = 6
    INITIAL_CAPITAL = 30000000
    DATA_FILE = '個股合-1.xlsx'

    prices, code_to_name = clean_data(DATA_FILE)
    bt = Backtester(prices, code_to_name, INITIAL_CAPITAL)

    # 全期間範圍 (2019-01-02 到 2025-12-31)
    start_date = prices.index[0]
    end_date = prices.index[-1]

    print(f"正在執行全期間回測 ({start_date.date()} 至 {end_date.date()})...")
    print(f"參數: SMA={SMA_PERIOD}, ROC={ROC_PERIOD}, SL={STOP_LOSS_PCT*100:.1f}%, Rebalance={REBALANCE}")

    eq_df, trades = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE, start_date, end_date)
    cagr, mdd, calmar = calculate_metrics(eq_df)

    print("\n=== 全期間回測績效 (2019-2025) ===")
    print(f"年化報酬率 (CAGR): {cagr:.2%}")
    print(f"最大回撤 (MaxDD): {mdd:.2%}")
    print(f"卡瑪比率 (Calmar Ratio): {calmar:.2f}")
    print(f"總交易次數: {trades}")
    print(f"期末淨值: {eq_df['權益'].iloc[-1]:,.0f}")

if __name__ == "__main__":
    main()
