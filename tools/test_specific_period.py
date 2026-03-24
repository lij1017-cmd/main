import pandas as pd
import numpy as np
from run_wfa import Backtester, clean_data, calculate_metrics

def main():
    # 參數設定 (先前針對 2020-2024 優化出且 Reb=9 的結果)
    SMA_PERIOD = 54
    ROC_PERIOD = 52
    STOP_LOSS_PCT = 0.09
    REBALANCE = 9
    INITIAL_CAPITAL = 30000000
    DATA_FILE = '個股合-1.xlsx'

    prices, code_to_name = clean_data(DATA_FILE)
    bt = Backtester(prices, code_to_name, INITIAL_CAPITAL)

    # 指定回測區間
    start_date = '2019-01-02'
    end_date = '2020-05-31'

    print(f"正在執行指定期間回測 ({start_date} 至 {end_date})...")
    print(f"參數: SMA={SMA_PERIOD}, ROC={ROC_PERIOD}, SL={STOP_LOSS_PCT*100:.1f}%, Reb={REBALANCE}")

    eq_df, trades = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE, start_date, end_date)
    cagr, mdd, calmar = calculate_metrics(eq_df)

    print("\n=== 指定期間回測績效 (2019-2020) ===")
    print(f"年化報酬率 (CAGR): {cagr:.2%}")
    print(f"最大回撤 (MaxDD): {mdd:.2%}")
    print(f"卡瑪比率 (Calmar Ratio): {calmar:.2f}")
    print(f"總交易次數: {trades}")
    print(f"期末淨值: {eq_df['權益'].iloc[-1]:,.0f}")

if __name__ == "__main__":
    main()
