import pandas as pd
import numpy as np
from run_backtest_equity2025新_動態版V1 import Backtester, clean_data, calculate_metrics

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

    # 回測區間 2019-2024
    start_date = '2019-01-02'
    end_date = '2024-12-31'

    # 執行回測 (動態版 V1 邏輯)
    # 我們需要過濾日期，因為 Backtester.run 會跑完整價格資料
    # 但為了計算精確的區間績效，我們在 run 之後切片，或者修改 run 邏輯。
    # 考慮到 Backtester 的實作，它會從 start_idx 開始跑。

    eq_df, trades, hold, trades2, daily = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE)

    # 切片目標區間
    mask = (eq_df['日期'] >= pd.to_datetime(start_date)) & (eq_df['日期'] <= pd.to_datetime(end_date))
    eq_period = eq_df[mask]

    cagr, mdd, calmar, ret = calculate_metrics(eq_period)

    print(f"\n=== 動態版 V1 測試結果 (2019-2024) ===")
    print(f"參數: SMA={SMA_PERIOD}, ROC={ROC_PERIOD}, SL={STOP_LOSS_PCT*100:.1f}%, Reb={REBALANCE}")
    print(f"年化報酬率 (CAGR): {cagr:.2%}")
    print(f"最大回撤 (MaxDD): {mdd:.2%}")
    print(f"卡瑪比率 (Calmar Ratio): {calmar:.2f}")
    print(f"期末淨值: {eq_period['權益'].iloc[-1]:,.0f}")
    print(f"總交易筆數: {len(trades[(trades['訊號日期'] >= pd.to_datetime(start_date)) & (trades['訊號日期'] <= pd.to_datetime(end_date))])}")

if __name__ == "__main__":
    main()
