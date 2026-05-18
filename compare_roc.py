import pandas as pd
from backtest_atr import clean_data, BacktesterATR, calculate_metrics

def compare():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterATR(prices, volumes, code_to_name)

    # 方案 A: 使用原 ROC 14
    print("Testing ROC 14 with optimized ATR...")
    eq14, _, _, _, _ = bt.run(
        sma_period=303, roc_period=14, stop_loss_type='atr',
        atr_period=20, atr_multiplier=4.5,
        rebalance_interval=9, use_market_filter=True, breadth_threshold=0.42, mkt_sma_window=14
    )
    cagr14, mdd14, calmar14, ret14 = calculate_metrics(eq14)
    y22_14 = calculate_metrics(eq14[eq14['日期'].dt.year == 2022])[3]

    # 方案 B: 使用優化後的 ROC 10
    print("Testing ROC 10 with optimized ATR...")
    eq10, _, _, _, _ = bt.run(
        sma_period=303, roc_period=10, stop_loss_type='atr',
        atr_period=15, atr_multiplier=4.3,
        rebalance_interval=9, use_market_filter=True, breadth_threshold=0.42, mkt_sma_window=14
    )
    cagr10, mdd10, calmar10, ret10 = calculate_metrics(eq10)
    y22_10 = calculate_metrics(eq10[eq10['日期'].dt.year == 2022])[3]

    print("\n--- 比較結果 ---")
    print(f"ROC 14: CAGR={cagr14:.2%}, Calmar={calmar14:.2f}, 2022 Ret={y22_14:.2%}")
    print(f"ROC 10: CAGR={cagr10:.2%}, Calmar={calmar10:.2f}, 2022 Ret={y22_10:.2%}")

if __name__ == "__main__":
    compare()
