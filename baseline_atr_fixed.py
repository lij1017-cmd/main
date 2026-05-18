import pandas as pd
from backtest_atr import clean_data, BacktesterATR, calculate_metrics

def run_baseline():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterATR(prices, volumes, code_to_name)

    # 參數來自 reproduce_equityV_filter.md (方案 B)
    # 使用 'fixed' 停損來確認新引擎的基準線 (固定預算模式)
    equity, trades, holdings, trades2, daily = bt.run(
        sma_period=303,
        roc_period=14,
        stop_loss_type='fixed',
        stop_loss_val=0.0999,
        rebalance_interval=9,
        use_market_filter=True,
        breadth_threshold=0.42,
        mkt_sma_window=14,
        breadth_window=290
    )

    cagr, mdd, calmar, total_ret = calculate_metrics(equity)
    print(f"全期間績效 (固定預算 10M, 固定停損 9.99%):")
    print(f"CAGR: {cagr:.2%}")
    print(f"MaxDD: {mdd:.2%}")
    print(f"Calmar: {calmar:.2f}")

    # 年度績效
    equity['Year'] = equity['日期'].dt.year
    for year, group in equity.groupby('Year'):
        y_cagr, y_mdd, y_calmar, y_ret = calculate_metrics(group)
        print(f"{year} - CAGR: {y_cagr:.2%}, MaxDD: {y_mdd:.2%}, Calmar: {y_calmar:.2f}, Return: {y_ret:.2%}")

if __name__ == "__main__":
    run_baseline()
