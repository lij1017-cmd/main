import pandas as pd
import numpy as np
from backtest_breadth import clean_data, BacktesterBreadth, calculate_metrics

def main():
    DATA_FILE = '樣本集-1.xlsx'
    INITIAL_CAPITAL = 30000000
    SMA_PERIOD = 303
    ROC_PERIOD = 14
    STOP_LOSS_PCT = 0.0999
    REBALANCE = 9

    print(f"Loading data for optimization...")
    prices, volumes, code_to_name = clean_data(DATA_FILE)
    bt = BacktesterBreadth(prices, volumes, code_to_name, initial_capital=INITIAL_CAPITAL)

    results = []
    # Search threshold from 10% to 60% with 1% steps
    thresholds = np.linspace(0.10, 0.60, 51)

    print(f"Starting brute-force search for optimal breadth threshold...")
    for t in thresholds:
        eq, _, _, _, _ = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE, use_market_filter=True, breadth_threshold=t)

        # Calculate full period metrics (2019-2025)
        mask = (eq['日期'] >= '2019-01-01') & (eq['日期'] <= '2025-12-31')
        res_p = eq[mask]
        cagr, mdd, calmar, _ = calculate_metrics(res_p)

        results.append({
            'Threshold': t,
            'CAGR': cagr,
            'MaxDD': mdd,
            'Calmar': calmar
        })
        print(f"Threshold: {t:.2%}, CAGR: {cagr:.2%}, MaxDD: {mdd:.2%}, Calmar: {calmar:.2f}")

    df_results = pd.DataFrame(results)
    best_row = df_results.loc[df_results['Calmar'].idxmax()]

    print("\nOptimization Finished!")
    print(f"Best Threshold: {best_row['Threshold']:.2%}")
    print(f"Max Calmar: {best_row['Calmar']:.2f}")
    print(f"Corresponding CAGR: {best_row['CAGR']:.2%}, MaxDD: {best_row['MaxDD']:.2%}")

    df_results.to_excel('breadth_optimization_results.xlsx', index=False)

if __name__ == "__main__":
    main()
