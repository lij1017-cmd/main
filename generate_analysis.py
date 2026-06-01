
import pandas as pd
import numpy as np
from backtest_stress import BacktesterVol, clean_data, calculate_metrics_dual
import json

def run_analysis():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterVol(prices, volumes, code_to_name)

    scenarios = {
        'Baseline': {
            'sma_period': 303, 'roc_period': 14,
            'stop_loss_type': 'fixed', 'stop_loss_val': 0.0999,
            'use_market_filter': False, 'use_breadth_weight': False
        },
        'Scenario A': {
            'sma_period': 303, 'roc_period': 14,
            'stop_loss_type': 'fixed', 'stop_loss_val': 0.10,
            'use_market_filter': True, 'use_breadth_weight': False
        },
        'Scenario B': {
            'sma_period': 303, 'roc_period': 14,
            'stop_loss_type': 'vol', 'vol_multiplier': 2.7,
            'use_market_filter': False, 'use_breadth_weight': False
        },
        'Scenario C': {
            'sma_period': 303, 'roc_period': 14,
            'stop_loss_type': 'vol', 'vol_multiplier': 2.7,
            'use_market_filter': True, 'use_breadth_weight': True
        }
    }

    costs = {
        'Standard': {'sl_slippage': 0.0, 'filter_slippage': 0.0},
        'Conservative': {'sl_slippage': 0.003, 'filter_slippage': 0.0},
        'Extreme': {'sl_slippage': 0.0, 'filter_slippage': 0.005}
    }

    results = {}

    for s_name, s_params in scenarios.items():
        results[s_name] = {}
        for c_name, c_params in costs.items():
            print(f"Running {s_name} under {c_name} cost...")
            res = bt.run(**s_params, **c_params)
            metrics = calculate_metrics_dual(res[0], 30000000, 150000000)

            # Extract relevant metrics
            results[s_name][c_name] = {
                'CAGR': metrics['Trading CAGR'],
                'MaxDD': metrics['Standard MaxDD'],
                'Calmar': metrics['Trading Calmar'],
                'Return2022': metrics['Yearly Performance'].loc[2022, '年度報酬率'] if 2022 in metrics['Yearly Performance'].index else 0.0
            }

    with open('analysis_results.json', 'w') as f:
        json.dump(results, f, indent=4)

    print("Analysis complete. Results saved to analysis_results.json")

if __name__ == "__main__":
    run_analysis()
