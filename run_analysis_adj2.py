
import pandas as pd
import numpy as np
from backtest_adj2 import BacktesterVol, clean_data, calculate_metrics_dual
import json
import xlsxwriter

def export_to_excel(equity_df, trades_df, trades2_df, daily_df, metrics, filename):
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        trades_df.to_excel(writer, sheet_name='Trades', index=False)
        trades2_df.to_excel(writer, sheet_name='Trades2', index=False)
        equity_df.to_excel(writer, sheet_name='Equity_Curve', index=False)
        # Equity_Hold placeholder (current holdings)
        # In this simple backtest, we don't track historical hold snapshots every day in a separate sheet usually,
        # but we can provide the last day's slots.

        daily_df.to_excel(writer, sheet_name='Daily', index=False)

        # Summary Sheet
        summary_data = [
            ['指標', '數值'],
            ['Trading CAGR (30M)', f"{metrics['Trading CAGR']:.2%}"],
            ['Authorized CAGR (150M)', f"{metrics['Authorized CAGR']:.2%}"],
            ['Standard MaxDD', f"{metrics['Standard MaxDD']:.2%}"],
            ['Fixed Base MaxDD', f"{metrics['Fixed Base MaxDD']:.2%}"],
            ['Trading Calmar', f"{metrics['Trading Calmar']:.2f}"]
        ]

        # Add Yearly Performance
        summary_data.append(['', ''])
        summary_data.append(['年度', '年度報酬率', '年度損益'])
        for year, row in metrics['Yearly Performance'].iterrows():
            summary_data.append([year, f"{row['年度報酬率']:.2%}", f"{row['年度損益']:,.0f}"])

        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False, header=False)

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
            equity_curve, trades, trades2, daily = bt.run(**s_params, **c_params)
            metrics = calculate_metrics_dual(equity_curve, 30000000, 150000000)

            # Extract relevant metrics
            results[s_name][c_name] = {
                'CAGR': metrics['Trading CAGR'],
                'MaxDD': metrics['Standard MaxDD'],
                'Calmar': metrics['Trading Calmar'],
                'Return2022': metrics['Yearly Performance'].loc[2022, '年度報酬率'] if 2022 in metrics['Yearly Performance'].index else 0.0
            }

            # Export Scenario C (Standard) as requested
            if s_name == 'Scenario C' and c_name == 'Standard':
                print(f"Exporting Scenario C (Standard) to equityV-adj2.xlsx...")
                export_to_excel(equity_curve, trades, trades2, daily, metrics, 'equityV-adj2.xlsx')

    with open('analysis_results.json', 'w') as f:
        json.dump(results, f, indent=4)

    print("Analysis complete. Results saved to analysis_results.json")

if __name__ == "__main__":
    run_analysis()
