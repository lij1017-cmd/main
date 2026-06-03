
import pandas as pd
import numpy as np
from backtest_adj4 import BacktesterVol, clean_data, calculate_metrics_dual
import xlsxwriter
import json

def export_to_excel_final(equity_df, trades_df, trades2_df, daily_df, metrics, filename):
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        trades_df.to_excel(writer, sheet_name='Trades', index=False)
        trades2_df.to_excel(writer, sheet_name='Trades2', index=False)
        equity_df.to_excel(writer, sheet_name='Equity_Curve', index=False)
        daily_df.to_excel(writer, sheet_name='Daily', index=False)

        # Summary Sheet
        summary_data = [
            ['策略指標 (全期間)', '數值'],
            ['Trading CAGR (30M)', f"{metrics['Trading CAGR']:.2%}"],
            ['Authorized CAGR (150M)', f"{metrics['Authorized CAGR']:.2%}"],
            ['Standard MaxDD', f"{metrics['Standard MaxDD']:.2%}"],
            ['Fixed Base MaxDD (150M)', f"{metrics['Fixed Base MaxDD']:.2%}"],
            ['Trading Calmar', f"{metrics['Trading Calmar']:.2f}"],
            ['', ''],
            ['年度績效 (實戰模式)', '年度報酬率', '年度損益 (TWD)', '年度MDD (150M基準)']
        ]

        # Calculate Yearly MDD (150M base)
        equity_df['Year'] = equity_df['日期'].dt.year
        yearly_groups = equity_df.groupby('Year')

        for year, row in metrics['Yearly Performance'].iterrows():
            # Get MDD for this specific year from fixed base column
            year_mdd_fixed = equity_df[equity_df['Year'] == year]['固定基準回撤'].min()
            summary_data.append([year, f"{row['年度報酬率']:.2%}", f"{row['年度損益']:,.0f}", f"{year_mdd_fixed:.2%}"])

        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False, header=False)

def run_analysis_adj4():
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterVol(prices, volumes, code_to_name)

    params_c = {
        'sma_period': 303, 'roc_period': 14,
        'stop_loss_type': 'vol', 'vol_multiplier': 2.7,
        'use_market_filter': True, 'breadth_window': 290, 'use_breadth_weight': True
    }

    # Run Full Period for Excel
    print("Generating full period deliverables...")
    eq, t, t2, d = bt.run(**params_c)
    metrics = calculate_metrics_dual(eq, 30000000, 150000000)
    export_to_excel_final(eq, t, t2, d, metrics, 'equityV-adj4.xlsx')

    # Gather other metrics for MD report (optional: could just use values from discussion)
    # But let's verify IS/OOS
    res_is = bt.run(**params_c, start_date='2019-01-01', end_date='2023-12-31')
    m_is = calculate_metrics_dual(res_is[0], 30000000, 150000000)

    res_oos = bt.run(**params_c, start_date='2024-01-01', end_date='2025-12-31')
    m_oos = calculate_metrics_dual(res_oos[0], 30000000, 150000000)

    print(f"IS CAGR: {m_is['Trading CAGR']:.2%}")
    print(f"OOS CAGR: {m_oos['Trading CAGR']:.2%}")

    print("Adj4 deliverables generation complete.")

if __name__ == "__main__":
    run_analysis_adj4()
