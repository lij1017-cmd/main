import pandas as pd
import numpy as np

def get_metrics_adj1(filepath):
    df = pd.read_excel(filepath, sheet_name='Summary')
    metrics = {}
    # Based on previous output
    # 3: 最初投入資金 CAGR
    # 5: 標準 MDD (對峰值)
    # 7: Trading Calmar Ratio
    metrics['CAGR'] = df.iloc[3, 1]
    metrics['MaxDD'] = df.iloc[5, 1]
    metrics['Calmar'] = df.iloc[7, 1]

    # Extract yearly returns
    yearly_returns = {}
    for i in range(10, len(df), 2):
        label = str(df.iloc[i, 0])
        if '年度報酬率' in label:
            year = label.split(' ')[0]
            yearly_returns[year] = df.iloc[i, 1]
    metrics['Yearly'] = yearly_returns
    return metrics

def get_metrics_orig(filepath):
    # The original file might not have a summary sheet, I might need to calculate it from Equity Curve
    # Let's check if there's an Equity_Curve sheet in trendstrategy_results_equityV.xlsx
    try:
        df_eq = pd.read_excel(filepath, sheet_name='Equity_Curve')
        # Assuming columns: '日期', '總權益' or similar
        # Based on backtest_vol.py, it should have '權益' and '回撤(Drawdown)'
        # Let's check columns first
        # print(df_eq.columns)

        equity = df_eq['權益']
        dates = pd.to_datetime(df_eq['日期'])

        days = (dates.iloc[-1] - dates.iloc[0]).days
        years = days / 365.25
        total_return = (equity.iloc[-1] / equity.iloc[0]) - 1
        cagr = (1 + total_return)**(1/years) - 1

        # MaxDD
        peak = equity.cummax()
        dd = (equity - peak) / peak
        max_dd = dd.min()

        calmar = cagr / abs(max_dd)

        # Yearly
        df_eq['Year'] = dates.dt.year
        yearly = df_eq.groupby('Year').last()['權益']
        yearly_prev = df_eq.groupby('Year').first()['權益']
        # This is not quite right for yearly return if we want to reset every year
        # Let's just do simple end-of-year / start-of-year

        years_list = df_eq['Year'].unique()
        yearly_ret = {}
        for y in years_list:
            year_data = df_eq[df_eq['Year'] == y]
            start_val = year_data['權益'].iloc[0]
            end_val = year_data['權益'].iloc[-1]
            yearly_ret[str(y)] = (end_val / start_val) - 1

        return {
            'CAGR': cagr,
            'MaxDD': max_dd,
            'Calmar': calmar,
            'Yearly': yearly_ret
        }
    except Exception as e:
        print(f"Error processing original: {e}")
        return None

adj1 = get_metrics_adj1('equityV-adj1.xlsx')
orig = get_metrics_orig('trendstrategy_results_equityV.xlsx')

print("--- Comparison ---")
print(f"{'Metric':<15} {'Original':<15} {'Adj1':<15}")
print(f"{'CAGR':<15} {orig['CAGR']:.2%} {adj1['CAGR']:.2%}")
print(f"{'MaxDD':<15} {orig['MaxDD']:.2%} {adj1['MaxDD']:.2%}")
print(f"{'Calmar':<15} {orig['Calmar']:.2f} {adj1['Calmar']:.2f}")

print("\n--- Yearly Returns ---")
years = sorted(list(set(adj1['Yearly'].keys()) | set(orig['Yearly'].keys())))
for y in years:
    o_ret = orig['Yearly'].get(y, np.nan)
    a_ret = adj1['Yearly'].get(y, np.nan)
    print(f"{y}: Original: {o_ret:.2%}, Adj1: {a_ret:.2%}")
