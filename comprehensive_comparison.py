import pandas as pd
import numpy as np
from tabulate import tabulate
from wfa_analysis import Backtester, clean_data, calculate_metrics

def main():
    data_file = '個股合-1.xlsx'
    prices, code_to_name = clean_data(data_file)
    bt = Backtester(prices, code_to_name)

    # Define Optimized Configurations
    configs = [
        {"name": "6-day Optimized", "sma": 87, "roc": 54, "sl": 0.09, "reb": 6},
        {"name": "7-day Optimized", "sma": 30, "roc": 52, "sl": 0.075, "reb": 7},
        {"name": "8-day Optimized", "sma": 27, "roc": 98, "sl": 0.075, "reb": 8}
    ]

    # 1. Full Period Comparison (2019-2025)
    full_results = []
    start_all = pd.to_datetime('2019/1/2')
    end_all = pd.to_datetime('2025/12/31')
    for cfg in configs:
        eq, trades, cost = bt.run(cfg["sma"], cfg["roc"], cfg["sl"], start_all, end_all, rebalance_interval=cfg["reb"])
        cagr, mdd, calmar = calculate_metrics(eq)

        full_results.append([
            cfg["name"],
            f"SMA:{cfg['sma']}, ROC:{cfg['roc']}, SL:{cfg['sl']*100:.1f}%",
            f"{cagr:.2%}",
            f"{mdd:.2%}",
            f"{calmar:.2f}",
            trades,
            f"{int(cost):,}"
        ])

    print("### Full Period Comparison (2019-2025)")
    headers = ["Strategy", "Parameters", "CAGR", "MaxDD", "Calmar", "Trades", "Costs (TWD)"]
    print(tabulate(full_results, headers=headers, tablefmt="pipe"))
    print("\n")

    # 2. WFA Period Average Performance
    periods = [
        ('2024/6/1', '2025/12/31'), ('2024/1/2', '2025/5/31'), ('2023/1/2', '2024/12/31'),
        ('2022/1/2', '2024/5/31'), ('2021/6/1', '2023/12/31'), ('2021/1/2', '2023/5/30'),
        ('2020/1/2', '2022/12/31'), ('2019/6/1', '2022/5/30'), ('2019/1/2', '2021/12/31'),
    ]

    wfa_summary = []
    for cfg in configs:
        calmars = []
        cagrs = []
        mdds = []
        for start, end in periods:
            start_dt = pd.to_datetime(start)
            end_dt = pd.to_datetime(end)

            eq, trades, cost = bt.run(cfg["sma"], cfg["roc"], cfg["sl"], start_dt, end_dt, rebalance_interval=cfg["reb"])

            c, m, cl = calculate_metrics(eq)
            calmars.append(cl)
            cagrs.append(c)
            mdds.append(m)

        wfa_summary.append([
            cfg["name"],
            f"{np.mean(cagrs):.2%}",
            f"{np.min(mdds):.2%}", # Worst case MDD across periods
            f"{np.mean(calmars):.2f}",
            f"{np.std(calmars):.2f}" # Stability of performance
        ])

    print("### WFA Sub-Period Summary (Averages)")
    wfa_headers = ["Strategy", "Avg CAGR", "Worst MDD", "Avg Calmar", "Calmar StdDev"]
    print(tabulate(wfa_summary, headers=wfa_headers, tablefmt="pipe"))

if __name__ == "__main__":
    main()
