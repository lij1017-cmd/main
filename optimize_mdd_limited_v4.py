import pandas as pd
import numpy as np
import os
import concurrent.futures

def clean_data(filepath):
    df_raw = pd.read_excel(filepath, header=None)
    stock_codes = df_raw.iloc[0, 2:].values
    stock_names = df_raw.iloc[1, 2:].values
    dates = pd.to_datetime(df_raw.iloc[2:, 1])
    prices = df_raw.iloc[2:, 2:].astype(float)
    prices.index = dates
    prices.columns = stock_codes
    prices = prices.ffill().bfill()
    return prices

class Backtester:
    def __init__(self, prices_df, initial_capital=30000000):
        self.prices_df = prices_df
        self.prices = prices_df.values
        self.dates = prices_df.index
        self.initial_capital = initial_capital
        self.market_price = prices_df.mean(axis=1).values

    def run(self, sma_period, roc_period, stop_loss_pct, rebalance_interval, start_idx, last_idx):
        sma = self.prices_df.rolling(window=int(sma_period)).mean().values
        roc = self.prices_df.pct_change(periods=int(roc_period)).values

        # Market Filter: stay in cash if market is below its own SMA
        # This helps reduce MDD during bear markets.
        mkt_ma = pd.Series(self.market_price).rolling(window=120).mean().values
        market_on = self.market_price > mkt_ma

        loop_start = max(start_idx, 150)
        surplus_pool = float(self.initial_capital)
        slots = {0: None, 1: None, 2: None}
        equity_curve = [float(self.initial_capital)] * (loop_start - start_idx)

        for i in range(loop_start, last_idx + 1):
            current_prices = self.prices[i]
            stock_mv = sum(info['shares'] * current_prices[info['asset_idx']] for info in slots.values() if info)
            equity_curve.append(surplus_pool + stock_mv)

            if i == last_idx: break
            next_prices = self.prices[i+1]

            # Stop Loss (Signal T, Execution T+1)
            for s_id, info in slots.items():
                if info:
                    if current_prices[info['asset_idx']] > info['max_p']: info['max_p'] = current_prices[info['asset_idx']]
                    if current_prices[info['asset_idx']] < info['max_p'] * (1 - stop_loss_pct):
                        surplus_pool += info['shares'] * next_prices[info['asset_idx']] * 0.995575
                        slots[s_id] = None

            # Rebalance (Signal T, Execution T+1)
            if (i - loop_start) % int(rebalance_interval) == 0:
                liquid_val = surplus_pool + sum(info['shares'] * next_prices[info['asset_idx']] * 0.995575 for info in slots.values() if info)

                # Rule D.1: Fixed 30M reset logic sizing
                if not market_on[i]:
                    surplus_pool = liquid_val
                    slots = {0: None, 1: None, 2: None}
                    continue

                sorted_idx = np.argsort(roc[i])[::-1]
                top_3 = []
                for idx in sorted_idx:
                    if len(top_3) >= 3: break
                    if current_prices[idx] > sma[i][idx] and roc[i][idx] > 0:
                        top_3.append(idx)

                new_slots = {0: None, 1: None, 2: None}
                kept = {info['asset_idx']: sid for sid, info in slots.items() if info and info['asset_idx'] in top_3}
                for idx, sid in kept.items(): new_slots[sid] = slots[sid]

                new_sigs = [idx for idx in top_3 if idx not in kept]
                for idx in new_sigs:
                    target_sid = None
                    for sid in range(3):
                        if new_slots[sid] is None:
                            target_sid = sid
                            break
                    if target_sid is None: break

                    budget = 10000000.0
                    buy_p = next_prices[idx]
                    shares = (int(budget // (buy_p * 1.001425)) // 1000) * 1000
                    if shares > 0:
                        new_slots[target_sid] = {'asset_idx': idx, 'shares': shares, 'max_p': buy_p}

                new_mv_cost = sum(info['shares'] * next_prices[info['asset_idx']] * 1.001425 for info in new_slots.values() if info)
                surplus_pool = liquid_val - new_mv_cost
                slots = new_slots

        eq = pd.Series(equity_curve)
        mdd = ((eq - eq.cummax()) / eq.cummax()).min()
        cagr = (eq.iloc[-1]/eq.iloc[0])**(252/len(eq)) - 1 if len(eq) > 0 else 0
        calmar = cagr / abs(mdd) if mdd < 0 else 0
        return cagr, mdd, calmar

def worker(params, bt, period_indices):
    sma, roc, sl, reb = params
    results = []
    for s, e in period_indices:
        results.append(bt.run(sma, roc, sl, reb, s, e))

    mdds = [r[1] for r in results]
    calmars = [r[2] for r in results]
    min_mdd = min(mdds)
    full_calmar = calmars[-1]
    min_sub_calmar = min(calmars[:-1])

    # Fitness priority: MDD > -0.25, then maximize Calmar
    if min_mdd < -0.25:
        score = -100 + min_mdd # Penalty
    else:
        score = full_calmar + min_sub_calmar

    return params, score, min_mdd, full_calmar, min_sub_calmar

if __name__ == "__main__":
    prices = clean_data('個股合-1.xlsx')
    bt = Backtester(prices)
    periods = [('2019-01-02', '2021-12-31'), ('2019-06-01', '2022-05-31'), ('2020-01-02', '2022-12-31'), ('2020-06-01', '2023-05-31'), ('2021-01-02', '2023-12-31'), ('2021-06-01', '2024-05-31'), ('2022-01-02', '2024-12-31'), ('2022-06-01', '2025-05-31'), ('2023-01-02', '2025-12-31'), ('2019-01-02', '2025-12-31')]
    p_indices = []
    for s, e in periods:
        mask = (bt.dates >= pd.to_datetime(s)) & (bt.dates <= pd.to_datetime(e))
        idx = np.where(mask)[0]
        p_indices.append((idx[0], idx[-1]))

    param_list = []
    # Narrowing SMA/ROC search for efficiency
    for sl in [0.03, 0.05, 0.07, 0.09]:
        for reb in [6, 8, 10]:
            for sma in range(30, 151, 15):
                for roc in range(30, 151, 15):
                    param_list.append((sma, roc, sl, reb))

    best_score = -1000
    best_params = None

    print(f"Searching {len(param_list)} sets...")
    with concurrent.futures.ProcessPoolExecutor() as executor:
        futures = [executor.submit(worker, p, bt, p_indices) for p in param_list]
        for f in concurrent.futures.as_completed(futures):
            p, s, mdd, fc, mc = f.result()
            if s > best_score:
                best_score = s
                best_params = (p, mdd, fc, mc)
                print(f"Best: {p} Score:{s:.2f} MinMDD:{mdd:.2%} FullCal:{fc:.2f} MinSub:{mc:.2f}")

    print(f"\nFINAL BEST: {best_params}")
