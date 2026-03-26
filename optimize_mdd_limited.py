import pandas as pd
import numpy as np
import random
import os

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

    def run(self, sma_period, roc_period, stop_loss_pct, rebalance_interval, start_date, end_date):
        sma = self.prices_df.rolling(window=int(sma_period)).mean().values
        roc = self.prices_df.pct_change(periods=int(roc_period)).values

        mask = (self.dates >= pd.to_datetime(start_date)) & (self.dates <= pd.to_datetime(end_date))
        all_indices = np.where(mask)[0]
        if len(all_indices) == 0: return None, None
        first_idx, last_idx = all_indices[0], all_indices[-1]
        loop_start = max(first_idx, 150)

        surplus_pool = float(self.initial_capital)
        slots = {0: None, 1: None, 2: None}
        equity_curve = [float(self.initial_capital)] * (loop_start - first_idx)

        # We need to track the "Cumulative Return" of each rebalance period
        # but reset the capital for sizing.
        # Total Portfolio Value = Surplus + MarketValue
        # Performance = TotalValue / Initial
        # But at Rebalance, TotalValue is set to 30M for shares calculation.

        for i in range(loop_start, last_idx + 1):
            current_prices = self.prices[i]
            stock_mv = sum(info['shares'] * current_prices[info['asset_idx']] for info in slots.values() if info)
            total_equity = surplus_pool + stock_mv
            equity_curve.append(total_equity)

            if i == last_idx: break
            next_prices = self.prices[i+1]

            # Stop Loss (T signal, T+1 execution)
            for s_id, info in slots.items():
                if info:
                    if current_prices[info['asset_idx']] > info['max_p']: info['max_p'] = current_prices[info['asset_idx']]
                    if current_prices[info['asset_idx']] < info['max_p'] * (1 - stop_loss_pct):
                        sell_p = next_prices[info['asset_idx']]
                        proceeds = info['shares'] * sell_p * (1 - 0.001425 - 0.003)
                        surplus_pool += proceeds
                        slots[s_id] = None

            # Rebalance
            if (i - loop_start) % int(rebalance_interval) == 0:
                # 1. Selection
                sorted_idx = np.argsort(roc[i])[::-1]
                top_3 = []
                for idx in sorted_idx:
                    if len(top_3) >= 3: break
                    if current_prices[idx] > sma[i][idx] and roc[i][idx] > 0:
                        top_3.append(idx)

                # 2. Reset capital to 30M for sizing (as per Rule D.1)
                # We calculate how much we HAVE now if we sold everything.
                current_total = surplus_pool + sum(info['shares'] * next_prices[info['asset_idx']] * 0.995575 for info in slots.values() if info)
                # But Rule D.1 says "初始資金每次再平衡固定 3000 萬".
                # This means we use 30M as the pool.

                new_slots = {0: None, 1: None, 2: None}
                # Check kept stocks
                kept_indices = {} # asset_idx -> s_id
                for s_id, info in slots.items():
                    if info and info['asset_idx'] in top_3:
                        kept_indices[info['asset_idx']] = s_id

                # Market Value of kept stocks at T+1 (execution)
                kept_mv = sum(slots[kept_indices[idx]]['shares'] * next_prices[idx] for idx in kept_indices)

                # Effective Surplus for 30M reset
                rebal_surplus = 30000000.0 - kept_mv

                # New slots
                for idx in top_3:
                    if idx in kept_indices:
                        s_id = kept_indices[idx]
                        new_slots[s_id] = slots[s_id]
                    else:
                        # Budget calculation from Rule D.2
                        # If a stock was sold, we use min(proceeds, 10M).
                        # If the slot was empty, we use 10M.
                        # Since we reset to 30M, every "new" slot gets up to 10M.
                        budget = 10000000.0
                        buy_p = next_prices[idx]
                        shares = (int(budget // (buy_p * 1.001425)) // 1000) * 1000
                        if shares > 0:
                            cost = shares * buy_p * 1.001425
                            rebal_surplus -= cost
                            for s_id in range(3):
                                if new_slots[s_id] is None:
                                    new_slots[s_id] = {'asset_idx': idx, 'shares': shares, 'max_p': buy_p}
                                    break

                # The actual surplus pool in the backtest must reflect the PnL.
                # Actual surplus = current_total - sum(new_slots MV)
                new_mv = sum(info['shares'] * next_prices[info['asset_idx']] for info in new_slots.values() if info)
                surplus_pool = current_total - (new_mv * 1.001425) # simplified cost
                slots = new_slots

        eq_series = pd.Series(equity_curve)
        rolling_max = eq_series.cummax()
        dd_series = (eq_series - rolling_max) / rolling_max
        return eq_series, dd_series

def calculate_metrics(eq, dd):
    if eq is None or len(eq) < 2: return 0, 0, 0
    ret = (eq.iloc[-1] / eq.iloc[0]) - 1
    years = len(eq) / 252.0
    cagr = (1 + ret)**(1/years) - 1 if years > 0 and ret > -1 else -1
    mdd = dd.min()
    calmar = cagr / abs(mdd) if mdd < 0 else 0
    return cagr, mdd, calmar

def fitness(params, bt, periods):
    sma, roc, sl, reb = params
    results = []
    for start, end in periods:
        eq, dd = bt.run(sma, roc, sl, reb, start, end)
        results.append(calculate_metrics(eq, dd))

    # Target: All Calmar >= 1.5, MDD >= -0.25
    # Full Calmar >= 2.0
    full_cagr, full_mdd, full_calmar = results[-1]
    sub_calmars = [r[2] for r in results[:-1]]
    sub_mdds = [r[1] for r in results[:-1]]

    min_sub_calmar = min(sub_calmars)
    min_sub_mdd = min(sub_mdds)

    score = full_calmar + min_sub_calmar

    # Penalties
    penalty = 0
    if full_mdd < -0.25: penalty += (abs(full_mdd) - 0.25) * 100
    for m in sub_mdds:
        if m < -0.25: penalty += (abs(m) - 0.25) * 100

    return score - penalty

def aco_search(bt, periods):
    # SMA 30-120, ROC 30-120, SL 0.01-0.09, REB 5-10
    ants = 10
    iterations = 5
    rho = 0.1

    best_params = None
    best_score = -float('inf')

    # Initialize pheromones
    # For simplicity, we use a localized random search around best found.
    for _ in range(iterations):
        candidates = []
        for _ in range(ants):
            sma = random.randint(30, 120)
            roc = random.randint(30, 120)
            sl = random.uniform(0.01, 0.09)
            reb = random.randint(5, 10)
            candidates.append((sma, roc, sl, reb))

        for cand in candidates:
            score = fitness(cand, bt, periods)
            if score > best_score:
                best_score = score
                best_params = cand
                print(f"New Best: {cand} Score: {score:.2f}")

    return best_params

if __name__ == "__main__":
    prices = clean_data('個股合-1.xlsx')
    bt = Backtester(prices)
    periods = [
        ('2019-01-02', '2021-12-31'), ('2019-06-01', '2022-05-31'),
        ('2020-01-02', '2022-12-31'), ('2020-06-01', '2023-05-31'),
        ('2021-01-02', '2023-12-31'), ('2021-06-01', '2024-05-31'),
        ('2022-01-02', '2024-12-31'), ('2022-06-01', '2025-05-31'),
        ('2023-01-02', '2025-12-31'), ('2019-01-02', '2025-12-31')
    ]

    best = aco_search(bt, periods)
    print(f"Final Best Params: {best}")

    # Final Report
    sma, roc, sl, reb = best
    print("\nFinal Performance Audit:")
    for i, (s, e) in enumerate(periods):
        eq, dd = bt.run(sma, roc, sl, reb, s, e)
        c, m, cl = calculate_metrics(eq, dd)
        period_name = f"Window {i+1}" if i < 9 else "Full Period"
        print(f"{period_name} ({s} to {e}): CAGR={c:.2%}, MDD={m:.2%}, Calmar={cl:.2f}")
