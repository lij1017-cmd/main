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
        # Market proxy: simple average of all stocks
        self.market_price = prices_df.mean(axis=1).values

    def run(self, sma_period, roc_period, stop_loss_pct, rebalance_interval, start_date, end_date, market_sma_period=150):
        sma = self.prices_df.rolling(window=int(sma_period)).mean().values
        roc = self.prices_df.pct_change(periods=int(roc_period)).values

        # Market Filter
        mkt_ma = pd.Series(self.market_price).rolling(window=market_sma_period).mean().values
        market_on = self.market_price > mkt_ma

        mask = (self.dates >= pd.to_datetime(start_date)) & (self.dates <= pd.to_datetime(end_date))
        all_indices = np.where(mask)[0]
        if len(all_indices) == 0: return None, None
        first_idx, last_idx = all_indices[0], all_indices[-1]
        loop_start = max(first_idx, 150)

        surplus_pool = float(self.initial_capital)
        slots = {0: None, 1: None, 2: None}
        equity_curve = [float(self.initial_capital)] * (loop_start - first_idx)

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
                current_total = surplus_pool + sum(info['shares'] * next_prices[info['asset_idx']] * 0.995575 for info in slots.values() if info)

                # If Market is OFF, go to cash
                if not market_on[i]:
                    surplus_pool = current_total
                    slots = {0: None, 1: None, 2: None}
                    continue

                # Selection
                sorted_idx = np.argsort(roc[i])[::-1]
                top_3 = []
                for idx in sorted_idx:
                    if len(top_3) >= 3: break
                    if current_prices[idx] > sma[i][idx] and roc[i][idx] > 0:
                        top_3.append(idx)

                new_slots = {0: None, 1: None, 2: None}
                kept_indices = {}
                for s_id, info in slots.items():
                    if info and info['asset_idx'] in top_3:
                        kept_indices[info['asset_idx']] = s_id

                kept_mv = sum(slots[kept_indices[idx]]['shares'] * next_prices[idx] for idx in kept_indices)

                # Budget for new
                for idx in top_3:
                    if idx in kept_indices:
                        new_slots[kept_indices[idx]] = slots[kept_indices[idx]]
                    else:
                        budget = 10000000.0
                        buy_p = next_prices[idx]
                        shares = (int(budget // (buy_p * 1.001425)) // 1000) * 1000
                        if shares > 0:
                            for s_id in range(3):
                                if new_slots[s_id] is None:
                                    new_slots[s_id] = {'asset_idx': idx, 'shares': shares, 'max_p': buy_p}
                                    break

                new_mv = sum(info['shares'] * next_prices[info['asset_idx']] for info in new_slots.values() if info)
                surplus_pool = current_total - (new_mv * 1.001425)
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

    full_calmar = results[-1][2]
    sub_calmars = [r[2] for r in results[:-1]]
    sub_mdds = [r[1] for r in results[:-1]]
    full_mdd = results[-1][1]

    # Penalize MDD > 25% very heavily
    penalty = 0
    for m in sub_mdds + [full_mdd]:
        if m < -0.25:
            penalty += (abs(m) - 0.25) * 1000 # Massive penalty

    if penalty > 0: return -penalty

    # Otherwise, score is sum of calmars
    return full_calmar + sum(sub_calmars)

def aco_search(bt, periods):
    # Try multiple iterations with more ants
    ants = 20
    iterations = 5
    best_params = None
    best_score = -float('inf')

    for it in range(iterations):
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
                print(f"Iter {it}: New Best {cand} Score {score:.2f}")

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
    if not best:
        # Fallback to a scan
        print("ACO failed to find valid set, starting grid scan...")
        for sl in [0.03, 0.05, 0.07]:
            for reb in [6, 10]:
                for sma in [80, 110]:
                    for roc in [50, 90]:
                        cand = (sma, roc, sl, reb)
                        score = fitness(cand, bt, periods)
                        if score > best_score:
                            best_score = score
                            best_params = cand
                            print(f"Grid: New Best {cand} Score {score:.2f}")
        best = best_params

    print(f"Final Best Params: {best}")

    sma, roc, sl, reb = best
    print("\nFinal Performance Audit:")
    for i, (s, e) in enumerate(periods):
        eq, dd = bt.run(sma, roc, sl, reb, s, e)
        c, m, cl = calculate_metrics(eq, dd)
        period_name = f"Window {i+1}" if i < 9 else "Full Period"
        print(f"{period_name} ({s} to {e}): CAGR={c:.2%}, MDD={m:.2%}, Calmar={cl:.2f}")
