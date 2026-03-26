import pandas as pd
import numpy as np

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

class FinalBacktester:
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

        # Performance Tracking
        surplus_pool = float(self.initial_capital)
        slots = {0: None, 1: None, 2: None}
        equity_curve = [float(self.initial_capital)] * (loop_start - first_idx)

        for i in range(loop_start, last_idx + 1):
            current_prices = self.prices[i]
            # Equity Calculation
            stock_mv = sum(info['shares'] * current_prices[info['asset_idx']] for info in slots.values() if info)
            total_equity = surplus_pool + stock_mv
            equity_curve.append(total_equity)

            if i == last_idx: break
            next_prices = self.prices[i+1] # Execution T+1

            # Stop Loss (Signal T, Execution T+1)
            for s_id, info in slots.items():
                if info:
                    if current_prices[info['asset_idx']] > info['max_p']: info['max_p'] = current_prices[info['asset_idx']]
                    if current_prices[info['asset_idx']] < info['max_p'] * (1 - stop_loss_pct):
                        # Sell T+1
                        proceeds = info['shares'] * next_prices[info['asset_idx']] * 0.995575
                        surplus_pool += proceeds
                        slots[s_id] = None

            # Rebalance (Signal T, Execution T+1)
            if (i - loop_start) % int(rebalance_interval) == 0:
                # Value at execution T+1
                liquid_val = surplus_pool + sum(info['shares'] * next_prices[info['asset_idx']] * 0.995575 for info in slots.values() if info)

                # Selection at signal date T
                sorted_idx = np.argsort(roc[i])[::-1]
                top_3 = []
                for idx in sorted_idx:
                    if len(top_3) >= 3: break
                    if current_prices[idx] > sma[i][idx] and roc[i][idx] > 0:
                        top_3.append(idx)

                # Execution rebalance T+1
                new_slots = {0: None, 1: None, 2: None}
                # Keepers from Signal T to Execution T+1
                kept_indices = {info['asset_idx']: sid for sid, info in slots.items() if info and info['asset_idx'] in top_3}
                for idx, sid in kept_indices.items():
                    new_slots[sid] = slots[sid]

                # Rule D.2: Sell proceeds reinvested immediately up to 10M
                # Logic: surplus_pool is handled by subtracting the new costs from the total liquid_val.
                new_signals = [idx for idx in top_3 if idx not in kept_indices]
                for idx in new_signals:
                    target_sid = None
                    for sid in range(3):
                        if new_slots[sid] is None:
                            target_sid = sid
                            break
                    if target_sid is None: break

                    budget = 10000000.0
                    buy_price_exec = next_prices[idx]
                    cost_per_share = buy_price_exec * 1.001425
                    shares = (int(budget // cost_per_share) // 1000) * 1000
                    if shares > 0:
                        new_slots[target_sid] = {'asset_idx': idx, 'shares': shares, 'max_p': buy_price_exec}

                # Actual Performance Adjustment
                new_mv_cost = sum(info['shares'] * next_prices[info['asset_idx']] * 1.001425 for info in new_slots.values() if info)
                surplus_pool = liquid_val - new_mv_cost
                slots = new_slots

        eq_series = pd.Series(equity_curve)
        rolling_max = eq_series.cummax()
        dd_series = (eq_series - rolling_max) / rolling_max
        return eq_series, dd_series

def get_metrics(eq, dd):
    if eq is None or len(eq) < 2: return 0, 0, 0
    total_ret = (eq.iloc[-1] / eq.iloc[0]) - 1
    years = len(eq) / 252.0
    cagr = (1 + total_ret)**(1/years) - 1 if years > 0 and total_ret > -1 else -1
    mdd = dd.min()
    calmar = cagr / abs(mdd) if mdd < 0 else 0
    return cagr, mdd, calmar

prices = clean_data('個股合-1.xlsx')
bt = FinalBacktester(prices)
periods = [('2019-01-02', '2021-12-31'), ('2019-06-01', '2022-05-31'), ('2020-01-02', '2022-12-31'), ('2020-06-01', '2023-05-31'), ('2021-01-02', '2023-12-31'), ('2021-06-01', '2024-05-31'), ('2022-01-02', '2024-12-31'), ('2022-06-01', '2025-05-31'), ('2023-01-02', '2025-12-31'), ('2019-01-02', '2025-12-31')]

# SMA 30-120, ROC 30-120, Reb 5-10
for sma in [112, 100, 115]:
    for roc in [80, 54, 85]:
        for sl in [0.08, 0.09]:
            for reb in [6, 10]:
                print(f"\n--- Checking {sma}, {roc}, {sl}, {reb} ---")
                all_m = []
                for i, (s, e) in enumerate(periods):
                    eq, dd = bt.run(sma, roc, sl, reb, s, e)
                    c, m, cl = get_metrics(eq, dd)
                    all_m.append((c, m, cl))

                min_mdd = min(r[1] for r in all_m)
                full_calmar = all_m[-1][2]
                min_calmar = min(r[2] for r in all_m[:-1])
                print(f"FullCal: {full_calmar:.2f} MinCal: {min_calmar:.2f} MinMDD: {min_mdd:.2%}")
