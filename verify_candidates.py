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

class Backtester:
    def __init__(self, prices_df, initial_capital=30000000):
        self.prices_df = prices_df
        self.prices = prices_df.values
        self.dates = prices_df.index
        self.initial_capital = initial_capital
        self.market_price = prices_df.mean(axis=1).values

    def run(self, sma_period, roc_period, stop_loss_pct, rebalance_interval, start_date, end_date):
        sma = self.prices_df.rolling(window=int(sma_period)).mean().values
        roc = self.prices_df.pct_change(periods=int(roc_period)).values
        mkt_ma = pd.Series(self.market_price).rolling(window=120).mean().values
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
            equity_curve.append(surplus_pool + stock_mv)
            if i == last_idx: break
            next_prices = self.prices[i+1]
            for s_id, info in slots.items():
                if info:
                    if current_prices[info['asset_idx']] > info['max_p']: info['max_p'] = current_prices[info['asset_idx']]
                    if current_prices[info['asset_idx']] < info['max_p'] * (1 - stop_loss_pct):
                        surplus_pool += info['shares'] * next_prices[info['asset_idx']] * 0.995575
                        slots[s_id] = None
            if (i - loop_start) % int(rebalance_interval) == 0:
                current_total = surplus_pool + sum(info['shares'] * next_prices[info['asset_idx']] * 0.995575 for info in slots.values() if info)
                if not market_on[i]:
                    surplus_pool = current_total
                    slots = {0: None, 1: None, 2: None}
                    continue
                sorted_idx = np.argsort(roc[i])[::-1]
                top_3 = []
                for idx in sorted_idx:
                    if len(top_3) >= 3: break
                    if current_prices[idx] > sma[i][idx] and roc[i][idx] > 0: top_3.append(idx)
                new_slots = {0: None, 1: None, 2: None}
                kept_indices = {info['asset_idx']: sid for sid, info in slots.items() if info and info['asset_idx'] in top_3}
                for idx in top_3:
                    if idx in kept_indices:
                        new_slots[kept_indices[idx]] = slots[kept_indices[idx]]
                    else:
                        shares = (int(10000000.0 // (next_prices[idx] * 1.001425)) // 1000) * 1000
                        if shares > 0:
                            for s_id in range(3):
                                if new_slots[s_id] is None:
                                    new_slots[s_id] = {'asset_idx': idx, 'shares': shares, 'max_p': next_prices[idx]}
                                    break
                new_mv = sum(info['shares'] * next_prices[info['asset_idx']] for info in new_slots.values() if info)
                surplus_pool = current_total - (new_mv * 1.001425)
                slots = new_slots
        return pd.Series(equity_curve)

def get_metrics(eq):
    if eq is None or len(eq) < 2: return 0, 0, 0
    cagr = (eq.iloc[-1]/eq.iloc[0])**(252/len(eq)) - 1
    mdd = ((eq - eq.cummax())/eq.cummax()).min()
    return cagr, mdd, cagr/abs(mdd) if mdd < 0 else 0

prices = clean_data('個股合-1.xlsx')
bt = Backtester(prices)
periods = [('2019-01-02', '2021-12-31'), ('2019-06-01', '2022-05-31'), ('2020-01-02', '2022-12-31'), ('2020-06-01', '2023-05-31'), ('2021-01-02', '2023-12-31'), ('2021-06-01', '2024-05-31'), ('2022-01-02', '2024-12-31'), ('2022-06-01', '2025-05-31'), ('2023-01-02', '2025-12-31'), ('2019-01-02', '2025-12-31')]

# (SMA, ROC, SL, REB)
candidates = [(87, 54, 0.08, 6), (54, 52, 0.09, 9), (72, 47, 0.048, 10), (105, 80, 0.08, 10), (120, 120, 0.09, 10), (30, 30, 0.03, 5)]

for cand in candidates:
    print(f"\n--- Candidate: {cand} ---")
    sma, roc, sl, reb = cand
    for i, (s, e) in enumerate(periods):
        eq = bt.run(sma, roc, sl, reb, s, e)
        c, m, cl = get_metrics(eq)
        period_name = f"Window {i+1}" if i < 9 else "Full Period"
        print(f"{period_name} ({s} to {e}): CAGR={c:.2%}, MDD={m:.2%}, Calmar={cl:.2f}")
