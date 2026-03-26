import pandas as pd
import numpy as np
import os
import sys

def clean_data(filepath):
    df_raw = pd.read_excel(filepath, header=None)
    stock_codes = df_raw.iloc[0, 2:].values
    stock_names = df_raw.iloc[1, 2:].values
    dates = pd.to_datetime(df_raw.iloc[2:, 1])
    prices = df_raw.iloc[2:, 2:].astype(float)
    prices.index = dates
    prices.columns = stock_codes
    code_to_name = dict(zip(stock_codes, stock_names))
    prices = prices.ffill().bfill()
    return prices, code_to_name

class Backtester:
    def __init__(self, prices_df, code_to_name, initial_capital=30000000):
        self.prices_df = prices_df
        self.prices = prices_df.values
        self.dates = prices_df.index
        self.assets = prices_df.columns
        self.code_to_name = code_to_name
        self.initial_capital = initial_capital

    def run(self, sma_period, roc_period, stop_loss_pct, rebalance_interval, start_date, end_date):
        sma = self.prices_df.rolling(window=int(sma_period)).mean().values
        roc = self.prices_df.pct_change(periods=int(roc_period)).values

        # MARKET FILTER (SMA 100 on average of all assets)
        market_avg = self.prices_df.mean(axis=1)
        market_sma = market_avg.rolling(window=100).mean()
        market_trend = market_avg > market_sma

        mask = (self.dates >= pd.to_datetime(start_date)) & (self.dates <= pd.to_datetime(end_date))
        all_indices = np.where(mask)[0]
        if len(all_indices) == 0: return None, None
        first_idx, last_idx = all_indices[0], all_indices[-1]
        start_buffer = max(int(sma_period), int(roc_period))
        loop_start = max(first_idx, start_buffer)

        surplus_pool = float(self.initial_capital)
        slots = {0: None, 1: None, 2: None}
        compounded_base = float(self.initial_capital)
        equity_curve = []

        for i in range(first_idx, loop_start):
            equity_curve.append(float(self.initial_capital))

        for i in range(loop_start, last_idx + 1):
            current_prices = self.prices[i]
            stock_mv = sum(info['shares'] * current_prices[info['asset_idx']] for info in slots.values() if info)
            total_equity_current = surplus_pool + stock_mv
            interval_return = total_equity_current / self.initial_capital
            display_equity = compounded_base * interval_return
            equity_curve.append(display_equity)

            if i == last_idx: break
            next_prices = self.prices[i+1]

            # 1. Stop Loss (T signal, T+1 exit)
            for s_id, info in slots.items():
                if info:
                    if current_prices[info['asset_idx']] > info['max_p']: info['max_p'] = current_prices[info['asset_idx']]
                    if current_prices[info['asset_idx']] < info['max_p'] * (1 - stop_loss_pct):
                        sell_p = next_prices[info['asset_idx']]
                        surplus_pool += info['shares'] * sell_p * (1 - 0.001425 - 0.003)
                        slots[s_id] = None

            # 2. Rebalance (T signal, T+1 execute)
            if (i - loop_start) % int(rebalance_interval) == 0:
                compounded_base = display_equity

                # Check Market Filter
                if not market_trend.iloc[i]:
                    # MARKET WEAK: Sell everything, go to cash
                    for s_id, info in slots.items():
                        if info:
                            sell_p = next_prices[info['asset_idx']]
                            surplus_pool += info['shares'] * sell_p * (1 - 0.001425 - 0.003)
                            slots[s_id] = None
                    # Pool reset
                    surplus_pool = float(self.initial_capital)
                else:
                    top_3 = []
                    sorted_all = np.argsort(roc[i])[::-1]
                    for idx in sorted_all:
                        if len(top_3) >= 3: break
                        if current_prices[idx] > sma[i][idx] and roc[i][idx] > 0: top_3.append(idx)

                    kept_sids = {info['asset_idx']: s_id for s_id, info in slots.items() if info and info['asset_idx'] in top_3}
                    kept_mv = sum(slots[sid]['shares'] * current_prices[sig] for sig, sid in kept_sids.items())
                    surplus_pool = float(self.initial_capital) - kept_mv

                    budgets = []
                    for s_id in range(3):
                        if s_id not in kept_sids.values():
                            if slots[s_id]:
                                sell_p = next_prices[slots[s_id]['asset_idx']]
                                budgets.append(min(slots[s_id]['shares'] * sell_p * (1 - 0.001425 - 0.003), 10000000.0))
                                slots[s_id] = None
                            else: budgets.append(10000000.0)

                    new_signals = [sig for sig in top_3 if sig not in kept_sids]
                    for sig in new_signals:
                        if not budgets: break
                        budget = budgets.pop(0)
                        buy_p_exec = next_prices[sig]
                        shares = (int(budget // (buy_p_exec * 1.001425)) // 1000) * 1000
                        if shares > 0:
                            cost = shares * buy_p_exec * 1.001425
                            if surplus_pool >= cost:
                                surplus_pool -= cost
                                for s_id in range(3):
                                    if slots[s_id] is None:
                                        slots[s_id] = {'asset_idx': sig, 'shares': shares, 'max_p': buy_p_exec}
                                        break

        eq_s = pd.Series(equity_curve)
        rolling_max = eq_s.cummax()
        return eq_s, (eq_s - rolling_max) / rolling_max

def calculate_metrics(eq_s, dd_s):
    if eq_s is None or eq_s.empty: return 0, 0, 0
    total_ret = (eq_s.iloc[-1] / eq_s.iloc[0]) - 1
    years = len(eq_s) / 252.0
    cagr = (1 + total_ret)**(1/years)-1 if years > 0 and total_ret > -1 else -1
    mdd = dd_s.min()
    calmar = cagr / abs(mdd) if mdd < 0 else (cagr if cagr > 0 else 0)
    return cagr, mdd, calmar

if __name__ == "__main__":
    prices, code_to_name = clean_data('個股合-1.xlsx')
    bt = Backtester(prices, code_to_name)
    periods = [('2019-01-02', '2021-12-31'), ('2019-06-01', '2022-05-31'), ('2020-01-02', '2022-12-31'), ('2020-06-01', '2023-05-31'), ('2021-01-02', '2023-12-31'), ('2021-06-01', '2024-05-31'), ('2022-01-02', '2024-12-31'), ('2022-06-01', '2025-05-31'), ('2023-01-02', '2025-12-31'), ('2019-01-02', '2025-12-31')]

    # Check if Market Filter helps MDD
    # Using previous good params
    p = (98, 80, 0.085, 10)
    print(f"Results for {p} WITH Market Filter:")
    for s, e in periods:
        eq, dd = bt.run(*p, s, e)
        c, m, cal = calculate_metrics(eq, dd)
        print(f"{s} - {e}: Calmar={cal:.2f} (CAGR:{c:.2%}, MDD:{m:.2%})")
