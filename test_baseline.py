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

        for i in range(loop_start, last_idx + 1):
            current_prices = self.prices[i]
            stock_mv = sum(info['shares'] * current_prices[info['asset_idx']] for info in slots.values() if info)
            equity_curve.append(surplus_pool + stock_mv)
            if i == last_idx: break

            next_prices = self.prices[i+1]
            # Stop Loss Check (Signal T, Exec T+1)
            for s_id, info in slots.items():
                if info:
                    if current_prices[info['asset_idx']] > info['max_p']: info['max_p'] = current_prices[info['asset_idx']]
                    if current_prices[info['asset_idx']] < info['max_p'] * (1 - stop_loss_pct):
                        surplus_pool += info['shares'] * next_prices[info['asset_idx']] * (1 - 0.001425 - 0.003)
                        slots[s_id] = None

            # Rebalance Check
            if (i - loop_start) % int(rebalance_interval) == 0:
                # Reset capital logic
                total_val = surplus_pool + sum(info['shares'] * next_prices[info['asset_idx']] * (1-0.004425) for info in slots.values() if info)
                surplus_pool = float(self.initial_capital)
                slots = {0: None, 1: None, 2: None}

                # Selection
                sorted_idx = np.argsort(roc[i])[::-1]
                top_3 = []
                for idx in sorted_idx:
                    if len(top_3) >= 3: break
                    if current_prices[idx] > sma[i][idx] and roc[i][idx] > 0: top_3.append(idx)

                for sig in top_3:
                    buy_p = next_prices[sig]
                    budget = 10000000.0
                    shares = (int(budget // (buy_p * 1.001425)) // 1000) * 1000
                    if shares > 0:
                        surplus_pool -= (shares * buy_p * 1.001425)
                        for s_id in range(3):
                            if slots[s_id] is None:
                                slots[s_id] = {'asset_idx': sig, 'shares': shares, 'max_p': buy_p}
                                break
        return pd.Series(equity_curve)

def get_metrics(eq):
    if eq is None or len(eq) < 2: return 0, 0, 0
    ret = (eq.iloc[-1] / eq.iloc[0]) - 1
    years = len(eq) / 252.0
    cagr = (1 + ret)**(1/years) - 1 if years > 0 and ret > -1 else -1
    mdd = ((eq - eq.cummax()) / eq.cummax()).min()
    calmar = cagr / abs(mdd) if mdd < 0 else 0
    return cagr, mdd, calmar

prices = clean_data('個股合-1.xlsx')
bt = Backtester(prices)
# Testing SMA 100, ROC 100, SL 9%, Reb 12
eq = bt.run(100, 100, 0.09, 12, '2019-01-02', '2025-12-31')
c, m, cal = get_metrics(eq)
print(f"Full Period -> CAGR: {c:.2%}, MDD: {m:.2%}, Calmar: {cal:.2f}")
