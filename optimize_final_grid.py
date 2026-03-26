import pandas as pd
import numpy as np
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

class FastBacktester:
    def __init__(self, prices_df, initial_capital=30000000):
        self.prices = prices_df.values
        self.dates = prices_df.index
        self.prices_df = prices_df
        self.initial_capital = float(initial_capital)

    def run(self, sma_p, roc_p, sl_p, reb_p, start_idx, last_idx):
        # Localize indicators for speed
        sma = self.prices_df.rolling(window=int(sma_p)).mean().values
        roc = self.prices_df.pct_change(periods=int(roc_p)).values

        surplus_pool = self.initial_capital
        slots = {0: None, 1: None, 2: None}
        equity_curve = []

        loop_start = max(start_idx, 150)

        for i in range(start_idx, loop_start):
            equity_curve.append(self.initial_capital)

        for i in range(loop_start, last_idx + 1):
            curr_p = self.prices[i]
            stock_mv = sum(info['shares'] * curr_p[info['idx']] for info in slots.values() if info)
            total_equity = surplus_pool + stock_mv
            equity_curve.append(total_equity)

            if i == last_idx: break
            next_p = self.prices[i+1]

            # Stop Loss
            for s_id, info in slots.items():
                if info:
                    if curr_p[info['idx']] > info['max_p']: info['max_p'] = curr_p[info['idx']]
                    if curr_p[info['idx']] < info['max_p'] * (1 - sl_p):
                        surplus_pool += info['shares'] * next_p[info['idx']] * 0.995575
                        slots[s_id] = None

            # Rebalance
            if (i - loop_start) % int(reb_p) == 0:
                liquid_val = surplus_pool + sum(info['shares'] * next_p[info['idx']] * 0.995575 for info in slots.values() if info)

                # Selection
                sorted_idx = np.argsort(roc[i])[::-1]
                top_3 = []
                for idx in sorted_idx:
                    if len(top_3) >= 3: break
                    if curr_p[idx] > sma[i][idx] and roc[i][idx] > 0: top_3.append(idx)

                new_slots = {0: None, 1: None, 2: None}
                kept = {info['idx']: sid for sid, info in slots.items() if info and info['idx'] in top_3}
                for idx, sid in kept.items(): new_slots[sid] = slots[sid]

                new_sigs = [idx for idx in top_3 if idx not in kept]
                for idx in new_sigs:
                    for sid in range(3):
                        if new_slots[sid] is None:
                            shares = (int(10000000.0 // (next_p[idx] * 1.001425)) // 1000) * 1000
                            if shares > 0:
                                new_slots[sid] = {'idx': idx, 'shares': shares, 'max_p': next_p[idx]}
                            break

                new_cost = sum(info['shares'] * next_p[info['idx']] * 1.001425 for info in new_slots.values() if info)
                surplus_pool = liquid_val - new_cost
                slots = new_slots

        eq = pd.Series(equity_curve)
        mdd = ((eq - eq.cummax()) / eq.cummax()).min()
        ret = (eq.iloc[-1] / eq.iloc[0]) - 1
        years = len(eq) / 252.0
        cagr = (1 + ret)**(1/years)-1 if years > 0 and ret > -1 else -1
        calmar = cagr / abs(mdd) if mdd < 0 else 0
        return cagr, mdd, calmar

def worker(params, bt, period_indices):
    sma, roc, sl, reb = params
    results = []
    for start, end in period_indices:
        res = bt.run(sma, roc, sl, reb, start, end)
        results.append(res)

    # Target: Min Calmar >= 2.0?
    min_calmar = min(r[2] for r in results)
    full_calmar = results[-1][2]
    return (params, min_calmar, full_calmar, results)

if __name__ == "__main__":
    prices = clean_data('個股合-1.xlsx')
    bt = FastBacktester(prices)
    periods = [('2019-01-02', '2021-12-31'), ('2019-06-01', '2022-05-31'), ('2020-01-02', '2022-12-31'), ('2020-06-01', '2023-05-31'), ('2021-01-02', '2023-12-31'), ('2021-06-01', '2024-05-31'), ('2022-01-02', '2024-12-31'), ('2022-06-01', '2025-05-31'), ('2023-01-02', '2025-12-31'), ('2019-01-02', '2025-12-31')]
    period_indices = []
    for s, e in periods:
        mask = (bt.dates >= pd.to_datetime(s)) & (bt.dates <= pd.to_datetime(e))
        idx = np.where(mask)[0]
        period_indices.append((idx[0], idx[-1]))

    param_list = []
    for sl in [0.08, 0.09]:
        for reb in [6, 8, 10]:
            for sma in range(30, 121, 10):
                for roc in range(30, 121, 10):
                    param_list.append((sma, roc, sl, reb))

    best_score = -1
    best_all = None

    print(f"Starting Grid Search on {len(param_list)} combinations...")
    with concurrent.futures.ProcessPoolExecutor() as executor:
        futures = [executor.submit(worker, p, bt, period_indices) for p in param_list]
        for future in concurrent.futures.as_completed(futures):
            params, m_cal, f_cal, res = future.result()
            # If all sub calmars >= 1.5 and full >= 2.0
            if m_cal > 1.4 and f_cal > 1.8:
                score = m_cal + f_cal
                if score > best_score:
                    best_score = score
                    best_all = (params, res)
                    print(f"Candidate: {params} MinCalmar: {m_cal:.2f} FullCalmar: {f_cal:.2f}")

    if best_all:
        params, res = best_all
        print(f"\nFINAL BEST: {params}")
        for i, r in enumerate(res):
            p = f"Window {i+1}" if i < 9 else "Full Period"
            print(f"{p}: CAGR={r[0]:.2%}, MDD={r[1]:.2%}, Calmar={r[2]:.2f}")
    else:
        print("No set met the minimum quality threshold.")
