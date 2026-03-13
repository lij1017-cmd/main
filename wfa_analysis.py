import pandas as pd
import numpy as np
from tabulate import tabulate

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
    def __init__(self, prices, code_to_name, initial_capital=30000000):
        self.prices_df = prices
        self.prices = prices.values
        self.dates = prices.index
        self.assets = prices.columns
        self.code_to_name = code_to_name
        self.initial_capital = initial_capital

    def run(self, sma_period, roc_period, stop_loss_pct, start_date, end_date):
        # Filter dates
        mask = (self.dates >= start_date) & (self.dates <= end_date)
        period_dates = self.dates[mask]

        # We need some buffer for SMA/ROC if we want to start exactly at start_date
        # But usually in these backtests, the indicators are calculated on the whole dataset
        # then sliced, or we just use the period.
        # Given the request, I will calculate indicators on the full dataset to ensure stability
        # and then slice the execution loop.

        sma = self.prices_df.rolling(window=sma_period).mean().values
        roc = self.prices_df.pct_change(periods=roc_period).values

        # Find global indices for the period
        all_indices = np.where(mask)[0]
        if len(all_indices) == 0:
            return None, None, 0

        first_idx = all_indices[0]
        last_idx = all_indices[-1]

        # Check if we have enough data for indicators at start
        start_buffer = max(sma_period, roc_period)
        loop_start = max(first_idx, start_buffer)

        cash = self.initial_capital
        portfolio = {} # {asset_idx: {'shares': ..., 'max_price': ..., 'buy_price': ..., 'buy_date': ...}}
        equity_curve_list = []
        equity_dates = []

        total_costs = 0
        trade_count = 0

        for i in range(loop_start, last_idx + 1):
            date = self.dates[i]
            current_prices = self.prices[i]

            # 1. Calculate current equity (at today's close)
            total_equity = cash
            for asset_idx, info in portfolio.items():
                total_equity += info['shares'] * current_prices[asset_idx]

            equity_curve_list.append(total_equity)
            equity_dates.append(date)

            if i == last_idx:
                break # End of period

            next_prices = self.prices[i+1]

            # 2. Check for daily Stop Loss
            triggered_sl_idxs = []
            for asset_idx, info in portfolio.items():
                curr_p = current_prices[asset_idx]
                if curr_p > info['max_price']:
                    info['max_price'] = curr_p

                if curr_p < info['max_price'] * (1 - stop_loss_pct):
                    triggered_sl_idxs.append(asset_idx)

            # 3. Check for Rebalancing (every 5 days)
            # Use relative index for rebalancing cycle
            is_rebalance_day = (i - loop_start) % 5 == 0

            top_3_signals = []
            if is_rebalance_day:
                eligible_mask = (current_prices > sma[i]) & (roc[i] > 0)
                if np.any(eligible_mask):
                    eligible_idxs = np.where(eligible_mask)[0]
                    eligible_rocs = roc[i][eligible_idxs]
                    num_to_pick = min(3, len(eligible_idxs))
                    top_idxs = eligible_idxs[np.argsort(eligible_rocs)[-num_to_pick:][::-1]]
                    top_3_signals = list(top_idxs)

            # 4. Execution Logic (T+1 at close)
            assets_to_sell = set(triggered_sl_idxs)
            if is_rebalance_day:
                for asset_idx in portfolio.keys():
                    if asset_idx not in top_3_signals:
                        assets_to_sell.add(asset_idx)

            for asset_idx in list(assets_to_sell):
                if asset_idx in portfolio:
                    info = portfolio.pop(asset_idx)
                    sell_price = next_prices[asset_idx]

                    shares = info['shares']
                    commission = shares * sell_price * 0.001425
                    tax = shares * sell_price * 0.003
                    proceeds = shares * sell_price - commission - tax

                    cash += proceeds
                    total_costs += (commission + tax)
                    trade_count += 1

            if is_rebalance_day:
                assets_to_buy = [a for a in top_3_signals if a not in portfolio]
                slot_cap = self.initial_capital / 3

                for asset_idx in assets_to_buy:
                    if len(portfolio) >= 3: break

                    buy_price_exec = next_prices[asset_idx]
                    commission = (slot_cap / 1.001425) * 0.001425
                    shares = int(slot_cap // (buy_price_exec * 1.001425))

                    if shares > 0:
                        buy_val = shares * buy_price_exec
                        buy_comm = buy_val * 0.001425
                        cost = buy_val + buy_comm

                        if cash >= cost:
                            cash -= cost
                            total_costs += buy_comm
                            trade_count += 1
                            portfolio[asset_idx] = {
                                'shares': shares,
                                'max_price': buy_price_exec,
                                'buy_price': buy_price_exec,
                                'buy_date': self.dates[i+1]
                            }

        eq_series = pd.Series(equity_curve_list, index=equity_dates)
        return eq_series, trade_count, total_costs

def calculate_metrics(equity_curve):
    if equity_curve.empty or len(equity_curve) < 2: return 0, 0, 0
    total_return = (equity_curve.iloc[-1] / equity_curve.iloc[0]) - 1
    days = (equity_curve.index[-1] - equity_curve.index[0]).days
    years = days / 365.25
    cagr = (1 + total_return) ** (1 / years) - 1 if years > 0 else 0

    rolling_max = equity_curve.cummax()
    drawdowns = (equity_curve - rolling_max) / rolling_max
    max_dd = drawdowns.min()
    calmar = cagr / abs(max_dd) if max_dd != 0 else 0

    return cagr, max_dd, calmar

def main():
    data_file = '個股合-1.xlsx'
    prices, code_to_name = clean_data(data_file)
    bt = Backtester(prices, code_to_name)

    periods = [
        ('2024/6/1', '2025/12/31'),
        ('2024/1/2', '2025/5/31'),
        ('2023/1/2', '2024/12/31'),
        ('2022/1/2', '2024/5/31'),
        ('2021/6/1', '2023/12/31'),
        ('2021/1/2', '2023/5/30'),
        ('2020/1/2', '2022/12/31'),
        ('2019/6/1', '2022/5/30'),
        ('2019/1/2', '2021/12/31'),
    ]

    results = []
    for start_str, end_str in periods:
        start_date = pd.to_datetime(start_str)
        end_date = pd.to_datetime(end_str)

        eq, trades, costs = bt.run(35, 56, 0.09, start_date, end_date)
        cagr, mdd, calmar = calculate_metrics(eq)

        results.append([
            f"{start_str} - {end_str}",
            f"{cagr:.2%}",
            f"{mdd:.2%}",
            f"{calmar:.2f}",
            trades,
            f"{int(costs):,}"
        ])

    headers = ["回測期間", "CAGR", "MDD", "Calmar Ratio", "交易筆數", "交易成本總計"]
    print(tabulate(results, headers=headers, tablefmt="pipe"))

if __name__ == "__main__":
    main()
