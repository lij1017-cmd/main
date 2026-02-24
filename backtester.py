import pandas as pd
import numpy as np

class Backtester:
    def __init__(self, prices, initial_capital=30000000):
        self.prices = prices.values
        self.dates = prices.index
        self.assets = prices.columns
        self.initial_capital = initial_capital

    def run(self, sma_period, roc_period, stop_loss_pct):
        prices_df = pd.DataFrame(self.prices, index=self.dates, columns=self.assets)
        sma = prices_df.rolling(window=sma_period).mean().values
        roc = prices_df.pct_change(periods=roc_period).values

        capital = self.initial_capital
        portfolio = {}
        equity_curve = np.zeros(len(self.dates))
        trades = []
        holdings_history = []
        action_log = []

        start_idx = max(sma_period, roc_period)

        for i in range(start_idx, len(self.dates) - 1):
            date = self.dates[i]
            current_prices = self.prices[i]
            next_prices = self.prices[i+1]
            total_equity = capital
            assets_to_stop_loss = []

            for asset_idx, info in list(portfolio.items()):
                curr_p = current_prices[asset_idx]
                total_equity += info['shares'] * curr_p
                if curr_p > info['max_price']:
                    info['max_price'] = curr_p
                if curr_p < info['max_price'] * (1 - stop_loss_pct):
                    assets_to_stop_loss.append(asset_idx)

            equity_curve[i] = total_equity
            is_rebalance_day = (i - start_idx) % 5 == 0

            new_portfolio_signals = []
            if is_rebalance_day:
                eligible_mask = (current_prices > sma[i]) & (roc[i] > 0)
                if np.any(eligible_mask):
                    eligible_idxs = np.where(eligible_mask)[0]
                    top_k = min(3, len(eligible_idxs))
                    top_idxs = eligible_idxs[np.argsort(roc[i][eligible_idxs])[-top_k:][::-1]]
                    new_portfolio_signals = list(top_idxs)

            assets_selling_now = set(assets_to_stop_loss)
            if is_rebalance_day:
                for asset_idx in list(portfolio.keys()):
                    if asset_idx not in new_portfolio_signals:
                        assets_selling_now.add(asset_idx)

            for asset_idx in assets_selling_now:
                if asset_idx in portfolio:
                    info = portfolio.pop(asset_idx)
                    sell_price = next_prices[asset_idx]
                    capital += info['shares'] * sell_price
                    reason = '停損出場' if asset_idx in assets_to_stop_loss else '再平衡賣出'
                    trades.append({
                        'Buy_Date': info['buy_date'], 'Asset': self.assets[asset_idx],
                        'Buy_Price': info['buy_price'], 'Sell_Date': self.dates[i+1],
                        'Sell_Price': sell_price, 'Shares': info['shares'],
                        'Return': (sell_price / info['buy_price']) - 1, 'Reason': reason,
                        'Entry_Momentum': info['momentum']
                    })
                    action_log.append({
                        '日期': date, '股票代號': self.assets[asset_idx],
                        '狀態': f"{reason} ({'停損' if asset_idx in assets_to_stop_loss else '剃除'})",
                        '價格': current_prices[asset_idx], '股數': 0, '動能值': roc[i][asset_idx]
                    })

            if is_rebalance_day:
                assets_to_buy = [a for a in new_portfolio_signals if a not in portfolio]
                slot_capital = self.initial_capital / 3 # 10,000,000

                for asset_idx in assets_to_buy:
                    buy_price = next_prices[asset_idx]
                    # Only buy if remaining capital is at least 10,000,000
                    if capital >= slot_capital:
                        shares = slot_capital // buy_price
                        if shares > 0:
                            actual_cost = shares * buy_price
                            capital -= actual_cost
                            portfolio[asset_idx] = {
                                'shares': shares, 'buy_price': buy_price,
                                'buy_date': self.dates[i+1], 'max_price': buy_price,
                                'momentum': roc[i][asset_idx]
                            }

                for asset_idx in range(len(self.assets)):
                    if asset_idx in portfolio:
                        status = "買進新持有商品" if asset_idx in assets_to_buy else "保留與上一期相同之商品"
                        action_log.append({
                            '日期': date, '股票代號': self.assets[asset_idx], '狀態': status,
                            '價格': current_prices[asset_idx], '股數': portfolio[asset_idx]['shares'],
                            '動能值': roc[i][asset_idx]
                        })
            holdings_history.append({'Date': date, 'Holdings': [self.assets[a] for a in portfolio.keys()], 'Equity': total_equity})

        eq_series = pd.Series(equity_curve, index=self.dates).dropna()
        return eq_series[eq_series > 0], pd.DataFrame(trades), pd.DataFrame(holdings_history), pd.DataFrame(action_log)

def calculate_metrics(equity_curve, trades):
    if equity_curve.empty: return 0, 0, 0, 0
    total_return = (equity_curve.iloc[-1] / equity_curve.iloc[0]) - 1
    days = (equity_curve.index[-1] - equity_curve.index[0]).days
    cagr = (1 + total_return) ** (365.25 / days) - 1
    max_dd = ((equity_curve - equity_curve.cummax()) / equity_curve.cummax()).min()
    calmar = cagr / abs(max_dd) if max_dd != 0 else 0
    win_rate = (trades['Return'] > 0).mean() if not trades.empty else 0
    return cagr, max_dd, calmar, win_rate
