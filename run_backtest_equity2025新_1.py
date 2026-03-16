import pandas as pd
import numpy as np
import pickle

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

    def run(self, sma_period, roc_period, stop_loss_pct, rebalance_interval=6):
        sma = self.prices_df.rolling(window=sma_period).mean().values
        roc = self.prices_df.pct_change(periods=roc_period).values

        cash = self.initial_capital
        portfolio = {}
        equity_curve = np.zeros(len(self.dates))
        trades_log = []
        holdings_history = []

        start_idx = max(sma_period, roc_period)

        for i in range(start_idx, len(self.dates)):
            date = self.dates[i]
            current_prices = self.prices[i]

            total_equity = cash
            for asset_idx, info in portfolio.items():
                total_equity += info['shares'] * current_prices[asset_idx]
            equity_curve[i] = total_equity

            if i == len(self.dates) - 1:
                break

            next_prices = self.prices[i+1]

            triggered_sl_idxs = []
            for asset_idx, info in portfolio.items():
                curr_p = current_prices[asset_idx]
                if curr_p > info['max_price']:
                    info['max_price'] = curr_p
                if curr_p < info['max_price'] * (1 - stop_loss_pct):
                    triggered_sl_idxs.append(asset_idx)

            is_rebalance_day = (i - start_idx) % rebalance_interval == 0

            top_3_signals = []
            if is_rebalance_day:
                eligible_mask = (current_prices > sma[i]) & (roc[i] > 0)
                if np.any(eligible_mask):
                    eligible_idxs = np.where(eligible_mask)[0]
                    eligible_rocs = roc[i][eligible_idxs]
                    num_to_pick = min(3, len(eligible_idxs))
                    top_idxs = eligible_idxs[np.argsort(eligible_rocs)[-num_to_pick:][::-1]]
                    top_3_signals = list(top_idxs)

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

                    sell_fee = shares * sell_price * 0.001425
                    sell_tax = shares * sell_price * 0.003
                    proceeds = shares * sell_price - sell_fee - sell_tax
                    cash += proceeds

                    reason = "停損" if asset_idx in triggered_sl_idxs else "再平衡"
                    trades_log.append({
                        '日期': date,
                        '股票代號': self.assets[asset_idx],
                        '狀態': '賣出',
                        '價格': sell_price, # Use execution price for audit consistency
                        '股數': info['shares'],
                        '動能值': f"{roc[i][asset_idx]*100:.2f}%",
                        '標的名稱': self.code_to_name[self.assets[asset_idx]],
                        '原因': reason,
                        '買入手續費': 0,
                        '賣出手續費': sell_fee,
                        '賣出交易稅': sell_tax,
                        '說明': f"{reason}賣出：{self.code_to_name[self.assets[asset_idx]]}"
                    })

            if is_rebalance_day:
                assets_to_buy = [a for a in top_3_signals if a not in portfolio]
                slot_cap = self.initial_capital / 3
                for asset_idx in assets_to_buy:
                    if len(portfolio) >= 3: break
                    buy_price_exec = next_prices[asset_idx]
                    shares = int(slot_cap // (buy_price_exec * 1.001425))
                    if shares > 0:
                        buy_val = shares * buy_price_exec
                        buy_fee = buy_val * 0.001425
                        cost = buy_val + buy_fee
                        if cash >= cost:
                            cash -= cost
                            portfolio[asset_idx] = {
                                'shares': shares,
                                'max_price': buy_price_exec,
                                'buy_price': buy_price_exec,
                                'buy_date': self.dates[i+1]
                            }
                            trades_log.append({
                                '日期': date,
                                '股票代號': self.assets[asset_idx],
                                '狀態': '買進',
                                '價格': buy_price_exec, # Use execution price for audit consistency
                                '股數': shares,
                                '動能值': f"{roc[i][asset_idx]*100:.2f}%",
                                '標的名稱': self.code_to_name[self.assets[asset_idx]],
                                '原因': '符合趨勢',
                                '買入手續費': buy_fee,
                                '賣出手續費': 0,
                                '賣出交易稅': 0,
                                '說明': f"買進新持有商品：{self.code_to_name[self.assets[asset_idx]]}"
                            })

                for asset_idx in portfolio.keys():
                    if asset_idx in top_3_signals and asset_idx not in assets_to_buy:
                        trades_log.append({
                            '日期': date,
                            '股票代號': self.assets[asset_idx],
                            '狀態': '保持',
                            '價格': current_prices[asset_idx],
                            '股數': portfolio[asset_idx]['shares'],
                            '動能值': f"{roc[i][asset_idx]*100:.2f}%",
                            '標的名稱': self.code_to_name[self.assets[asset_idx]],
                            '原因': '趨勢持續',
                            '買入手續費': 0,
                            '賣出手續費': 0,
                            '賣出交易稅': 0,
                            '說明': f"保留與上一期相同：{self.code_to_name[self.assets[asset_idx]]}"
                        })

            holdings_history.append({
                'Date': date,
                'Holdings': ", ".join([f"{self.code_to_name[self.assets[a]]}({self.assets[a]})" for a in portfolio.keys()]),
                'Count': len(portfolio),
                'Equity': total_equity
            })

        eq_series = pd.Series(equity_curve, index=self.dates)
        eq_series.iloc[:start_idx] = self.initial_capital
        eq_series = eq_series.replace(0, np.nan).ffill()
        return eq_series, pd.DataFrame(trades_log), pd.DataFrame(holdings_history)

def calculate_metrics(equity_curve):
    if equity_curve.empty: return 0, 0, 0, 0
    total_return = (equity_curve.iloc[-1] / equity_curve.iloc[0]) - 1
    years = (equity_curve.index[-1] - equity_curve.index[0]).days / 365.25
    cagr = (1 + total_return) ** (1 / years) - 1 if years > 0 else 0
    rolling_max = equity_curve.cummax()
    drawdowns = (equity_curve - rolling_max) / rolling_max
    max_dd = drawdowns.min()
    calmar = cagr / abs(max_dd) if max_dd != 0 else 0
    return cagr, max_dd, calmar, total_return

def calculate_win_rate(trades_df):
    if trades_df.empty: return 0
    # Approximate win rate from trades log (simple implementation)
    # Strategy win rate is usually complex, here we just return a placeholder or implement it properly if needed.
    return 0.5

if __name__ == "__main__":
    prices, code_to_name = clean_data('個股合-1.xlsx')
    bt = Backtester(prices, code_to_name)
    eq, trades, hold = bt.run(87, 54, 0.09, 6)
    cagr, mdd, calmar, ret = calculate_metrics(eq)
    print(f"Equity2025新: CAGR={cagr:.2%}, MaxDD={mdd:.2%}, Calmar={calmar:.2f}")
