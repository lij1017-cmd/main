import pandas as pd
import numpy as np
import os
from datetime import datetime

def clean_data(filepath):
    """
    清洗並預處理輸入的 Excel 資料檔。
    """
    df_prices = pd.read_excel(filepath, sheet_name='還原收盤價', header=None)
    df_volume = pd.read_excel(filepath, sheet_name='成交量', header=None)
    stock_codes = df_prices.iloc[0, 1:].values
    stock_names = df_prices.iloc[1, 1:].values
    date_strings = df_prices.iloc[2:, 0].astype(str).str[:8]
    dates = pd.to_datetime(date_strings, format='%Y%m%d')
    prices = df_prices.iloc[2:, 1:].astype(float)
    prices.index = dates
    prices.columns = stock_codes
    volumes = df_volume.iloc[2:, 1:].astype(float)
    volumes.index = dates
    volumes.columns = stock_codes
    code_to_name = dict(zip(stock_codes, stock_names))
    prices = prices.ffill().bfill()
    volumes = volumes.fillna(0)
    return prices, volumes, code_to_name

def calculate_metrics(equity_curve_df, annual=False):
    if equity_curve_df.empty: return 0, 0, 0, 0
    equity = equity_curve_df['權益']
    total_return = (equity.iloc[-1] / equity.iloc[0]) - 1
    days = (equity_curve_df['日期'].iloc[-1] - equity_curve_df['日期'].iloc[0]).days
    years = max(days / 365.25, 0.001)
    cagr = (1 + total_return) ** (1 / years) - 1 if years > 0 else 0
    if annual:
        peak = equity.cummax()
        dd = (equity - peak) / peak
        max_dd = dd.min()
    else:
        max_dd = equity_curve_df['回撤(Drawdown)'].min()
    calmar = cagr / abs(max_dd) if max_dd != 0 else 0
    return cagr, max_dd, calmar, total_return

class BacktesterBreadth:
    def __init__(self, prices, volumes, code_to_name, initial_capital=30000000):
        self.prices_df = prices
        self.volumes_df = volumes
        self.prices = prices.values
        self.volumes = volumes.values
        self.dates = prices.index
        self.assets = prices.columns
        self.code_to_name = code_to_name
        self.initial_capital = initial_capital

    def run(self, sma_period=303, roc_period=14, stop_loss_pct=0.0999, rebalance_interval=9,
            breadth_window=300, breadth_threshold=0.45, market_sma_window=15):

        sma = self.prices_df.rolling(window=sma_period).mean().values
        roc = self.prices_df.pct_change(periods=roc_period).values
        sma5 = self.prices_df.rolling(window=5).mean().values
        sma10 = self.prices_df.rolling(window=10).mean().values
        sma20 = self.prices_df.rolling(window=20).mean().values

        b_sma = self.prices_df.rolling(window=breadth_window).mean().values
        breadth = np.mean(self.prices > b_sma, axis=1)
        market_avg = self.prices_df.mean(axis=1).values
        market_sma = self.prices_df.mean(axis=1).rolling(window=market_sma_window).mean().values
        mkt_filter = (breadth >= breadth_threshold) | (market_avg >= market_sma)

        surplus_pool = float(self.initial_capital)
        slots = {0: None, 1: None, 2: None}
        equity_curve_data = []
        trades_log = []
        trades2_log = []
        holdings_history = []
        peak_equity = float(self.initial_capital)
        start_idx = max(sma_period, roc_period, 20, breadth_window, market_sma_window)

        for i in range(start_idx, len(self.dates)):
            date = self.dates[i]
            stock_mv = sum(info['shares'] * self.prices[i, info['asset_idx']] for info in slots.values() if info and 'asset_idx' in info)
            total_equity = surplus_pool + stock_mv
            if total_equity > peak_equity: peak_equity = total_equity
            drawdown = (total_equity - peak_equity) / peak_equity
            equity_curve_data.append({'日期': date, '權益': total_equity, '回撤(Drawdown)': drawdown, '市場寬度': breadth[i], '市場濾網': '持倉' if mkt_filter[i] else '清倉'})

            h_names = [f"{self.code_to_name[self.assets[info['asset_idx']]]}({self.assets[info['asset_idx']]})" for info in slots.values() if info and 'asset_idx' in info]
            holdings_history.append({'Date': date, 'Holdings': ", ".join(h_names), 'Count': len(h_names), '現金': surplus_pool, '股票市值': stock_mv, '總資產': total_equity, '市場濾網': '持倉' if mkt_filter[i] else '清倉'})

            if i == len(self.dates) - 1: break
            next_prices = self.prices[i+1]

            # 市場濾網
            if not mkt_filter[i]:
                for s_id, info in slots.items():
                    if info and 'asset_idx' in info:
                        a_idx = info['asset_idx']
                        sell_p = next_prices[a_idx]
                        shares = info['shares']
                        sell_fee = shares * sell_p * 0.001425
                        sell_tax = shares * sell_p * 0.003
                        proceeds = shares * sell_p - sell_fee - sell_tax
                        surplus_pool += proceeds

                        trades2_log.append({
                            '買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx], '股票名稱': self.code_to_name[self.assets[a_idx]],
                            'T+1日買進價格': info['entry_price'], '股數': shares, '賣出訊號日期': date, 'T+1日賣出價格': sell_p,
                            '損益': proceeds - info['budget'], '報酬率': (proceeds/info['budget'])-1, '買進原因': '趨勢買進', '賣出原因': '市場濾網清倉'
                        })
                        trades_log.append({
                            '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出', '價格': sell_p, '股數': shares,
                            '標的名稱': self.code_to_name[self.assets[a_idx]], '原因': '市場濾網', '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax
                        })
                        slots[s_id] = None
                continue

            # 停損
            for s_id, info in slots.items():
                if info and 'asset_idx' in info:
                    a_idx = info['asset_idx']
                    curr_p = self.prices[i, a_idx]
                    if curr_p > info['max_price']: info['max_price'] = curr_p
                    if curr_p < info['max_price'] * (1 - stop_loss_pct):
                        sell_p = next_prices[a_idx]
                        shares = info['shares']
                        sell_fee = shares * sell_p * 0.001425
                        sell_tax = shares * sell_p * 0.003
                        proceeds = shares * sell_p - sell_fee - sell_tax
                        surplus_pool += proceeds

                        trades2_log.append({
                            '買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx], '股票名稱': self.code_to_name[self.assets[a_idx]],
                            'T+1日買進價格': info['entry_price'], '股數': shares, '賣出訊號日期': date, 'T+1日賣出價格': sell_p,
                            '損益': proceeds - info['budget'], '報酬率': (proceeds/info['budget'])-1, '買進原因': '趨勢買進', '賣出原因': '停損出場'
                        })
                        trades_log.append({
                            '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出', '價格': sell_p, '股數': shares,
                            '標的名稱': self.code_to_name[self.assets[a_idx]], '原因': '停損', '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax
                        })
                        slots[s_id] = None

            # 再平衡
            if (i - start_idx) % rebalance_interval == 0:
                top_3 = []
                sorted_all = np.argsort(roc[i])[::-1]
                for idx in sorted_all:
                    if len(top_3) >= 3: break
                    p, v = self.prices[i, idx], self.volumes[i, idx]
                    if (p > sma[i, idx] and roc[i, idx] > 0 and (p * v * 1000 > 30000000) and p > sma5[i, idx] and p > sma10[i, idx] and p > sma20[i, idx]):
                        top_3.append(idx)

                for s_id, info in slots.items():
                    if info and 'asset_idx' in info:
                        if info['asset_idx'] not in top_3:
                            a_idx = info['asset_idx']
                            sell_p = next_prices[a_idx]
                            shares = info['shares']
                            sell_fee = shares * sell_p * 0.001425
                            sell_tax = shares * sell_p * 0.003
                            proceeds = shares * sell_p - sell_fee - sell_tax
                            surplus_pool += proceeds

                            trades2_log.append({
                                '買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx], '股票名稱': self.code_to_name[self.assets[a_idx]],
                                'T+1日買進價格': info['entry_price'], '股數': shares, '賣出訊號日期': date, 'T+1日賣出價格': sell_p,
                                '損益': proceeds - info['budget'], '報酬率': (proceeds/info['budget'])-1, '買進原因': '趨勢買進', '賣出原因': '再平衡賣出'
                            })
                            trades_log.append({
                                '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出', '價格': sell_p, '股數': shares,
                                '標的名稱': self.code_to_name[self.assets[a_idx]], '原因': '再平衡', '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax
                            })
                            slots[s_id] = {'pending_budget': 10000000.0}
                    else:
                        slots[s_id] = {'pending_budget': 10000000.0}

                existing = [info['asset_idx'] for info in slots.values() if info and 'asset_idx' in info]
                available_sids = [sid for sid, val in slots.items() if val and 'pending_budget' in val]
                for sig in top_3:
                    if sig not in existing and available_sids:
                        s_id = available_sids.pop(0)
                        buy_p = next_prices[sig]
                        buy_fee_rate = 0.001425
                        shares = (int(10000000 // (buy_p * (1 + buy_fee_rate))) // 1000) * 1000
                        if shares > 0:
                            actual_cost = shares * buy_p * (1 + buy_fee_rate)
                            surplus_pool -= actual_cost
                            slots[s_id] = {'asset_idx': sig, 'shares': shares, 'max_price': buy_p, 'budget': actual_cost, 'entry_date': date, 'entry_price': buy_p}
                            trades_log.append({
                                '訊號日期': date, '股票代號': self.assets[sig], '狀態': '買進', '價格': buy_p, '股數': shares,
                                '標的名稱': self.code_to_name[self.assets[sig]], '原因': '趨勢買進', '買入手續費': shares * buy_p * buy_fee_rate, '賣出手續費': 0, '賣出交易稅': 0
                            })
                        else:
                            slots[s_id] = None

                for sid in list(slots.keys()):
                    if slots[sid] and 'pending_budget' in slots[sid]:
                        slots[sid] = None

        return (pd.DataFrame(equity_curve_data), pd.DataFrame(trades_log), pd.DataFrame(holdings_history), pd.DataFrame(trades2_log))
