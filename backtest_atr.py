import pandas as pd
import numpy as np

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

class BacktesterATR:
    """
    回測引擎：加入市場寬度與趨勢雙重確認濾網 (方案 B)，並使用 ATR 動態停損。
    固定單筆投入金額 (10M)，不進行報酬滾入。
    """
    def __init__(self, prices, volumes, code_to_name, initial_capital=30000000):
        self.prices_df = prices
        self.volumes_df = volumes
        self.prices = prices.values
        self.volumes = volumes.values
        self.dates = prices.index
        self.assets = prices.columns
        self.code_to_name = code_to_name
        self.initial_capital = initial_capital

        # 交易成本
        self.commission_rate = 0.001425
        self.tax_rate = 0.003

    def run(self, sma_period, roc_period, stop_loss_type='fixed', stop_loss_val=0.0999,
            atr_period=14, atr_multiplier=3.0,
            rebalance_interval=9, use_market_filter=True,
            breadth_threshold=0.42, mkt_sma_window=14, breadth_window=290):

        # 1. 指標預計算
        sma = self.prices_df.rolling(window=sma_period).mean().values
        roc = self.prices_df.pct_change(periods=roc_period).values
        sma5 = self.prices_df.rolling(window=5).mean().values
        sma10 = self.prices_df.rolling(window=10).mean().values
        sma20 = self.prices_df.rolling(window=20).mean().values

        # ATR 計算 (使用收盤價價差作為 TR 代理)
        tr = self.prices_df.diff().abs()
        atr = tr.ewm(alpha=1/atr_period, adjust=False).mean().values

        # 市場寬度濾網
        breadth_sma_all = self.prices_df.rolling(window=breadth_window).mean().values
        breadth = np.mean(self.prices > breadth_sma_all, axis=1)

        market_avg = self.prices_df.mean(axis=1).values
        market_sma = self.prices_df.mean(axis=1).rolling(window=mkt_sma_window).mean().values

        mkt_filter = (breadth >= breadth_threshold) | (market_avg >= market_sma)

        # 2. 帳戶與槽位初始化
        surplus_pool = float(self.initial_capital)
        slots = {0: None, 1: None, 2: None}

        equity_curve_data = []
        trades_log = []
        trades2_log = []
        holdings_history = []
        daily_details = []

        peak_equity = float(self.initial_capital)
        start_idx = max(sma_period, roc_period, breadth_window, mkt_sma_window, atr_period + 1)

        for i in range(start_idx, len(self.dates)):
            date = self.dates[i]
            current_prices = self.prices[i]

            stock_mv = 0.0
            h_names = []
            for s_id, info in slots.items():
                if info and 'asset_idx' in info:
                    a_idx = info['asset_idx']
                    mv = info['shares'] * current_prices[a_idx]
                    stock_mv += mv
                    h_names.append(f"{self.code_to_name[self.assets[a_idx]]}({self.assets[a_idx]})")
                    daily_details.append({
                        '日期': date, '股票代號': self.assets[a_idx], '股票名稱': self.code_to_name[self.assets[a_idx]],
                        '持有股數': info['shares'], '本日收盤價': current_prices[a_idx], '市值': mv
                    })

            total_equity = surplus_pool + stock_mv
            if total_equity > peak_equity: peak_equity = total_equity
            drawdown = (total_equity - peak_equity) / peak_equity

            equity_curve_data.append({
                '日期': date, '權益': total_equity, '回撤(Drawdown)': drawdown, '市場寬度': breadth[i]
            })

            if i == len(self.dates) - 1: break
            next_prices = self.prices[i+1]

            # 市場濾網全清倉
            if use_market_filter and not mkt_filter[i]:
                for s_id, info in slots.items():
                    if info and 'asset_idx' in info:
                        a_idx = info['asset_idx']
                        sell_price = next_prices[a_idx]
                        sell_fee = info['shares'] * sell_price * self.commission_rate
                        sell_tax = info['shares'] * sell_price * self.tax_rate
                        proceeds = info['shares'] * sell_price - sell_fee - sell_tax
                        surplus_pool += proceeds

                        trades2_log.append({
                            '買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx],
                            '股票名稱': self.code_to_name[self.assets[a_idx]],
                            'T+1日買進價格': info['entry_price'], '股數': info['shares'],
                            '賣出訊號日期': date, 'T+1日賣出價格': sell_price, '損益': proceeds - info['budget'],
                            '報酬率': (proceeds / info['budget']) - 1, '買進原因': info['entry_reason'],
                            '賣出原因': f"市場濾網：寬度({breadth[i]:.1%})與大盤皆弱"
                        })
                        trades_log.append({
                            '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出',
                            '價格': sell_price, '股數': info['shares'], '動能值': f"{roc[i][a_idx]*100:.2f}%",
                            '標的名稱': self.code_to_name[self.assets[a_idx]], '原因': '市場濾網',
                            '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax,
                            '說明': f"市場濾網賣出：{self.code_to_name[self.assets[a_idx]]}"
                        })
                        slots[s_id] = None
                continue

            # 每日檢查停損
            for s_id, info in slots.items():
                if info and 'asset_idx' in info:
                    a_idx = info['asset_idx']
                    if current_prices[a_idx] > info['max_price']:
                        info['max_price'] = current_prices[a_idx]

                    exit_triggered = False
                    reason_str = ""

                    if stop_loss_type == 'fixed':
                        if current_prices[a_idx] < info['max_price'] * (1 - stop_loss_val):
                            exit_triggered = True
                            reason_str = f"固定停損：價格自最高點回落達{stop_loss_val*100:.2f}%"
                    elif stop_loss_type == 'atr':
                        # 動態 ATR 停損：最高價 - N * ATR
                        # 注意：ATR[i] 是當前日期的 ATR
                        current_atr = atr[i][a_idx]
                        stop_price = info['max_price'] - atr_multiplier * current_atr
                        if current_prices[a_idx] < stop_price:
                            exit_triggered = True
                            reason_str = f"ATR 停損：價格({current_prices[a_idx]:.2f})低於停損價({stop_price:.2f}, ATR倍數={atr_multiplier})"

                    if exit_triggered:
                        sell_price = next_prices[a_idx]
                        sell_fee = info['shares'] * sell_price * self.commission_rate
                        sell_tax = info['shares'] * sell_price * self.tax_rate
                        proceeds = info['shares'] * sell_price - sell_fee - sell_tax
                        surplus_pool += proceeds

                        trades2_log.append({
                            '買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx],
                            '股票名稱': self.code_to_name[self.assets[a_idx]],
                            'T+1日買進價格': info['entry_price'], '股數': info['shares'],
                            '賣出訊號日期': date, 'T+1日賣出價格': sell_price, '損益': proceeds - info['budget'],
                            '報酬率': (proceeds / info['budget']) - 1, '買進原因': info['entry_reason'], '賣出原因': reason_str
                        })
                        trades_log.append({
                            '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出',
                            '價格': sell_price, '股數': info['shares'], '動能值': f"{roc[i][a_idx]*100:.2f}%",
                            '標的名稱': self.code_to_name[self.assets[a_idx]], '原因': '停損',
                            '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax,
                            '說明': f"停損賣出：{self.code_to_name[self.assets[a_idx]]}"
                        })
                        slots[s_id] = None

            # 再平衡
            if (i - start_idx) % rebalance_interval == 0:
                top_3_signals = []
                sorted_all = np.argsort(roc[i])[::-1]
                for idx in sorted_all:
                    if len(top_3_signals) >= 3: break
                    p, s, r = current_prices[idx], sma[i][idx], roc[i][idx]
                    amount = p * self.volumes[i][idx] * 1000
                    c2 = p > sma5[i][idx] and p > sma10[i][idx] and p > sma20[i][idx]
                    if p > s and r > 0 and amount > 30000000 and c2:
                        top_3_signals.append(idx)

                signal_to_slot_map = {}
                for s_id, info in slots.items():
                    if info and 'asset_idx' in info:
                        if info['asset_idx'] in top_3_signals:
                            signal_to_slot_map[info['asset_idx']] = s_id
                        else:
                            # 排名外賣出
                            a_idx = info['asset_idx']
                            sell_price = next_prices[a_idx]
                            sell_fee = info['shares'] * sell_price * self.commission_rate
                            sell_tax = info['shares'] * sell_price * self.tax_rate
                            proceeds = info['shares'] * sell_price - sell_fee - sell_tax
                            surplus_pool += proceeds

                            trades2_log.append({
                                '買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx],
                                '股票名稱': self.code_to_name[self.assets[a_idx]],
                                'T+1日買進價格': info['entry_price'], '股數': info['shares'],
                                '賣出訊號日期': date, 'T+1日賣出價格': sell_price, '損益': proceeds - info['budget'],
                                '報酬率': (proceeds / info['budget']) - 1, '買進原因': info['entry_reason'],
                                '賣出原因': f"再平衡賣出：排名外"
                            })
                            trades_log.append({
                                '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出',
                                '價格': sell_price, '股數': info['shares'], '動能值': f"{roc[i][a_idx]*100:.2f}%",
                                '標的名稱': self.code_to_name[self.assets[a_idx]], '原因': '再平衡',
                                '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax,
                                '說明': f"再平衡賣出：{self.code_to_name[self.assets[a_idx]]}"
                            })
                            slots[s_id] = None

                # 買進新標的
                for s_id in slots:
                    if slots[s_id] is None and top_3_signals:
                        # 找下一個不在槽位中的訊號
                        next_sig = None
                        for sig in top_3_signals:
                            if sig not in [slots[sid]['asset_idx'] for sid in slots if slots[sid]]:
                                next_sig = sig
                                break

                        if next_sig is not None:
                            # 固定投入金額 10M
                            budget = 10000000.0
                            buy_price_exec = next_prices[next_sig]
                            shares = (int(budget // (buy_price_exec * (1 + self.commission_rate))) // 1000) * 1000
                            if shares > 0:
                                actual_cost = shares * buy_price_exec * (1 + self.commission_rate)
                                buy_fee = shares * buy_price_exec * self.commission_rate
                                surplus_pool -= actual_cost
                                slots[s_id] = {
                                    'asset_idx': next_sig, 'shares': shares, 'max_price': buy_price_exec,
                                    'budget': actual_cost, 'entry_date': date, 'entry_price': buy_price_exec,
                                    'entry_reason': f"符合趨勢與濾網，ROC:{roc[i][next_sig]*100:.2f}%"
                                }
                                trades_log.append({
                                    '訊號日期': date, '股票代號': self.assets[next_sig], '狀態': '買進',
                                    '價格': buy_price_exec, '股數': shares, '動能值': f"{roc[i][next_sig]*100:.2f}%",
                                    '標的名稱': self.code_to_name[self.assets[next_sig]], '原因': '符合趨勢',
                                    '買入手續費': buy_fee, '賣出手續費': 0, '賣出交易稅': 0,
                                    '說明': f"買進：{self.code_to_name[self.assets[next_sig]]}"
                                })

                # 續抱日誌
                for sig, s_id in signal_to_slot_map.items():
                    trades_log.append({
                        '訊號日期': date, '股票代號': self.assets[sig], '狀態': '保持',
                        '價格': current_prices[sig], '股數': slots[s_id]['shares'],
                        '動能值': f"{roc[i][sig]*100:.2f}%", '標的名稱': self.code_to_name[self.assets[sig]], '原因': '趨勢持續',
                        '買入手續費': 0, '賣出手續費': 0, '賣出交易稅': 0, '說明': f"續抱：{self.code_to_name[self.assets[sig]]}"
                    })

            # Snapshot
            count = sum(1 for s in slots.values() if s and 'asset_idx' in s)
            holdings_history.append({
                'Date': date, 'Holdings': ", ".join(h_names), 'Count': count, '現金': surplus_pool, '股票市值': stock_mv, '總資產': total_equity
            })

        return pd.DataFrame(equity_curve_data), pd.DataFrame(trades_log), pd.DataFrame(holdings_history), pd.DataFrame(trades2_log), pd.DataFrame(daily_details)

def calculate_metrics(equity_curve_df):
    if equity_curve_df.empty: return 0, 0, 0, 0
    equity = equity_curve_df['權益']
    total_return = (equity.iloc[-1] / equity.iloc[0]) - 1
    days = (equity_curve_df['日期'].iloc[-1] - equity_curve_df['日期'].iloc[0]).days
    years = days / 365.25
    cagr = (1 + total_return) ** (1 / years) - 1 if years > 0 else 0
    max_dd = equity_curve_df['回撤(Drawdown)'].min()
    calmar = cagr / abs(max_dd) if max_dd != 0 else 0
    return cagr, max_dd, calmar, total_return
