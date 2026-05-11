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

class BacktesterBreadth:
    """
    回測引擎：加入市場寬度與趨勢雙重確認濾網 (方案 B)。
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

    def run(self, sma_period, roc_period, stop_loss_pct, rebalance_interval=9, use_market_filter=True, breadth_threshold=0.35, mkt_sma_window=20):
        # 1. 指標預計算
        sma = self.prices_df.rolling(window=sma_period).mean().values
        roc = self.prices_df.pct_change(periods=roc_period).values
        sma5 = self.prices_df.rolling(window=5).mean().values
        sma10 = self.prices_df.rolling(window=10).mean().values
        sma20 = self.prices_df.rolling(window=20).mean().values

        # 市場濾網：方案 B (寬度 35% OR 大盤 SMA 20)
        b200_all = self.prices_df.rolling(window=200).mean().values
        breadth = np.mean(self.prices > b200_all, axis=1)

        market_avg = self.prices_df.mean(axis=1).values
        market_sma = self.prices_df.mean(axis=1).rolling(window=mkt_sma_window).mean().values

        # 邏輯：滿足寬度門檻 或 大盤確認趨勢 (OR Logic)
        mkt_filter = (breadth >= breadth_threshold) | (market_avg >= market_sma)

        # 2. 帳戶與槽位初始化
        surplus_pool = float(self.initial_capital)
        slots = {0: None, 1: None, 2: None}

        equity_curve_data = []
        trades_log = []
        trades2_log = []
        holdings_history = []
        daily_details = []

        current_reasons = []
        peak_equity = float(self.initial_capital)

        start_idx = max(sma_period, roc_period, 200, mkt_sma_window)

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

            # B. 每日檢查市場濾網 (全清倉)
            if use_market_filter and not mkt_filter[i]:
                triggered_slots = []
                for s_id, info in slots.items():
                    if info and 'asset_idx' in info:
                        triggered_slots.append((s_id, f"雙重確認濾網觸發：寬度({breadth[i]:.1%})與大盤皆弱"))

                if triggered_slots:
                    for s_id, reason_str in triggered_slots:
                        info = slots[s_id]
                        a_idx = info['asset_idx']
                        sell_price = next_prices[a_idx]
                        sell_fee = info['shares'] * sell_price * 0.001425
                        sell_tax = info['shares'] * sell_price * 0.003
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
                            '標的名稱': self.code_to_name[self.assets[a_idx]], '原因': '市場濾網',
                            '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax,
                            '說明': f"市場濾網賣出：{self.code_to_name[self.assets[a_idx]]}"
                        })
                        slots[s_id] = None
                continue

            # C. 每日檢查停損
            triggered_sl = []
            for s_id, info in slots.items():
                if info and 'asset_idx' in info:
                    a_idx = info['asset_idx']
                    if current_prices[a_idx] > info['max_price']: info['max_price'] = current_prices[a_idx]
                    if current_prices[a_idx] < info['max_price'] * (1 - stop_loss_pct):
                        triggered_sl.append((s_id, f"停損機制：價格自最高點回落達{stop_loss_pct*100:.2f}%"))

            for s_id, reason_str in triggered_sl:
                info = slots[s_id]
                a_idx = info['asset_idx']
                sell_price = next_prices[a_idx]
                sell_fee = info['shares'] * sell_price * 0.001425
                sell_tax = info['shares'] * sell_price * 0.003
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

            # D. 再平衡
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
                            a_idx = info['asset_idx']
                            sell_price = next_prices[a_idx]
                            sell_fee = info['shares'] * sell_price * 0.001425
                            sell_tax = info['shares'] * sell_price * 0.003
                            proceeds = info['shares'] * sell_price - sell_fee - sell_tax

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
                            slots[s_id] = {'pending_budget': proceeds}
                    else:
                        alloc = min(surplus_pool, 10000000.0)
                        surplus_pool -= alloc
                        slots[s_id] = {'pending_budget': alloc}

                new_signals = [sig for sig in top_3_signals if sig not in signal_to_slot_map]
                available_ids = [sid for sid, data in slots.items() if data and 'pending_budget' in data]
                for sig in new_signals:
                    if not available_ids: break
                    target_sid = available_ids.pop(0)
                    budget = slots[target_sid]['pending_budget']
                    buy_price_exec = next_prices[sig]
                    shares = (int(budget // (buy_price_exec * 1.001425)) // 1000) * 1000
                    if shares > 0:
                        actual_cost = shares * buy_price_exec * 1.001425
                        buy_fee = shares * buy_price_exec * 0.001425
                        surplus_pool += (budget - actual_cost)
                        slots[target_sid] = {
                            'asset_idx': sig, 'shares': shares, 'max_price': buy_price_exec,
                            'budget': actual_cost, 'entry_date': date, 'entry_price': buy_price_exec,
                            'entry_reason': f"符合趨勢與濾網，ROC:{roc[i][sig]*100:.2f}%"
                        }
                        trades_log.append({
                            '訊號日期': date, '股票代號': self.assets[sig], '狀態': '買進',
                            '價格': buy_price_exec, '股數': shares, '動能值': f"{roc[i][sig]*100:.2f}%",
                            '標的名稱': self.code_to_name[self.assets[sig]], '原因': '符合趨勢',
                            '買入手續費': buy_fee, '賣出手續費': 0, '賣出交易稅': 0,
                            '說明': f"買進：{self.code_to_name[self.assets[sig]]}"
                        })
                    else:
                        surplus_pool += budget
                        slots[target_sid] = None

                for sid in list(slots.keys()):
                    if slots[sid] and 'pending_budget' in slots[sid]:
                        surplus_pool += slots[sid]['pending_budget']
                        slots[sid] = None
                for sig, sid in signal_to_slot_map.items():
                    trades_log.append({
                        '訊號日期': date, '股票代號': self.assets[sig], '狀態': '保持',
                        '價格': current_prices[sig], '股數': slots[sid]['shares'],
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
