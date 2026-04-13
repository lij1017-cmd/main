import pandas as pd
import numpy as np

def clean_data(filepath):
    """
    清洗並預處理輸入的 Excel 資料檔。
    此函數會讀取 '還原收盤價' 與 '成交量' 工作表，並處理時間序列。
    """
    # 讀取 Excel 檔案，header=None 表示不使用預設標頭
    df_prices = pd.read_excel(filepath, sheet_name='還原收盤價', header=None)
    df_volume = pd.read_excel(filepath, sheet_name='成交量', header=None)

    # 提取第 0 列為股票代號，第 1 列為股票名稱 (排除第 0 欄的日期列)
    stock_codes = df_prices.iloc[0, 1:].values
    stock_names = df_prices.iloc[1, 1:].values

    # 提取日期資訊 (從第 2 列開始)，格式如 '20190102收盤價'，擷取前 8 位轉為日期物件
    date_strings = df_prices.iloc[2:, 0].astype(str).str[:8]
    dates = pd.to_datetime(date_strings, format='%Y%m%d')

    # 提取價格數據 (從第 2 列、第 1 欄開始)，轉為浮點數並設定索引與欄名
    prices = df_prices.iloc[2:, 1:].astype(float)
    prices.index = dates
    prices.columns = stock_codes

    # 提取成交量數據 (從第 2 列、第 1 欄開始)，轉為浮點數並設定索引與欄名
    volumes = df_volume.iloc[2:, 1:].astype(float)
    volumes.index = dates
    volumes.columns = stock_codes

    # 建立股票代號對應名稱的字典，方便後續查詢
    code_to_name = dict(zip(stock_codes, stock_names))

    # 數據補齊邏輯：價格採前值補後值 (ffill)，若開頭即缺失則採後值補前值 (bfill)
    prices = prices.ffill().bfill()
    # 成交量缺失則補 0
    volumes = volumes.fillna(0)

    return prices, volumes, code_to_name

class BacktesterV2:
    """
    回測引擎：動態分配版 V2 (支援 2026/04/01 選股池自動擴充，並新增每日操作備註)。
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

    def run(self, sma_period, roc_period, stop_loss_pct, rebalance_interval=6, stop_loss_type='peak', ma_stop_period=5):
        # 1. 指標預計算階段
        sma = self.prices_df.rolling(window=sma_period).mean().values
        roc = self.prices_df.pct_change(periods=roc_period).values
        sma5 = self.prices_df.rolling(window=5).mean().values
        sma10 = self.prices_df.rolling(window=10).mean().values
        sma20 = self.prices_df.rolling(window=20).mean().values

        # 2. 回測帳戶與持股設定初始化
        surplus_pool = float(self.initial_capital)
        slots = {0: None, 1: None, 2: None}

        equity_curve_data = []
        trades_log = []
        trades2_log = []
        holdings_history = []
        daily_details = []

        peak_equity = float(self.initial_capital)
        start_idx = max(sma_period, roc_period, 20)

        # 遍歷每一日進行模擬回測
        for i in range(start_idx, len(self.dates)):
            date = self.dates[i]
            current_prices = self.prices[i]
            daily_remarks = [] # 用於記錄當日的特殊操作備註

            # A. 計算今日帳戶權益 (按今日收盤價估值)
            stock_mv = 0.0
            h_names = []
            for s_id, info in slots.items():
                if info and 'asset_idx' in info:
                    a_idx = info['asset_idx']
                    mv = info['shares'] * current_prices[a_idx]
                    stock_mv += mv
                    h_names.append(f"{self.code_to_name[self.assets[a_idx]]}({self.assets[a_idx]})")
                    daily_details.append({
                        '日期': date, '股票代號': self.assets[a_idx],
                        '股票名稱': self.code_to_name[self.assets[a_idx]],
                        '持有股數': info['shares'], '本日收盤價': current_prices[a_idx], '市值': mv
                    })

            total_equity = surplus_pool + stock_mv
            if total_equity > peak_equity: peak_equity = total_equity
            drawdown = (total_equity - peak_equity) / peak_equity

            equity_curve_data.append({
                '日期': date, '權益': total_equity, '回撤(Drawdown)': drawdown
            })

            # B. 每日停損檢查邏輯 (訊號產生於今日收盤，執行於明日)
            triggered_slots = []
            for s_id, info in slots.items():
                if info and 'asset_idx' in info:
                    a_idx = info['asset_idx']
                    curr_p = current_prices[a_idx]
                    if stop_loss_type == 'peak':
                        if curr_p > info['max_price']: info['max_price'] = curr_p
                        if curr_p < info['max_price'] * (1 - stop_loss_pct):
                            triggered_slots.append((s_id, f"觸發停損：價格由高點回落 {stop_loss_pct*100}%"))
                            daily_remarks.append(f"今日觸發 {self.code_to_name[self.assets[a_idx]]}({self.assets[a_idx]}) 停損機制出場。")

            # C. 再平衡選股篩選邏輯 (每 rebalance_interval 天執行一次)
            is_rebalance_day = (i - start_idx) % rebalance_interval == 0

            if is_rebalance_day:
                # 動態選股池
                if date < pd.Timestamp('2026-04-01'): pool_indices = list(range(131))
                else: pool_indices = list(range(138))

                top_3_signals = []
                pool_roc = roc[i][pool_indices]
                sorted_pool_idx = np.argsort(pool_roc)[::-1]
                sorted_all = [pool_indices[idx] for idx in sorted_pool_idx]

                for idx in sorted_all:
                    if len(top_3_signals) >= 3: break
                    p, s, r = current_prices[idx], sma[i][idx], roc[i][idx]
                    v = self.volumes[i][idx]
                    amount = p * v * 1000
                    cond1 = amount > 30000000
                    cond2 = p > sma5[i][idx] and p > sma10[i][idx] and p > sma20[i][idx]
                    if p > s and r > 0 and cond1 and cond2:
                        top_3_signals.append(idx)

                # 產生指令備註
                current_assets_in_slots = [info['asset_idx'] for info in slots.values() if info and 'asset_idx' in info]

                to_sell = []
                for s_id, info in slots.items():
                    if info and 'asset_idx' in info and info['asset_idx'] not in top_3_signals:
                        to_sell.append(f"{self.code_to_name[self.assets[info['asset_idx']]]}({self.assets[info['asset_idx']]})")

                to_buy = []
                for sig in top_3_signals:
                    if sig not in current_assets_in_slots:
                        to_buy.append(f"{self.code_to_name[self.assets[sig]]}({self.assets[sig]})")

                to_hold = []
                for sig in top_3_signals:
                    if sig in current_assets_in_slots:
                        to_hold.append(f"{self.code_to_name[self.assets[sig]]}({self.assets[sig]})")

                action_str = "再平衡日，建議次日交易指令："
                if to_buy: action_str += f"【買進】{', '.join(to_buy)} "
                if to_sell: action_str += f"【賣出】{', '.join(to_sell)} "
                if to_hold: action_str += f"【續抱】{', '.join(to_hold)} "
                if not to_buy and not to_sell and not to_hold: action_str += "無符合標的，建議空倉。"
                daily_remarks.append(action_str)

            # D. 記錄持股總覽 (含備註)
            holding_count = sum(1 for s in slots.values() if s and 'asset_idx' in s)
            holdings_history.append({
                'Date': date,
                'Holdings': ", ".join(h_names),
                'Count': holding_count,
                '現金': surplus_pool,
                '股票市值': stock_mv,
                '總資產': total_equity,
                '備註': "；".join(daily_remarks)
            })

            # 如果是最後一天，雖然無法執行明日交易，但已產出今日訊號並記錄於備註
            if i == len(self.dates) - 1:
                break

            # E. 執行交易 (T+1 執行)
            next_prices = self.prices[i+1]

            # 1. 執行停損賣出
            for s_id, reason_str in triggered_slots:
                if slots[s_id]: # 確保槽位還在
                    info = slots[s_id]
                    a_idx = info['asset_idx']
                    sell_price = next_prices[a_idx]
                    shares = info['shares']
                    sell_fee = shares * sell_price * 0.001425
                    sell_tax = shares * sell_price * 0.003
                    proceeds = shares * sell_price - sell_fee - sell_tax
                    surplus_pool += proceeds

                    name = self.code_to_name[self.assets[a_idx]]
                    pnl = proceeds - info['budget']
                    ret_pct = (proceeds / info['budget']) - 1
                    trades2_log.append({
                        '買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx],
                        '股票名稱': name, 'T+1日買進價格': info['entry_price'], '股數': shares,
                        '賣出訊號日期': date, 'T+1日賣出價格': sell_price, '損益': pnl,
                        '報酬率': ret_pct, '買進原因': info['entry_reason'], '賣出原因': reason_str
                    })
                    trades_log.append({
                        '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出',
                        '價格': sell_price, '股數': shares, '動能值': f"{roc[i][a_idx]*100:.2f}%",
                        '標的名稱': name, '原因': '停損', '買入手續費': 0, '賣出手續費': sell_fee,
                        '賣出交易稅': sell_tax, '說明': f"停損賣出：{name} ({reason_str})"
                    })
                    slots[s_id] = None

            # 2. 執行再平衡買賣
            if is_rebalance_day:
                signal_to_slot_map = {}
                for s_id, info in slots.items():
                    if info and 'asset_idx' in info:
                        if info['asset_idx'] in top_3_signals:
                            signal_to_slot_map[info['asset_idx']] = s_id
                        else:
                            a_idx = info['asset_idx']
                            sell_price = next_prices[a_idx]
                            shares = info['shares']
                            sell_fee = shares * sell_price * 0.001425
                            sell_tax = shares * sell_price * 0.003
                            proceeds = shares * sell_price - sell_fee - sell_tax

                            name = self.code_to_name[self.assets[a_idx]]
                            pnl = proceeds - info['budget']
                            ret_pct = (proceeds / info['budget']) - 1
                            trades2_log.append({
                                '買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx],
                                '股票名稱': name, 'T+1日買進價格': info['entry_price'], '股數': shares,
                                '賣出訊號日期': date, 'T+1日賣出價格': sell_price, '損益': pnl,
                                '報酬率': ret_pct, '買進原因': info['entry_reason'], '賣出原因': "再平衡賣出"
                            })
                            trades_log.append({
                                '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出',
                                '價格': sell_price, '股數': shares, '動能值': f"{roc[i][a_idx]*100:.2f}%",
                                '標的名稱': name, '原因': '再平衡', '買入手續費': 0, '賣出手續費': sell_fee,
                                '賣出交易稅': sell_tax, '說明': f"再平衡賣出：{name}"
                            })
                            slots[s_id] = {'pending_budget': proceeds}
                    else:
                        alloc = min(surplus_pool, 10000000.0)
                        surplus_pool -= alloc
                        slots[s_id] = {'pending_budget': alloc}

                new_signals = [sig for sig in top_3_signals if sig not in signal_to_slot_map]
                available_slot_ids = [sid for sid, data in slots.items() if data and 'pending_budget' in data]

                for sig in new_signals:
                    if not available_slot_ids: break
                    target_sid = available_slot_ids.pop(0)
                    budget = slots[target_sid]['pending_budget']
                    invest_budget = min(budget, 10000000.0)
                    buy_price_exec = next_prices[sig]
                    cost_per_share = buy_price_exec * 1.001425
                    shares = (int(invest_budget // cost_per_share) // 1000) * 1000
                    if shares > 0:
                        actual_cost = shares * buy_price_exec * 1.001425
                        buy_fee = shares * buy_price_exec * 0.001425
                        surplus_pool += (budget - actual_cost)
                        name = self.code_to_name[self.assets[sig]]
                        slots[target_sid] = {
                            'asset_idx': sig, 'shares': shares, 'max_price': buy_price_exec,
                            'budget': actual_cost, 'entry_date': date, 'entry_price': buy_price_exec,
                            'entry_reason': f"動能符合, ROC:{roc[i][sig]*100:.2f}%"
                        }
                        trades_log.append({
                            '訊號日期': date, '股票代號': self.assets[sig], '狀態': '買進',
                            '價格': buy_price_exec, '股數': shares, '動能值': f"{roc[i][sig]*100:.2f}%",
                            '標的名稱': name, '原因': '再平衡買入', '買入手續費': buy_fee, '賣出手續費': 0,
                            '賣出交易稅': 0, '說明': f"買進標的：{name}"
                        })
                    else:
                        surplus_pool += budget
                        slots[target_sid] = None

                for sid in list(slots.keys()):
                    if slots[sid] and 'pending_budget' in slots[sid]:
                        surplus_pool += slots[sid]['pending_budget']
                        slots[sid] = None
                for sig, sid in signal_to_slot_map.items():
                    name = self.code_to_name[self.assets[sig]]
                    trades_log.append({
                        '訊號日期': date, '股票代號': self.assets[sig], '狀態': '保持',
                        '價格': current_prices[sig], '股數': slots[sid]['shares'],
                        '動能值': f"{roc[i][sig]*100:.2f}%", '標的名稱': name, '原因': '續抱',
                        '買入手續費': 0, '賣出手續費': 0, '賣出交易稅': 0, '說明': f"保持持有：{name}"
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
