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
    回測引擎：動態分配版 V2 (支援 2026/04/01 選股池自動擴充)。
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
        # 計算簡單移動平均線 (SMA) 作為趨勢判讀基礎
        sma = self.prices_df.rolling(window=sma_period).mean().values
        # 計算變動率 (ROC) 作為動能強弱判讀
        roc = self.prices_df.pct_change(periods=roc_period).values

        # 計算短、中、長期移動平均線 (5, 10, 20日)，用於多重均線濾網
        sma5 = self.prices_df.rolling(window=5).mean().values
        sma10 = self.prices_df.rolling(window=10).mean().values
        sma20 = self.prices_df.rolling(window=20).mean().values

        # 2. 回測帳戶與持股設定初始化
        surplus_pool = float(self.initial_capital) # 目前可用現金
        # slots 為 3 個持股槽位，每個槽位固定配置上限 1000 萬，若有空缺則找新標的
        slots = {0: None, 1: None, 2: None}

        equity_curve_data = [] # 權益曲線數據集
        trades_log = []        # 訊號發出紀錄
        trades2_log = []       # 買賣對帳明細 (計算單筆損益)
        holdings_history = []  # 每日持股清單與資金總結
        daily_details = []     # 每日持股明細 (含股數、市值)

        current_reasons = []   # 記錄當下決策原因
        peak_equity = float(self.initial_capital) # 歷史最高淨值，用於計算回撤

        # 設定回測起始日，需確保所有指標 (包含長週期 SMA) 都有數值
        start_idx = max(sma_period, roc_period, 20)

        # 遍歷每一日進行模擬回測
        for i in range(start_idx, len(self.dates)):
            date = self.dates[i]
            current_prices = self.prices[i]

            # A. 計算今日帳戶權益 (按今日收盤價估值)
            stock_mv = 0.0
            h_names = []
            for s_id, info in slots.items():
                if info and 'asset_idx' in info:
                    a_idx = info['asset_idx']
                    # 市值 = 持有股數 * 今日收盤價
                    mv = info['shares'] * current_prices[a_idx]
                    stock_mv += mv
                    h_names.append(f"{self.code_to_name[self.assets[a_idx]]}({self.assets[a_idx]})")
                    # 記錄持股明細
                    daily_details.append({
                        '日期': date,
                        '股票代號': self.assets[a_idx],
                        '股票名稱': self.code_to_name[self.assets[a_idx]],
                        '持有股數': info['shares'],
                        '本日收盤價': current_prices[a_idx],
                        '市值': mv
                    })

            # 總資產 = 現金池 + 股票總市值
            total_equity = surplus_pool + stock_mv
            # 更新歷史最高淨值並計算當前最大回撤 (Drawdown)
            if total_equity > peak_equity: peak_equity = total_equity
            drawdown = (total_equity - peak_equity) / peak_equity

            # 記錄每日權益曲線
            equity_curve_data.append({
                '日期': date,
                '權益': total_equity,
                '回撤(Drawdown)': drawdown
            })

            # 模擬結束
            if i == len(self.dates) - 1:
                break

            # 交易執行日價格：本策略採 T 日訊號，T+1 日收盤價執行
            next_prices = self.prices[i+1]

            # B. 每日停損檢查邏輯
            triggered_slots = []
            for s_id, info in slots.items():
                if info and 'asset_idx' in info:
                    a_idx = info['asset_idx']
                    curr_p = current_prices[a_idx]

                    stop_triggered = False
                    reason_str = ""

                    # 採用最高價回落停損規則 (Peak Stop Loss)
                    if stop_loss_type == 'peak':
                        # 更新持股期間最高價
                        if curr_p > info['max_price']:
                            info['max_price'] = curr_p
                        # 檢查回落比例
                        if curr_p < info['max_price'] * (1 - stop_loss_pct):
                            stop_triggered = True
                            reason_str = f"觸發停損：價格由波段高點回落 {stop_loss_pct*100}%"

                    if stop_triggered:
                        triggered_slots.append((s_id, reason_str))

            # 執行停損部位賣出 (於 T+1 日執行)
            for s_id, reason_str in triggered_slots:
                info = slots[s_id]
                a_idx = info['asset_idx']
                sell_price = next_prices[a_idx]
                shares = info['shares']

                # 扣除手續費 (0.1425%) 與交易稅 (0.3%)
                sell_fee = shares * sell_price * 0.001425
                sell_tax = shares * sell_price * 0.003
                proceeds = shares * sell_price - sell_fee - sell_tax
                surplus_pool += proceeds # 資金回籠至現金池

                name = self.code_to_name[self.assets[a_idx]]
                current_reasons.append(f"{name} 停損出場")

                # 記錄單筆成對交易
                pnl = proceeds - info['budget']
                ret_pct = (proceeds / info['budget']) - 1
                trades2_log.append({
                    '買進訊號日期': info['entry_date'],
                    '股票代號': self.assets[a_idx],
                    '股票名稱': name,
                    'T+1日買進價格': info['entry_price'],
                    '股數': shares,
                    '賣出訊號日期': date,
                    'T+1日賣出價格': sell_price,
                    '損益': pnl,
                    '報酬率': ret_pct,
                    '買進原因': info['entry_reason'],
                    '賣出原因': reason_str
                })

                # 記錄交易流水帳
                trades_log.append({
                    '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出',
                    '價格': sell_price, '股數': shares, '動能值': f"{roc[i][a_idx]*100:.2f}%",
                    '標的名稱': name, '原因': '停損',
                    '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax,
                    '說明': f"停損賣出：{name} ({reason_str})"
                })
                slots[s_id] = None # 清空槽位

            # C. 再平衡與選股篩選邏輯 (每 Rebalance Interval 天執行一次)
            is_rebalance_day = (i - start_idx) % rebalance_interval == 0
            if is_rebalance_day:
                current_reasons = []

                # --- 動態選股池調整邏輯 ---
                # 依據 MD 檔規範：2026/04/01 之前維持 131 檔，之後擴充為 138 檔。
                if date < pd.Timestamp('2026-04-01'):
                    pool_indices = list(range(131)) # 原始標的
                else:
                    pool_indices = list(range(138)) # 包含新納入標的

                top_3_signals = [] # 本期入選標的索引
                exclusion_reasons = [] # 排除原因紀錄

                # 在選股池範圍內，依 ROC (動能) 由高至低排序
                pool_roc = roc[i][pool_indices]
                sorted_pool_idx = np.argsort(pool_roc)[::-1]
                sorted_all = [pool_indices[idx] for idx in sorted_pool_idx]

                for idx in sorted_all:
                    if len(top_3_signals) >= 3: break # 只選前三名
                    p, s, r = current_prices[idx], sma[i][idx], roc[i][idx]
                    v = self.volumes[i][idx]
                    amount = p * v * 1000 # 日成交金額 (TWD)

                    # 濾網一：成交金額必須大於 3,000 萬 TWD
                    cond1 = amount > 30000000
                    # 濾網二：價格必須高於 5, 10, 20 日移動平均線 (短中長期多頭)
                    cond2 = p > sma5[i][idx] and p > sma10[i][idx] and p > sma20[i][idx]

                    # 核心訊號條件：價格高於 SMA 且 ROC 為正，且通過以上兩項濾網
                    if p > s and r > 0 and cond1 and cond2:
                        top_3_signals.append(idx)
                    else:
                        # 紀錄不符原因供除錯參考
                        name = self.code_to_name[self.assets[idx]]
                        if r <= 0: exclusion_reasons.append(f"{name} ROC 非正")
                        elif p <= s: exclusion_reasons.append(f"{name} 價格低於 SMA")
                        elif not cond1: exclusion_reasons.append(f"{name} 流動性不足")
                        elif not cond2: exclusion_reasons.append(f"{name} 均線濾網未過")

                # 處理現有持股：若入選前三則續抱，否則賣出釋出資金
                signal_to_slot_map = {}
                for s_id, info in slots.items():
                    if info and 'asset_idx' in info:
                        if info['asset_idx'] in top_3_signals:
                            signal_to_slot_map[info['asset_idx']] = s_id # 標的仍在名單內，紀錄槽位
                        else:
                            # 標的不在名單內，執行再平衡賣出
                            a_idx = info['asset_idx']
                            sell_price = next_prices[a_idx]
                            shares = info['shares']
                            sell_fee = shares * sell_price * 0.001425
                            sell_tax = shares * sell_price * 0.003
                            proceeds = shares * sell_price - sell_fee - sell_tax

                            name = self.code_to_name[self.assets[a_idx]]
                            sell_reason = "再平衡賣出：排名落後或未過濾網"

                            # 紀錄成對交易損益
                            pnl = proceeds - info['budget']
                            ret_pct = (proceeds / info['budget']) - 1
                            trades2_log.append({
                                '買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx],
                                '股票名稱': name, 'T+1日買進價格': info['entry_price'], '股數': shares,
                                '賣出訊號日期': date, 'T+1日賣出價格': sell_price, '損益': pnl,
                                '報酬率': ret_pct, '買進原因': info['entry_reason'], '賣出原因': sell_reason
                            })

                            # 紀錄流水帳
                            trades_log.append({
                                '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出',
                                '價格': sell_price, '股數': shares, '動能值': f"{roc[i][a_idx]*100:.2f}%",
                                '標的名稱': name, '原因': '再平衡',
                                '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax,
                                '說明': f"再平衡賣出：{name}"
                            })
                            slots[s_id] = {'pending_budget': proceeds} # 暫存資金供買入新標的
                    else:
                        # 閒置槽位分配資金 (上限 1000 萬)
                        alloc = min(surplus_pool, 10000000.0)
                        surplus_pool -= alloc
                        slots[s_id] = {'pending_budget': alloc}

                # 買入新入選標的
                new_signals = [sig for sig in top_3_signals if sig not in signal_to_slot_map]
                available_slot_ids = [sid for sid, data in slots.items() if data and 'pending_budget' in data]

                for sig in new_signals:
                    if not available_slot_ids: break # 若無空缺槽位則略過
                    target_sid = available_slot_ids.pop(0)
                    budget = slots[target_sid]['pending_budget']
                    invest_budget = min(budget, 10000000.0) # 單一標的上限制
                    buy_price_exec = next_prices[sig]
                    cost_per_share = buy_price_exec * 1.001425 # 含手續費
                    # 股數計算：必須為 1,000 股 (一張) 的整數倍
                    shares = (int(invest_budget // cost_per_share) // 1000) * 1000

                    if shares > 0:
                        actual_cost = shares * buy_price_exec * 1.001425
                        buy_fee = shares * buy_price_exec * 0.001425
                        surplus_pool += (budget - actual_cost) # 找零存回現金池

                        name = self.code_to_name[self.assets[sig]]
                        entry_reason = f"動能訊號符合，ROC: {roc[i][sig]*100:.2f}%"

                        # 更新槽位狀態為持有中
                        slots[target_sid] = {
                            'asset_idx': sig, 'shares': shares, 'max_price': buy_price_exec,
                            'budget': actual_cost, 'entry_date': date, 'entry_price': buy_price_exec,
                            'entry_reason': entry_reason
                        }

                        # 紀錄流水帳
                        trades_log.append({
                            '訊號日期': date, '股票代號': self.assets[sig], '狀態': '買進',
                            '價格': buy_price_exec, '股數': shares, '動能值': f"{roc[i][sig]*100:.2f}%",
                            '標的名稱': name, '原因': '符合趨勢',
                            '買入手續費': buy_fee, '賣出手續費': 0, '賣出交易稅': 0,
                            '說明': f"買進標的：{name}"
                        })
                    else:
                        # 資金不足以買入一張，資金歸還
                        surplus_pool += budget
                        slots[target_sid] = None

                # 清理未使用的暫存預算，並記錄「保持」持有狀態的部位
                for sid in list(slots.keys()):
                    if slots[sid] and 'pending_budget' in slots[sid]:
                        surplus_pool += slots[sid]['pending_budget']
                        slots[sid] = None
                for sig, sid in signal_to_slot_map.items():
                    name = self.code_to_name[self.assets[sig]]
                    trades_log.append({
                        '訊號日期': date, '股票代號': self.assets[sig], '狀態': '保持',
                        '價格': current_prices[sig], '股數': slots[sid]['shares'],
                        '動能值': f"{roc[i][sig]*100:.2f}%", '標的名稱': name, '原因': '趨勢持續',
                        '買入手續費': 0, '賣出手續費': 0, '賣出交易稅': 0, '說明': f"保持持有：{name}"
                    })

            # D. 每日總結紀錄
            count = sum(1 for s in slots.values() if s and 'asset_idx' in s)
            holdings_history.append({
                'Date': date,
                'Holdings': ", ".join(h_names),
                'Count': count,
                '現金': surplus_pool,
                '股票市值': stock_mv,
                '總資產': total_equity
            })

        # 回傳回測結果的各種 DataFrame 資料表
        return pd.DataFrame(equity_curve_data), pd.DataFrame(trades_log), pd.DataFrame(holdings_history), pd.DataFrame(trades2_log), pd.DataFrame(daily_details)

def calculate_metrics(equity_curve_df):
    """
    計算策略關鍵績效指標。
    輸入為包含日期與權益的資料表。
    """
    if equity_curve_df.empty: return 0, 0, 0, 0
    equity = equity_curve_df['權益']
    total_return = (equity.iloc[-1] / equity.iloc[0]) - 1
    # 計算回測天數與換算年份
    days = (equity_curve_df['日期'].iloc[-1] - equity_curve_df['日期'].iloc[0]).days
    years = days / 365.25
    # CAGR 年化報酬率 (複利公式)
    cagr = (1 + total_return) ** (1 / years) - 1 if years > 0 else 0
    # 最大回撤 (Max Drawdown)
    max_dd = equity_curve_df['回撤(Drawdown)'].min()
    # 卡瑪比率 (Calmar Ratio)
    calmar = cagr / abs(max_dd) if max_dd != 0 else 0
    return cagr, max_dd, calmar, total_return
