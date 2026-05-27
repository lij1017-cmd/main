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

class BacktesterAdjusted:
    """
    回測引擎：調整版 (equityV(調))。
    固定使用資金 1200 萬，授權資金 6000 萬。
    """
    def __init__(self, prices, volumes, code_to_name, authorized_capital=60000000, trading_capital=12000000):
        self.prices_df = prices
        self.volumes_df = volumes
        self.prices = prices.values
        self.volumes = volumes.values
        self.dates = prices.index
        self.assets = prices.columns
        self.code_to_name = code_to_name
        self.authorized_capital = authorized_capital
        self.trading_capital = trading_capital
        self.slot_budget = trading_capital / 3 # 每檔 400 萬

    def run(self, sma_period, roc_period, stop_loss_pct, rebalance_interval=9, stop_loss_type='peak'):
        # 1. 指標預計算
        sma = self.prices_df.rolling(window=sma_period).mean().values
        roc = self.prices_df.pct_change(periods=roc_period).values
        sma5 = self.prices_df.rolling(window=5).mean().values
        sma10 = self.prices_df.rolling(window=10).mean().values
        sma20 = self.prices_df.rolling(window=20).mean().values

        # 2. 帳戶與槽位初始化
        # surplus_pool 初始化為 trading_capital
        surplus_pool = float(self.trading_capital)
        # slots: {id: {asset_idx, shares, max_price, budget, entry_date, entry_price, entry_reason, remark}}
        slots = {0: None, 1: None, 2: None}

        equity_curve_data = []
        trades_log = []
        trades2_log = []
        holdings_history = []
        daily_details = []

        current_remarks = []
        peak_equity = float(self.trading_capital)
        # 固定基準 MDD 使用 6000 萬
        # 但權益曲線的基點應該是 trading_capital 嗎？
        # 使用者要求：年初損益歸零，但持有部位延續。
        # 全期間權益曲線則持續累計。

        start_idx = max(sma_period, roc_period, 20)

        for i in range(start_idx, len(self.dates)):
            date = self.dates[i]
            current_prices = self.prices[i]

            # A. 計算今日權益與明細
            stock_mv = 0.0
            h_names = []
            for s_id, info in slots.items():
                if info and 'asset_idx' in info:
                    a_idx = info['asset_idx']
                    mv = info['shares'] * current_prices[a_idx]
                    stock_mv += mv
                    name = self.code_to_name[self.assets[a_idx]]
                    h_names.append(f"{name}({self.assets[a_idx]})")

                    daily_details.append({
                        '日期': date,
                        '股票代號': self.assets[a_idx],
                        '股票名稱': name,
                        '持有股數': info['shares'],
                        '本日收盤價': current_prices[a_idx],
                        '市值': mv,
                        '備註': info.get('remark', '')
                    })

            total_equity = surplus_pool + stock_mv
            if total_equity > peak_equity: peak_equity = total_equity

            # 標準 MDD (相對於最高權益)
            drawdown = (total_equity - peak_equity) / peak_equity if peak_equity != 0 else 0
            # 固定基準 MDD (相對於 6000 萬)
            fixed_drawdown = (total_equity - peak_equity) / self.authorized_capital

            equity_curve_data.append({
                '日期': date,
                '權益': total_equity,
                '回撤(Drawdown)': drawdown,
                '固定基準 MDD': fixed_drawdown
            })

            if i == len(self.dates) - 1:
                break

            next_prices = self.prices[i+1] # T+1 執行價格

            # B. 每日檢查停損
            triggered_slots = []
            for s_id, info in slots.items():
                if info and 'asset_idx' in info:
                    a_idx = info['asset_idx']
                    curr_p = current_prices[a_idx]

                    stop_triggered = False
                    reason_str = ""

                    if stop_loss_type == 'peak':
                        if curr_p > info['max_price']:
                            info['max_price'] = curr_p
                        if curr_p < info['max_price'] * (1 - stop_loss_pct):
                            stop_triggered = True
                            reason_str = f"停損機制，價格自最高點回落達{stop_loss_pct*100}%"

                    if stop_triggered:
                        triggered_slots.append((s_id, reason_str))

            # 執行停損賣出
            for s_id, reason_str in triggered_slots:
                info = slots[s_id]
                a_idx = info['asset_idx']
                sell_price = next_prices[a_idx]
                shares = info['shares']

                sell_fee = shares * sell_price * 0.001425
                sell_tax = shares * sell_price * 0.003
                proceeds = shares * sell_price - sell_fee - sell_tax
                surplus_pool += proceeds

                name = self.code_to_name[self.assets[a_idx]]
                current_remarks.append(f"{name}因達停損標準需剔除")

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

                trades_log.append({
                    '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出',
                    '價格': sell_price, '股數': shares, '動能值': f"{roc[i][a_idx]*100:.2f}%",
                    '標的名稱': name, '原因': '停損',
                    '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax,
                    '說明': f"停損賣出：{name} ({reason_str})"
                })
                slots[s_id] = None

            # C. 再平衡邏輯
            is_rebalance_day = (i - start_idx) % rebalance_interval == 0
            if is_rebalance_day:
                current_remarks = []

                # 篩選訊號
                top_signals = []
                sorted_indices = np.argsort(roc[i])[::-1]
                for idx in sorted_indices:
                    if len(top_signals) >= 3: break
                    p, s, r = current_prices[idx], sma[i][idx], roc[i][idx]
                    v = self.volumes[i][idx]
                    amount = p * v * 1000

                    # 濾網
                    cond1 = amount > 30000000
                    cond2 = p > sma5[i][idx] and p > sma10[i][idx] and p > sma20[i][idx]

                    if p > s and r > 0 and cond1 and cond2:
                        top_signals.append(idx)

                # 處理槽位與賣出
                signal_to_slot_map = {}
                for s_id, info in slots.items():
                    if info and 'asset_idx' in info:
                        if info['asset_idx'] in top_signals:
                            signal_to_slot_map[info['asset_idx']] = s_id # 續抱
                        else:
                            # 賣出
                            a_idx = info['asset_idx']
                            sell_price = next_prices[a_idx]
                            shares = info['shares']
                            sell_fee = shares * sell_price * 0.001425
                            sell_tax = shares * sell_price * 0.003
                            proceeds = shares * sell_price - sell_fee - sell_tax

                            name = self.code_to_name[self.assets[a_idx]]
                            sell_reason = f"再平衡賣出，排名外或不符濾網"

                            pnl = proceeds - info['budget']
                            ret_pct = (proceeds / info['budget']) - 1
                            trades2_log.append({
                                '買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx],
                                '股票名稱': name, 'T+1日買進價格': info['entry_price'], '股數': shares,
                                '賣出訊號日期': date, 'T+1日賣出價格': sell_price, '損益': pnl,
                                '報酬率': ret_pct, '買進原因': info['entry_reason'], '賣出原因': sell_reason
                            })

                            trades_log.append({
                                '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出',
                                '價格': sell_price, '股數': shares, '動能值': f"{roc[i][a_idx]*100:.2f}%",
                                '標的名稱': name, '原因': '再平衡',
                                '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax,
                                '說明': f"再平衡賣出：{name}"
                            })
                            surplus_pool += proceeds
                            slots[s_id] = None

                # 買入新訊號
                new_signals = [sig for sig in top_signals if sig not in signal_to_slot_map]
                available_slot_ids = [sid for sid, data in slots.items() if data is None]

                for sig in new_signals:
                    if not available_slot_ids: break
                    target_sid = available_slot_ids.pop(0)

                    invest_budget = self.slot_budget # 固定 400 萬
                    buy_price_exec = next_prices[sig]
                    cost_per_share = buy_price_exec * 1.001425
                    max_shares = int(invest_budget // cost_per_share)

                    remark = ""
                    if max_shares >= 1000:
                        shares = (max_shares // 1000) * 1000
                    else:
                        shares = max_shares
                        remark = "無法滿足1000股"

                    actual_cost = shares * buy_price_exec * 1.001425
                    buy_fee = shares * buy_price_exec * 0.001425
                    surplus_pool -= actual_cost

                    name = self.code_to_name[self.assets[sig]]
                    entry_reason = f"符合趨勢與濾網，ROC:{roc[i][sig]*100:.2f}%"

                    slots[target_sid] = {
                        'asset_idx': sig, 'shares': shares, 'max_price': buy_price_exec,
                        'budget': actual_cost, 'entry_date': date, 'entry_price': buy_price_exec,
                        'entry_reason': entry_reason, 'remark': remark
                    }

                    trades_log.append({
                        '訊號日期': date, '股票代號': self.assets[sig], '狀態': '買進',
                        '價格': buy_price_exec, '股數': shares, '動能值': f"{roc[i][sig]*100:.2f}%",
                        '標的名稱': name, '原因': '符合趨勢',
                        '買入手續費': buy_fee, '賣出手續費': 0, '賣出交易稅': 0,
                        '說明': f"買進新持有商品：{name}" + (f" ({remark})" if remark else "")
                    })

                # 保持紀錄
                for sig, sid in signal_to_slot_map.items():
                    name = self.code_to_name[self.assets[sig]]
                    trades_log.append({
                        '訊號日期': date, '股票代號': self.assets[sig], '狀態': '保持',
                        '價格': current_prices[sig], '股數': slots[sid]['shares'],
                        '動能值': f"{roc[i][sig]*100:.2f}%", '標的名稱': name, '原因': '趨勢持續',
                        '買入手續費': 0, '賣出手續費': 0, '賣出交易稅': 0, '說明': f"保留與上一期相同：{name}"
                    })

            # D. 持股備註與快照
            count = sum(1 for s in slots.values() if s and 'asset_idx' in s)
            holdings_history.append({
                'Date': date,
                'Holdings': ", ".join(h_names),
                'Count': count,
                '現金': surplus_pool,
                '股票市值': stock_mv,
                '總資產': total_equity,
                '補充說明': "；".join(current_remarks) if current_remarks else ""
            })

        return pd.DataFrame(equity_curve_data), pd.DataFrame(trades_log), pd.DataFrame(holdings_history), pd.DataFrame(trades2_log), pd.DataFrame(daily_details)

def calculate_metrics_adj(equity_curve_df, authorized_capital=60000000):
    if equity_curve_df.empty: return 0, 0, 0, 0, 0
    equity = equity_curve_df['權益']
    total_return = (equity.iloc[-1] / equity.iloc[0]) - 1
    days = (equity_curve_df['日期'].iloc[-1] - equity_curve_df['日期'].iloc[0]).days
    years = days / 365.25
    cagr = (1 + total_return) ** (1 / years) - 1 if years > 0 else 0
    max_dd = equity_curve_df['回撤(Drawdown)'].min()
    fixed_max_dd = equity_curve_df['固定基準 MDD'].min()
    calmar = cagr / abs(max_dd) if max_dd != 0 else 0
    return cagr, max_dd, fixed_max_dd, calmar, total_return
