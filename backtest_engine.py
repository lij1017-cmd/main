import pandas as pd
import numpy as np

def clean_data(filepath):
    """
    清洗並預處理輸入的 Excel 資料檔。
    """
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
    """
    回測引擎：動態分配版 V1 (Dynamic Allocation V1)。
    支援槽位管理、賣出資金立即再投入、1000股整張單位交易。
    """
    def __init__(self, prices, code_to_name, initial_capital=30000000):
        self.prices_df = prices
        self.prices = prices.values
        self.dates = prices.index
        self.assets = prices.columns
        self.code_to_name = code_to_name
        self.initial_capital = initial_capital

    def run(self, sma_period, roc_period, stop_loss_pct, rebalance_interval=9, start_date=None, end_date=None):
        # 1. 指標預計算
        sma = self.prices_df.rolling(window=sma_period).mean().values
        roc = self.prices_df.pct_change(periods=roc_period).values

        # 確定回測區間
        if start_date is None: start_date = self.dates[0]
        if end_date is None: end_date = self.dates[-1]

        mask = (self.dates >= pd.to_datetime(start_date)) & (self.dates <= pd.to_datetime(end_date))
        all_indices = np.where(mask)[0]
        if len(all_indices) == 0: return None, None, None, None, None

        first_idx = all_indices[0]
        last_idx = all_indices[-1]
        start_buffer = max(sma_period, roc_period)
        loop_start = max(first_idx, start_buffer)

        # 2. 帳戶與槽位初始化
        surplus_pool = float(self.initial_capital)
        # slots: {id: {asset_idx, shares, max_price, budget, entry_date, entry_price, entry_reason}}
        slots = {0: None, 1: None, 2: None}

        equity_curve_data = []
        trades_log = []
        trades2_log = []
        holdings_history = []
        daily_details = []

        peak_equity = float(self.initial_capital)
        current_reasons = []

        # 初始權益紀錄 (若 loop_start > first_idx)
        for i in range(first_idx, loop_start):
            equity_curve_data.append({'日期': self.dates[i], '權益': float(self.initial_capital), '回撤': 0.0})

        for i in range(loop_start, last_idx + 1):
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
                    h_names.append(f"{self.code_to_name[self.assets[a_idx]]}({self.assets[a_idx]})")
                    daily_details.append({
                        '日期': date, '股票代號': self.assets[a_idx], '股票名稱': self.code_to_name[self.assets[a_idx]],
                        '持有股數': info['shares'], '本日收盤價': current_prices[a_idx], '市值': mv
                    })

            total_equity = surplus_pool + stock_mv
            if total_equity > peak_equity: peak_equity = total_equity
            drawdown = (total_equity - peak_equity) / peak_equity if peak_equity > 0 else 0
            equity_curve_data.append({'日期': date, '權益': total_equity, '回撤': drawdown})

            if i == last_idx: break
            next_prices = self.prices[i+1]

            # B. 每日檢查停損
            triggered_slots = []
            for s_id, info in slots.items():
                if info and 'asset_idx' in info:
                    a_idx = info['asset_idx']
                    curr_p = current_prices[a_idx]
                    if curr_p > info['max_price']: info['max_price'] = curr_p
                    if curr_p < info['max_price'] * (1 - stop_loss_pct): triggered_slots.append(s_id)

            # 執行停損賣出
            for s_id in triggered_slots:
                info = slots[s_id]
                a_idx = info['asset_idx']
                sell_price = next_prices[a_idx]
                shares = info['shares']
                sell_fee = shares * sell_price * 0.001425
                sell_tax = shares * sell_price * 0.003
                proceeds = shares * sell_price - sell_fee - sell_tax
                surplus_pool += proceeds

                name = self.code_to_name[self.assets[a_idx]]
                trades2_log.append({
                    '買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx], '股票名稱': name,
                    'T+1日買進價格': info['entry_price'], '股數': shares, '賣出訊號日期': date,
                    'T+1日賣出價格': sell_price, '損益': proceeds - info['budget'],
                    '報酬率': (proceeds / info['budget']) - 1, '買進原因': '符合趨勢', '賣出原因': '停損機制'
                })
                trades_log.append({
                    '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出', '價格': sell_price,
                    '股數': shares, '動能值': f"{roc[i][a_idx]*100:.2f}%", '標的名稱': name, '原因': '停損',
                    '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax
                })
                slots[s_id] = None

            # C. 再平衡邏輯
            is_rebalance_day = (i - loop_start) % rebalance_interval == 0
            if is_rebalance_day:
                current_reasons = []
                top_3_signals = []
                exclusion_reasons = []
                sorted_all = np.argsort(roc[i])[::-1]
                for idx in sorted_all:
                    if len(top_3_signals) >= 3: break
                    p, s, r = current_prices[idx], sma[i][idx], roc[i][idx]
                    if p > s and r > 0: top_3_signals.append(idx)
                    else:
                        name = self.code_to_name[self.assets[idx]]
                        if r <= 0: exclusion_reasons.append(f"{name} ROC < 0")
                        elif p <= s: exclusion_reasons.append(f"{name} 價格 < SMA")

                # 處理槽位與賣出 (不在 Top 3 的賣出)
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
                            trades2_log.append({
                                '買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx], '股票名稱': name,
                                'T+1日買進價格': info['entry_price'], '股數': shares, '賣出訊號日期': date,
                                'T+1日賣出價格': sell_price, '損益': proceeds - info['budget'],
                                '報酬率': (proceeds / info['budget']) - 1, '買進原因': '符合趨勢', '賣出原因': '再平衡排名外'
                            })
                            trades_log.append({
                                '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出', '價格': sell_price,
                                '股數': shares, '動能值': f"{roc[i][a_idx]*100:.2f}%", '標的名稱': name, '原因': '再平衡',
                                '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax
                            })
                            slots[s_id] = {'pending_budget': proceeds}
                    else:
                        # 已經是空位的，給予預設分配額 (1000萬)
                        alloc = min(surplus_pool, 10000000.0)
                        surplus_pool -= alloc
                        slots[s_id] = {'pending_budget': alloc}

                # 買入新訊號
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
                            'budget': actual_cost, 'entry_date': date, 'entry_price': buy_price_exec
                        }
                        trades_log.append({
                            '訊號日期': date, '股票代號': self.assets[sig], '狀態': '買進', '價格': buy_price_exec,
                            '股數': shares, '動能值': f"{roc[i][sig]*100:.2f}%", '標的名稱': name, '原因': '符合趨勢',
                            '買入手續費': buy_fee, '賣出手續費': 0, '賣出交易稅': 0
                        })
                    else:
                        surplus_pool += budget
                        slots[target_sid] = None
                        current_reasons.append(f"因資金配置限制，無法配置 {self.code_to_name[self.assets[sig]]}")

                # 清理臨時槽位資訊
                for sid in list(slots.keys()):
                    if slots[sid] and 'pending_budget' in slots[sid]:
                        surplus_pool += slots[sid]['pending_budget']
                        slots[sid] = None

                # 紀錄保持持股
                for sig, sid in signal_to_slot_map.items():
                    name = self.code_to_name[self.assets[sig]]
                    trades_log.append({
                        '訊號日期': date, '股票代號': self.assets[sig], '狀態': '保持', '價格': current_prices[sig],
                        '股數': slots[sid]['shares'], '動能值': f"{roc[i][sig]*100:.2f}%", '標的名稱': name,
                        '原因': '趨勢持續', '買入手續費': 0, '賣出手續費': 0, '賣出交易稅': 0
                    })

                if len(top_3_signals) < 3:
                    rebal_msg = f"當次再平衡僅有 {len(top_3_signals)} 檔符合標準"
                    if exclusion_reasons: rebal_msg += f"；排除原因: {'、'.join(exclusion_reasons[:2])}"
                    current_reasons.append(rebal_msg)

            # D. 持股備註與歷史紀錄
            count = sum(1 for s in slots.values() if s and 'asset_idx' in s)
            final_remark = "；".join(list(dict.fromkeys(current_reasons))) if count < 3 else ""
            holdings_history.append({
                'Date': date, 'Holdings': ", ".join(h_names), 'Count': count, '現金': surplus_pool, '股票市值': stock_mv, '總資產': total_equity, '補充說明': final_remark
            })

        return pd.DataFrame(equity_curve_data), pd.DataFrame(trades_log), pd.DataFrame(holdings_history), pd.DataFrame(trades2_log), pd.DataFrame(daily_details)

def calculate_metrics(eq_df):
    if eq_df is None or eq_df.empty: return 0, 0, 0
    equity = eq_df['權益']
    total_return = (equity.iloc[-1] / equity.iloc[0]) - 1
    days = (eq_df['日期'].iloc[-1] - eq_df['日期'].iloc[0]).days
    years = days / 365.25
    cagr = (1 + total_return) ** (1 / years) - 1 if years > 0 else 0
    max_dd = eq_df['回撤'].min()
    calmar = cagr / abs(max_dd) if max_dd != 0 else 0
    return cagr, max_dd, calmar
