import pandas as pd
import numpy as np
import pickle

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
    回測引擎：動態分配版。
    """
    def __init__(self, prices, code_to_name, initial_capital=30000000):
        self.prices_df = prices
        self.prices = prices.values
        self.dates = prices.index
        self.assets = prices.columns
        self.code_to_name = code_to_name
        self.initial_capital = initial_capital

    def run(self, sma_period, roc_period, stop_loss_pct, rebalance_interval=6):
        # 1. 指標預計算
        sma = self.prices_df.rolling(window=sma_period).mean().values
        roc = self.prices_df.pct_change(periods=roc_period).values

        # 2. 帳戶與槽位初始化
        # 中央資金池：存放初始資金、停損回流、再平衡超額預算及 rounding 餘額
        surplus_pool = float(self.initial_capital)
        # 3 個部位槽位，每個槽位包含: {asset_idx, shares, max_price, budget}
        slots = {0: None, 1: None, 2: None}

        equity_curve = np.zeros(len(self.dates))
        trades_log = []
        holdings_history = []

        # 用於紀錄狀態的持久變數
        current_reasons = []

        start_idx = max(sma_period, roc_period)

        for i in range(start_idx, len(self.dates)):
            date = self.dates[i]
            current_prices = self.prices[i]

            # 每日權益計算
            total_equity = surplus_pool
            for s_id, info in slots.items():
                if info and 'asset_idx' in info:
                    total_equity += info['shares'] * current_prices[info['asset_idx']]
            equity_curve[i] = total_equity

            if i == len(self.dates) - 1:
                break

            next_prices = self.prices[i+1] # T+1 執行價格

            # 3. 追蹤停損檢查 (每日)
            triggered_slots = []
            for s_id, info in slots.items():
                if info and 'asset_idx' in info:
                    a_idx = info['asset_idx']
                    curr_p = current_prices[a_idx]
                    if curr_p > info['max_price']:
                        info['max_price'] = curr_p
                    if curr_p < info['max_price'] * (1 - stop_loss_pct):
                        triggered_slots.append(s_id)

            # 執行停損賣出
            for s_id in triggered_slots:
                info = slots[s_id]
                a_idx = info['asset_idx']
                sell_price = next_prices[a_idx]
                shares = info['shares']

                # 計算成本與所得
                sell_fee = shares * sell_price * 0.001425
                sell_tax = shares * sell_price * 0.003
                proceeds = shares * sell_price - sell_fee - sell_tax

                # 停損資金全數回到中央池，等下次再平衡
                surplus_pool += proceeds

                name = self.code_to_name[self.assets[a_idx]]
                current_reasons.append(f"{name}因達停損標準需剔除")

                trades_log.append({
                    '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出',
                    '價格': sell_price, '股數': shares, '動能值': f"{roc[i][a_idx]*100:.2f}%",
                    '標的名稱': name, '原因': '停損',
                    '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax,
                    '說明': f"停損賣出：{name}"
                })
                slots[s_id] = None

            # 4. 再平衡邏輯 (每 rebalance_interval 天)
            is_rebalance_day = (i - start_idx) % rebalance_interval == 0
            if is_rebalance_day:
                current_reasons = [] # 重置理由

                # A. 篩選 Top 3 訊號
                eligible_mask = (current_prices > sma[i]) & (roc[i] > 0)
                top_3_signals = []
                exclusion_reasons = []
                if np.any(eligible_mask):
                    eligible_idxs = np.where(eligible_mask)[0]
                    eligible_rocs = roc[i][eligible_idxs]
                    # 全體按 ROC 排序以便記錄排除原因
                    sorted_all = np.argsort(roc[i])[::-1]
                    for idx in sorted_all:
                        if len(top_3_signals) >= 3: break
                        p = current_prices[idx]
                        s = sma[i][idx]
                        r = roc[i][idx]
                        if p > s and r > 0:
                            top_3_signals.append(idx)
                        else:
                            name = self.code_to_name[self.assets[idx]]
                            if r <= 0: exclusion_reasons.append(f"{name} ROC < 0")
                            elif p <= s: exclusion_reasons.append(f"{name} 價格 < SMA")

                # B. 第一步：處理槽位與預算
                signal_to_slot_map = {}
                for s_id, info in slots.items():
                    if info and 'asset_idx' in info:
                        if info['asset_idx'] in top_3_signals:
                            signal_to_slot_map[info['asset_idx']] = s_id # 續抱
                        else:
                            # 賣出：所得直接作為該槽位預算
                            a_idx = info['asset_idx']
                            sell_price = next_prices[a_idx]
                            shares = info['shares']
                            sell_fee = shares * sell_price * 0.001425
                            sell_tax = shares * sell_price * 0.003
                            proceeds = shares * sell_price - sell_fee - sell_tax
                            slots[s_id] = {'pending_budget': proceeds}

                            name = self.code_to_name[self.assets[a_idx]]
                            trades_log.append({
                                '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出',
                                '價格': sell_price, '股數': shares, '動能值': f"{roc[i][a_idx]*100:.2f}%",
                                '標的名稱': name, '原因': '再平衡',
                                '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax,
                                '說明': f"再平衡賣出：{name}"
                            })
                    else:
                        # 槽位為空，嘗試從池中分配 1000 萬
                        alloc = min(surplus_pool, 10000000.0)
                        surplus_pool -= alloc
                        slots[s_id] = {'pending_budget': alloc}

                # C. 第二步：買入新訊號
                new_signals = [sig for sig in top_3_signals if sig not in signal_to_slot_map]
                available_slot_ids = [sid for sid, data in slots.items() if data and 'pending_budget' in data]

                for sig in new_signals:
                    if not available_slot_ids: break
                    target_sid = available_slot_ids.pop(0)
                    budget = slots[target_sid]['pending_budget']

                    # 動態預算規則：使用賣出所得，上限 1000 萬
                    invest_budget = min(budget, 10000000.0)
                    buy_price_exec = next_prices[sig]
                    cost_per_share = buy_price_exec * 1.001425
                    shares = (int(invest_budget // cost_per_share) // 1000) * 1000

                    if shares > 0:
                        actual_cost = shares * buy_price_exec * 1.001425
                        buy_fee = shares * buy_price_exec * 0.001425
                        # 餘額回流中央池
                        surplus_pool += (budget - actual_cost)

                        slots[target_sid] = {
                            'asset_idx': sig, 'shares': shares,
                            'max_price': buy_price_exec, 'budget': actual_cost
                        }

                        name = self.code_to_name[self.assets[sig]]
                        trades_log.append({
                            '訊號日期': date, '股票代號': self.assets[sig], '狀態': '買進',
                            '價格': buy_price_exec, '股數': shares, '動能值': f"{roc[i][sig]*100:.2f}%",
                            '標的名稱': name, '原因': '符合趨勢',
                            '買入手續費': buy_fee, '賣出手續費': 0, '賣出交易稅': 0,
                            '說明': f"買進新持有商品：{name}"
                        })
                    else:
                        surplus_pool += budget
                        slots[target_sid] = None
                        name = self.code_to_name[self.assets[sig]]
                        current_reasons.append(f"因資金配置限制，無法配置第 {target_sid+1} 檔 ({name})")

                # D. 清理未使用的槽位預算
                for sid in list(slots.keys()):
                    if slots[sid] and 'pending_budget' in slots[sid]:
                        surplus_pool += slots[sid]['pending_budget']
                        slots[sid] = None

                # E. 紀錄保持
                for sig, sid in signal_to_slot_map.items():
                    name = self.code_to_name[self.assets[sig]]
                    trades_log.append({
                        '訊號日期': date, '股票代號': self.assets[sig], '狀態': '保持',
                        '價格': current_prices[sig], '股數': slots[sid]['shares'],
                        '動能值': f"{roc[i][sig]*100:.2f}%",
                        '標的名稱': name, '原因': '趨勢持續',
                        '買入手續費': 0, '賣出手續費': 0, '賣出交易稅': 0,
                        '說明': f"保留與上一期相同：{name}"
                    })

                # F. 組合再平衡摘要
                if len(top_3_signals) < 3:
                    rebal_msg = f"當次再平衡僅有 {len(top_3_signals)} 檔符合標準"
                    if exclusion_reasons:
                        rebal_msg += f"；排除原因: {'、'.join(exclusion_reasons[:2])}"
                    current_reasons.append(rebal_msg)

            # 5. 備註整理與快照
            count = sum(1 for s in slots.values() if s and 'asset_idx' in s)
            final_remark = "；".join(list(dict.fromkeys(current_reasons))) if count < 3 else ""

            h_names = []
            for sid in range(3):
                if slots[sid] and 'asset_idx' in slots[sid]:
                    h_names.append(f"{self.code_to_name[self.assets[slots[sid]['asset_idx']]]}({self.assets[slots[sid]['asset_idx']]})")

            holdings_history.append({
                'Date': date, 'Holdings': ", ".join(h_names), 'Count': count,
                'Equity': total_equity, '補充說明': final_remark
            })

        eq_series = pd.Series(equity_curve, index=self.dates)
        eq_series.iloc[:start_idx] = self.initial_capital
        eq_series = eq_series.replace(0, np.nan).ffill()
        return eq_series, pd.DataFrame(trades_log), pd.DataFrame(holdings_history)

def calculate_metrics(equity_curve):
    if equity_curve.empty: return 0, 0, 0, 0
    total_return = (equity_curve.iloc[-1] / equity_curve.iloc[0]) - 1
    days = (equity_curve.index[-1] - equity_curve.index[0]).days
    years = days / 365.25
    cagr = (1 + total_return) ** (1 / years) - 1 if years > 0 else 0
    rolling_max = equity_curve.cummax()
    drawdowns = (equity_curve - rolling_max) / rolling_max
    max_dd = drawdowns.min()
    calmar = cagr / abs(max_dd) if max_dd != 0 else 0
    return cagr, max_dd, calmar, total_return

if __name__ == "__main__":
    prices, code_to_name = clean_data('個股合-1.xlsx')
    bt = Backtester(prices, code_to_name)
    eq, trades, hold = bt.run(87, 54, 0.09, 6)
    cagr, mdd, calmar, ret = calculate_metrics(eq)
    print(f"Equity2025新-動態版 回測完成: CAGR={cagr:.2%}, MDD={mdd:.2%}, Calmar={calmar:.2f}")
