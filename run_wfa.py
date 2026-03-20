import pandas as pd
import numpy as np
import xlsxwriter

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
    回測引擎：動態分配版 V1 (適用於指定區間)。
    """
    def __init__(self, prices, code_to_name, initial_capital=30000000):
        self.prices_df = prices
        self.prices = prices.values
        self.dates = prices.index
        self.assets = prices.columns
        self.code_to_name = code_to_name
        self.initial_capital = initial_capital

    def run(self, sma_period, roc_period, stop_loss_pct, rebalance_interval, start_date, end_date):
        # 1. 指標預計算 (全域計算以保證穩定性)
        sma = self.prices_df.rolling(window=sma_period).mean().values
        roc = self.prices_df.pct_change(periods=roc_period).values

        # 2. 確定起止索引
        mask = (self.dates >= pd.to_datetime(start_date)) & (self.dates <= pd.to_datetime(end_date))
        all_indices = np.where(mask)[0]
        if len(all_indices) == 0:
            return None, None

        first_idx = all_indices[0]
        last_idx = all_indices[-1]

        # 緩衝期檢查
        start_buffer = max(sma_period, roc_period)
        loop_start = max(first_idx, start_buffer)

        # 3. 帳戶與槽位初始化
        surplus_pool = float(self.initial_capital)
        slots = {0: None, 1: None, 2: None}

        equity_curve_data = []
        trade_count = 0
        peak_equity = float(self.initial_capital)

        for i in range(loop_start, last_idx + 1):
            date = self.dates[i]
            current_prices = self.prices[i]

            # A. 計算今日權益
            stock_mv = 0.0
            for s_id, info in slots.items():
                if info and 'asset_idx' in info:
                    a_idx = info['asset_idx']
                    mv = info['shares'] * current_prices[a_idx]
                    stock_mv += mv

            total_equity = surplus_pool + stock_mv
            if total_equity > peak_equity: peak_equity = total_equity
            drawdown = (total_equity - peak_equity) / peak_equity

            equity_curve_data.append({
                '日期': date,
                '權益': total_equity,
                '回撤': drawdown
            })

            if i == last_idx:
                break

            next_prices = self.prices[i+1] # T+1 執行價格

            # B. 每日檢查停損
            triggered_slots = []
            for s_id, info in slots.items():
                if info and 'asset_idx' in info:
                    a_idx = info['asset_idx']
                    curr_p = current_prices[a_idx]
                    if curr_p > info['max_price']:
                        info['max_price'] = curr_p
                    if curr_p < info['max_price'] * (1 - stop_loss_pct):
                        triggered_slots.append(s_id)

            for s_id in triggered_slots:
                info = slots[s_id]
                a_idx = info['asset_idx']
                sell_price = next_prices[a_idx]
                shares = info['shares']

                sell_fee = shares * sell_price * 0.001425
                sell_tax = shares * sell_price * 0.003
                proceeds = shares * sell_price - sell_fee - sell_tax
                surplus_pool += proceeds
                trade_count += 1
                slots[s_id] = None

            # C. 再平衡邏輯
            is_rebalance_day = (i - loop_start) % rebalance_interval == 0
            if is_rebalance_day:
                # 篩選訊號
                top_3_signals = []
                sorted_all = np.argsort(roc[i])[::-1]
                for idx in sorted_all:
                    if len(top_3_signals) >= 3: break
                    p, s, r = current_prices[idx], sma[i][idx], roc[i][idx]
                    if p > s and r > 0:
                        top_3_signals.append(idx)

                # 處理槽位與賣出
                signal_to_slot_map = {}
                for s_id, info in slots.items():
                    if info and 'asset_idx' in info:
                        if info['asset_idx'] in top_3_signals:
                            signal_to_slot_map[info['asset_idx']] = s_id # 續抱
                        else:
                            a_idx = info['asset_idx']
                            sell_price = next_prices[a_idx]
                            shares = info['shares']
                            sell_fee = shares * sell_price * 0.001425
                            sell_tax = shares * sell_price * 0.003
                            proceeds = shares * sell_price - sell_fee - sell_tax
                            surplus_pool += proceeds
                            trade_count += 1
                            slots[s_id] = {'pending_budget': proceeds}
                    else:
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
                        surplus_pool += (budget - actual_cost)
                        slots[target_sid] = {
                            'asset_idx': sig, 'shares': shares, 'max_price': buy_price_exec
                        }
                        trade_count += 1
                    else:
                        surplus_pool += budget
                        slots[target_sid] = None

                # 清理
                for sid in list(slots.keys()):
                    if slots[sid] and 'pending_budget' in slots[sid]:
                        surplus_pool += slots[sid]['pending_budget']
                        slots[sid] = None

        return pd.DataFrame(equity_curve_data), trade_count

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

def main():
    DATA_FILE = '個股合-1.xlsx'
    SMA_PERIOD = 87
    ROC_PERIOD = 54
    STOP_LOSS_PCT = 0.09
    REBALANCE = 6
    INITIAL_CAPITAL = 30000000

    periods = [
        ('2019-01-02', '2022-12-31'),
        ('2019-06-01', '2023-05-31'),
        ('2020-01-02', '2023-12-31'),
        ('2020-06-01', '2024-05-31'),
        ('2021-01-02', '2024-12-31'),
        ('2021-06-01', '2025-05-31'),
        ('2022-01-02', '2025-12-31'),
    ]

    prices, code_to_name = clean_data(DATA_FILE)
    bt = Backtester(prices, code_to_name, INITIAL_CAPITAL)

    summary_results = []
    all_equity_curves = []

    for start_str, end_str in periods:
        print(f"正在執行回測: {start_str} 至 {end_str}")
        eq_df, trades = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE, start_str, end_str)
        cagr, mdd, calmar = calculate_metrics(eq_df)

        period_label = f"{start_str} - {end_str}"
        summary_results.append({
            '回測期間': period_label,
            'CAGR': cagr,
            'MaxDD': mdd,
            'Calmar Ratio': calmar,
            '交易次數': trades
        })

        # 準備 Equity Curve 資料供合併
        # 我們將每個期間的日期與權益分開存放，或者放在一起
        temp_eq = eq_df[['日期', '權益']].copy()
        temp_eq.columns = [f'日期_{period_label}', f'權益_{period_label}']
        all_equity_curves.append(temp_eq)

    # 彙整 Summary
    summary_df = pd.DataFrame(summary_results)

    # 產出 Excel
    OUTPUT_FILE = 'walk-forward.xlsx'
    writer = pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter')

    # 1. Summary Sheet
    summary_df.to_excel(writer, sheet_name='Summary', index=False)
    workbook = writer.book
    summary_sheet = writer.sheets['Summary']

    # 格式化
    percent_fmt = workbook.add_format({'num_format': '0.00%'})
    num_fmt = workbook.add_format({'num_format': '0.00'})
    summary_sheet.set_column('B:C', 15, percent_fmt)
    summary_sheet.set_column('D:D', 15, num_fmt)
    summary_sheet.set_column('A:A', 30)

    # 2. Equity_Curves Sheet
    # 合併所有 Equity Curve
    # 由於各期間長度不同，我們用 concat
    curves_df = pd.concat(all_equity_curves, axis=1)
    curves_df.to_excel(writer, sheet_name='Equity_Curves', index=False)
    curves_sheet = writer.sheets['Equity_Curves']

    # 3. 插入圖表
    for idx, (start_str, end_str) in enumerate(periods):
        period_label = f"{start_str} - {end_str}"
        chart = workbook.add_chart({'type': 'line'})

        # 資料欄位位置 (Excel 索引從 0 開始，所以 A=0, B=1, C=2, D=3...)
        # 日期在 2*idx, 權益在 2*idx + 1
        date_col = 2 * idx
        val_col = 2 * idx + 1

        max_row = len(all_equity_curves[idx])

        chart.add_series({
            'name':       f'Equity {period_label}',
            'categories': ['Equity_Curves', 1, date_col, max_row, date_col],
            'values':     ['Equity_Curves', 1, val_col, max_row, val_col],
        })

        chart.set_title({'name': f'Equity Curve ({period_label})'})
        chart.set_x_axis({'name': 'Date'})
        chart.set_y_axis({'name': 'Equity'})
        chart.set_legend({'position': 'none'})

        # 放置圖表
        row_pos = idx * 15
        curves_sheet.insert_chart(row_pos, 2 * len(periods) + 2, chart)

    writer.close()
    print(f"已成功產出匯總報表: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
