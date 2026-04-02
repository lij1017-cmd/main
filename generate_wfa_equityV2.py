import pandas as pd
import numpy as np
import xlsxwriter
import datetime

# =============================================================================
# 函式：資料清洗與預處理 (與使用者資料格式同步)
# =============================================================================
def clean_data(filepath):
    """
    讀取並清洗 Excel 資料，包含價格與成交量分頁。
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

    # 資料清洗：價格向後/向前填充，成交量補 0
    prices = prices.ffill().bfill()
    volumes = volumes.fillna(0)

    return prices, volumes, code_to_name

# =============================================================================
# 類別：資產類別趨勢追蹤回測引擎 (動態分配版 V1)
# =============================================================================
class Backtester:
    """
    回測引擎，支援動態資金分配 (Dynamic V1)，確保每個槽位投資上限為 1000 萬。
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

    def run(self, sma_period, roc_period, stop_loss_pct, rebalance_interval, start_date, end_date):
        """
        執行單一區間的回測模擬。
        """
        # 1. 技術指標預計算
        sma = self.prices_df.rolling(window=sma_period).mean().values
        roc = self.prices_df.pct_change(periods=roc_period).values
        sma5 = self.prices_df.rolling(window=5).mean().values
        sma10 = self.prices_df.rolling(window=10).mean().values
        sma20 = self.prices_df.rolling(window=20).mean().values

        # 確定本次回測區間在資料集中的索引位置
        mask = (self.dates >= pd.to_datetime(start_date)) & (self.dates <= pd.to_datetime(end_date))
        all_indices = np.where(mask)[0]
        if len(all_indices) == 0:
            return None, 0

        first_idx = all_indices[0]
        last_idx = all_indices[-1]

        # 緩衝期：303日均線需至少303天數據。設定全域基準點確保再平衡日期不位移。
        global_start_buffer = max(sma_period, roc_period, 20)
        loop_start = max(first_idx, global_start_buffer)

        # 2. 帳戶狀態與槽位初始化 (Dynamic V1)
        surplus_pool = float(self.initial_capital)
        slots = {0: None, 1: None, 2: None}

        equity_curve_list = []
        trade_count = 0

        # 3. 每日模擬循環
        # 核心：從 loop_start 開始記錄權益並執行交易。
        for i in range(loop_start, last_idx + 1):
            date = self.dates[i]
            current_prices = self.prices[i]

            # A. 計算今日權益總額
            stock_mv = 0.0
            for s_id, info in slots.items():
                if info and 'asset_idx' in info:
                    stock_mv += info['shares'] * current_prices[info['asset_idx']]

            total_equity = surplus_pool + stock_mv
            equity_curve_list.append({'日期': date, '權益': total_equity})

            # 若未達可計算指標的緩衝期，或已達最後一日，不執行交易邏輯
            if i == last_idx:
                continue

            next_prices = self.prices[i+1] # T+1 執行價格

            # B. 檢查停損機制 (最高價回落)
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
                sell_fee = shares * sell_price * 0.001425
                sell_tax = shares * sell_price * 0.003
                proceeds = shares * sell_price - sell_fee - sell_tax

                surplus_pool += proceeds
                trade_count += 1
                slots[s_id] = None

            # C. 再平衡邏輯 (錨定 loop_start，解決各區間再平衡日期不一致之問題)
            is_rebalance_day = (i - loop_start) % rebalance_interval == 0
            if is_rebalance_day:
                # 篩選符合條件標的
                signals = []
                sorted_all = np.argsort(roc[i])[::-1]
                for idx in sorted_all:
                    if len(signals) >= 3: break
                    p, s, r = current_prices[idx], sma[i][idx], roc[i][idx]
                    v = self.volumes[i][idx]
                    amount = p * v * 1000
                    if (p > s and r > 0 and amount > 30000000 and
                        p > sma5[i][idx] and p > sma10[i][idx] and p > sma20[i][idx]):
                        signals.append(idx)

                # 處理現有持股：排名外則賣出並保留預算在槽位中
                signal_to_slot = {}
                for s_id, info in slots.items():
                    if info and 'asset_idx' in info:
                        if info['asset_idx'] in signals:
                            signal_to_slot[info['asset_idx']] = s_id
                        else:
                            a_idx = info['asset_idx']
                            sell_price = next_prices[a_idx]
                            shares = info['shares']
                            sell_fee = shares * sell_price * 0.001425
                            sell_tax = shares * sell_price * 0.003
                            proceeds = shares * sell_price - sell_fee - sell_tax
                            slots[s_id] = {'pending_budget': proceeds}
                            trade_count += 1
                    elif info and 'pending_budget' in info:
                        pass
                    else:
                        # 槽位空缺，撥入預算
                        alloc = min(surplus_pool, 10000000.0)
                        surplus_pool -= alloc
                        slots[s_id] = {'pending_budget': alloc}

                # 買入新標的 (上限 1000 萬)
                new_signals = [s for s in signals if s not in signal_to_slot]
                available_sids = [sid for sid, data in slots.items() if data and 'pending_budget' in data]

                for sig in new_signals:
                    if not available_sids: break
                    target_sid = available_sids.pop(0)
                    budget = slots[target_sid]['pending_budget']
                    invest_budget = min(budget, 10000000.0)

                    buy_price_exec = next_prices[sig]
                    cost_per_share = buy_price_exec * 1.001425
                    shares = (int(invest_budget // cost_per_share) // 1000) * 1000

                    if shares > 0:
                        actual_cost = shares * buy_price_exec * 1.001425
                        surplus_pool += (budget - actual_cost)
                        slots[target_sid] = {
                            'asset_idx': sig,
                            'shares': shares,
                            'max_price': buy_price_exec,
                            'budget': actual_cost,
                            'entry_date': date,
                            'entry_price': buy_price_exec
                        }
                        trade_count += 1
                    else:
                        surplus_pool += budget
                        slots[target_sid] = None

                # 清理本期未使用的槽位預算
                for sid in list(slots.keys()):
                    if slots[sid] and 'pending_budget' in slots[sid]:
                        surplus_pool += slots[sid]['pending_budget']
                        slots[sid] = None

        return pd.DataFrame(equity_curve_list), trade_count

# =============================================================================
# 績效指標計算 (同步使用者之 WFA 基準：僅計算主動交易期間)
# =============================================================================
def calculate_metrics(eq_df):
    if eq_df is None or eq_df.empty: return 0, 0, 0
    equity = eq_df['權益']
    # 將分母設為實際有數據的天數 (即排除指標緩衝期後的交易日)
    total_return = (equity.iloc[-1] / equity.iloc[0]) - 1
    days = (eq_df['日期'].iloc[-1] - eq_df['日期'].iloc[0]).days
    years = days / 365.25
    cagr = (1 + total_return) ** (1 / years) - 1 if years > 0 else 0
    rolling_max = equity.cummax()
    drawdowns = (equity - rolling_max) / rolling_max
    max_dd = drawdowns.min()
    calmar = cagr / abs(max_dd) if max_dd != 0 else 0
    return cagr, max_dd, calmar

# =============================================================================
# 執行與 Excel 生成 (V2 版本)
# =============================================================================
if __name__ == "__main__":
    DATA_FILE = '資料-1.xlsx'
    SMA_PERIOD = 303
    ROC_PERIOD = 14
    STOP_LOSS_PCT = 0.0999
    REBALANCE = 9
    INITIAL_CAPITAL = 30000000

    # 九個調整後的 WFA 區間 (V2)
    periods = [
        ('2019-01-02', '2021-12-31'),
        ('2019-06-01', '2022-05-31'),
        ('2020-01-02', '2022-12-31'),
        ('2020-06-01', '2023-05-31'),
        ('2021-01-02', '2023-12-31'),
        ('2021-06-01', '2024-05-31'),
        ('2022-01-02', '2024-12-31'),
        ('2022-06-01', '2025-05-31'),
        ('2023-01-02', '2025-12-31'),
    ]

    prices, volumes, code_to_name = clean_data(DATA_FILE)
    bt = Backtester(prices, volumes, code_to_name, INITIAL_CAPITAL)

    summary_results = []
    all_equity_curves = []

    for idx, (start_str, end_str) in enumerate(periods):
        print(f"正在執行 WFA 區間 {idx+1}: {start_str} 至 {end_str}")
        eq_df, trades = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE, start_str, end_str)
        cagr, mdd, calmar = calculate_metrics(eq_df)

        period_label = f"{idx+1}. {start_str.replace('-', '.')} - {end_str.replace('-', '.')}"
        summary_results.append({
            '回測區間': period_label,
            '年化報酬率 (CAGR)': cagr,
            '最大回撤 (MaxDD)': mdd,
            'Calmar Ratio': calmar,
            '交易次數': trades
        })

        temp_eq = eq_df.copy()
        temp_eq.columns = ['日期', f'區間{idx+1}_權益']
        all_equity_curves.append(temp_eq)

    OUTPUT_FILE = 'walk-forward-equityV2.xlsx'
    writer = pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter')

    summary_df = pd.DataFrame(summary_results)
    summary_df.to_excel(writer, sheet_name='Summary', index=False)
    workbook = writer.book
    summary_sheet = writer.sheets['Summary']

    percent_fmt = workbook.add_format({'num_format': '0.00%', 'align': 'center'})
    num_fmt = workbook.add_format({'num_format': '0.00', 'align': 'center'})
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1, 'align': 'center'})

    for col_num, value in enumerate(summary_df.columns.values):
        summary_sheet.write(0, col_num, value, header_fmt)

    summary_sheet.set_column('A:A', 35)
    summary_sheet.set_column('B:C', 18, percent_fmt)
    summary_sheet.set_column('D:D', 15, num_fmt)
    summary_sheet.set_column('E:E', 12, num_fmt)

    final_curves_df = all_equity_curves[0]
    for next_df in all_equity_curves[1:]:
        final_curves_df = pd.merge(final_curves_df, next_df, on='日期', how='outer')

    final_curves_df = final_curves_df.sort_values('日期')
    final_curves_df.to_excel(writer, sheet_name='Equity_Curve', index=False)
    curves_sheet = writer.sheets['Equity_Curve']

    date_fmt = workbook.add_format({'num_format': 'yyyy/mm/dd', 'align': 'center'})
    curves_sheet.set_column('A:A', 12, date_fmt)

    for idx in range(len(periods)):
        chart = workbook.add_chart({'type': 'line'})
        period_label = f"Interval {idx+1}"
        col_idx = idx + 1
        max_row = len(final_curves_df)
        chart.add_series({
            'name':       ['Equity_Curve', 0, col_idx],
            'categories': ['Equity_Curve', 1, 0, max_row, 0],
            'values':     ['Equity_Curve', 1, col_idx, max_row, col_idx],
            'line':       {'width': 1.5},
        })
        chart.set_title({'name': f'Equity Curve - {period_label}'})
        chart.set_legend({'position': 'none'})
        chart.set_size({'width': 600, 'height': 350})
        curves_sheet.insert_chart(idx * 18, len(periods) + 2, chart)

    writer.close()
    print(f"回測完成！結果已儲存至: {OUTPUT_FILE}")
