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
    def __init__(self, prices, code_to_name, initial_capital=30000000):
        self.prices_df = prices
        self.prices = prices.values
        self.dates = prices.index
        self.assets = prices.columns
        self.code_to_name = code_to_name
        self.initial_capital = initial_capital

    def run(self, sma_period, roc_period, stop_loss_pct, rebalance_interval, start_date, end_date):
        # 1. 指標預計算
        sma = self.prices_df.rolling(window=sma_period).mean().values
        roc = self.prices_df.pct_change(periods=roc_period).values

        # 2. 確定區間索引
        mask = (self.dates >= pd.to_datetime(start_date)) & (self.dates <= pd.to_datetime(end_date))
        all_indices = np.where(mask)[0]
        if len(all_indices) == 0:
            return None, 0

        first_idx = all_indices[0]
        last_idx = all_indices[-1]

        start_buffer = max(sma_period, roc_period)
        loop_start = max(first_idx, start_buffer)

        # 3. 帳戶初始化
        cash = float(self.initial_capital)
        portfolio = {}
        equity_curve_list = []
        trade_count = 0

        for i in range(first_idx, loop_start):
            equity_curve_list.append({
                '日期': self.dates[i],
                '權益': float(self.initial_capital)
            })

        for i in range(loop_start, last_idx + 1):
            date = self.dates[i]
            current_prices = self.prices[i]

            total_equity = cash
            for a_idx, info in portfolio.items():
                total_equity += info['shares'] * current_prices[a_idx]

            equity_curve_list.append({
                '日期': date,
                '權益': total_equity
            })

            if i == last_idx:
                break

            next_prices = self.prices[i+1]

            # A. 停損
            triggered_sl_idxs = []
            for a_idx, info in portfolio.items():
                curr_p = current_prices[a_idx]
                if curr_p > info['max_price']:
                    info['max_price'] = curr_p
                if curr_p < info['max_price'] * (1 - stop_loss_pct):
                    triggered_sl_idxs.append(a_idx)

            # B. 再平衡 (以區間起始點為錨點)
            is_rebalance_day = (i - loop_start) % rebalance_interval == 0
            top_3_signals = []
            if is_rebalance_day:
                sorted_indices = np.argsort(roc[i])[::-1]
                for idx in sorted_indices:
                    if len(top_3_signals) >= 3: break
                    if current_prices[idx] > sma[i][idx] and roc[i][idx] > 0:
                        top_3_signals.append(idx)

            # C. 賣出
            assets_to_sell = set(triggered_sl_idxs)
            if is_rebalance_day:
                for a_idx in portfolio.keys():
                    if a_idx not in top_3_signals:
                        assets_to_sell.add(a_idx)

            for a_idx in list(assets_to_sell):
                if a_idx in portfolio:
                    info = portfolio.pop(a_idx)
                    sell_price = next_prices[a_idx]
                    shares = info['shares']
                    sell_fee = shares * sell_price * 0.001425
                    sell_tax = shares * sell_price * 0.003
                    cash += (shares * sell_price - sell_fee - sell_tax)
                    trade_count += 1

            # D. 買進
            if is_rebalance_day:
                slot_cap = self.initial_capital / 3
                for a_idx in top_3_signals:
                    if a_idx in portfolio or len(portfolio) >= 3: continue
                    buy_price_exec = next_prices[a_idx]
                    shares = (int(slot_cap // (buy_price_exec * 1.001425)) // 1000) * 1000
                    if shares > 0:
                        cost = shares * buy_price_exec * 1.001425
                        if cash >= cost:
                            cash -= cost
                            portfolio[a_idx] = {'shares': shares, 'max_price': buy_price_exec}
                            trade_count += 1

        return pd.DataFrame(equity_curve_list), trade_count

def calculate_metrics(eq_df):
    if eq_df is None or eq_df.empty: return 0, 0, 0
    equity = eq_df['權益']
    total_return = (equity.iloc[-1] / equity.iloc[0]) - 1
    days = (eq_df['日期'].iloc[-1] - eq_df['日期'].iloc[0]).days
    years = days / 365.25
    cagr = (1 + total_return) ** (1 / years) - 1 if years > 0 else 0
    rolling_max = equity.cummax()
    drawdowns = (equity - rolling_max) / rolling_max
    max_dd = drawdowns.min()
    calmar = cagr / abs(max_dd) if max_dd != 0 else 0
    return cagr, max_dd, calmar

def main():
    DATA_FILE = '個股合-1.xlsx'
    SMA_PERIOD = 87
    ROC_PERIOD = 54
    STOP_LOSS_PCT = 0.09
    REBALANCE = 6
    INITIAL_CAPITAL = 30000000

    # 最新要求的 WFA 區間
    periods = [
        ('2024-06-01', '2025-12-31'),
        ('2024-01-02', '2025-05-31'),
        ('2023-01-02', '2024-12-31'),
        ('2022-01-02', '2024-05-31'),
        ('2021-06-01', '2023-12-31'),
        ('2021-01-02', '2023-05-31'),
        ('2020-01-02', '2022-12-31'),
        ('2019-06-01', '2022-05-31'),
        ('2019-01-02', '2021-12-31'),
    ]

    prices, code_to_name = clean_data(DATA_FILE)
    bt = Backtester(prices, code_to_name, INITIAL_CAPITAL)

    summary_results = []
    all_equity_curves = []

    for start_str, end_str in periods:
        print(f"Executing WFA: {start_str} to {end_str}")
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

        temp_eq = eq_df.copy()
        temp_eq.columns = [f'日期_{period_label}', f'權益_{period_label}']
        all_equity_curves.append(temp_eq)

    summary_df = pd.DataFrame(summary_results)
    OUTPUT_FILE = 'walk-forward-3.xlsx'
    writer = pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter')
    summary_df.to_excel(writer, sheet_name='Summary', index=False)
    workbook = writer.book
    summary_sheet = writer.sheets['Summary']
    percent_fmt = workbook.add_format({'num_format': '0.00%'})
    num_fmt = workbook.add_format({'num_format': '0.00'})
    summary_sheet.set_column('B:C', 15, percent_fmt)
    summary_sheet.set_column('D:D', 15, num_fmt)
    summary_sheet.set_column('A:A', 30)

    curves_df = pd.concat(all_equity_curves, axis=1)
    curves_df.to_excel(writer, sheet_name='Equity_Curves', index=False)
    curves_sheet = writer.sheets['Equity_Curves']

    for idx, (start_str, end_str) in enumerate(periods):
        period_label = f"{start_str} - {end_str}"
        chart = workbook.add_chart({'type': 'line'})
        date_col = 2 * idx
        val_col = 2 * idx + 1
        max_row = len(all_equity_curves[idx])
        chart.add_series({
            'name':       f'Equity {period_label}',
            'categories': ['Equity_Curves', 1, date_col, max_row, date_col],
            'values':     ['Equity_Curves', 1, val_col, max_row, val_col],
        })
        chart.set_title({'name': f'Equity Curve ({period_label})'})
        chart.set_legend({'position': 'none'})
        row_pos = idx * 15
        curves_sheet.insert_chart(row_pos, 2 * len(periods) + 2, chart)

    writer.close()
    print(f"WFA Done: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
