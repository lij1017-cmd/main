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

def calculate_indicators(prices_df):
    # EMA 200
    ema200 = prices_df.ewm(span=200, adjust=False).mean()

    # MACD
    exp1 = prices_df.ewm(span=12, adjust=False).mean()
    exp2 = prices_df.ewm(span=26, adjust=False).mean()
    macd_line = exp1 - exp2
    signal_line = macd_line.ewm(span=9, adjust=False).mean()
    macd_hist = macd_line - signal_line

    # Rolling Volatility (Proxy for ATR) - using 20-day standard deviation
    volatility = prices_df.rolling(window=20).std()

    # Swing High/Low (Proxy for Pivot S/R) - using 10-day rolling max/min
    swing_high = prices_df.rolling(window=10).max()
    swing_low = prices_df.rolling(window=10).min()

    return ema200, macd_line, signal_line, macd_hist, volatility, swing_high, swing_low

class PhantomBacktester:
    def __init__(self, prices, volumes, code_to_name, initial_capital=30000000, mode='LS'):
        self.prices_df = prices
        self.volumes_df = volumes
        self.prices = prices.values
        self.dates = prices.index
        self.assets = prices.columns
        self.code_to_name = code_to_name
        self.initial_capital = initial_capital
        self.mode = mode # 'L' for Long-only, 'LS' for Long/Short

        # Strategy Parameters
        self.ema_len = 200
        self.macd_fast = 12
        self.macd_slow = 26
        self.macd_signal = 9
        self.swing_window = 10
        self.vol_frac = 0.5
        self.rr_target = 1.5

        # Transaction Costs
        self.commission_rate = 0.001425
        self.tax_rate = 0.003
        self.short_fee_rate = 0.001 # 融券手續費

    def run(self):
        # print(f"Calculating indicators (Mode: {self.mode})...")
        ema200, macd_line, signal_line, macd_hist, volatility, swing_high, swing_low = calculate_indicators(self.prices_df)

        ema200_v = ema200.values
        macd_line_v = macd_line.values
        signal_line_v = signal_line.values
        vol_v = volatility.values
        swing_high_v = swing_high.values
        swing_low_v = swing_low.values

        surplus_pool = float(self.initial_capital)
        slots = {i: None for i in range(10)}

        equity_curve = []
        trades_log = []

        break_low_active = np.zeros(len(self.assets), dtype=bool)
        break_high_active = np.zeros(len(self.assets), dtype=bool)

        start_idx = max(self.ema_len, self.macd_slow, 20)

        peak_equity = float(self.initial_capital)
        for i in range(start_idx, len(self.dates)):
            curr_prices = self.prices[i]

            # A. Equity Calculation
            long_mv = 0.0
            short_mv = 0.0
            for s_id, info in slots.items():
                if info:
                    mv = info['shares'] * curr_prices[info['asset_idx']]
                    if info['type'] == 'Long':
                        long_mv += mv
                    else:
                        short_mv += mv

            total_equity = surplus_pool + long_mv - short_mv
            if total_equity > peak_equity: peak_equity = total_equity
            drawdown = (total_equity - peak_equity) / peak_equity if peak_equity != 0 else 0

            equity_curve.append({'日期': self.dates[i], '權益': total_equity, '回撤(Drawdown)': drawdown})

            if i == len(self.dates) - 1: break

            next_prices = self.prices[i+1]

            # B. Check for Exit (TP/SL)
            for s_id, info in slots.items():
                if info:
                    a_idx = info['asset_idx']
                    cp = curr_prices[a_idx]
                    exit_triggered = False
                    exit_reason = ""

                    if info['type'] == 'Long':
                        if cp <= info['stop_loss']:
                            exit_triggered = True
                            exit_reason = "Stop Loss"
                        elif cp >= info['take_profit']:
                            exit_triggered = True
                            exit_reason = "Take Profit"
                    else: # Short
                        if cp >= info['stop_loss']:
                            exit_triggered = True
                            exit_reason = "Stop Loss"
                        elif cp <= info['take_profit']:
                            exit_triggered = True
                            exit_reason = "Take Profit"

                    if exit_triggered:
                        exit_price = next_prices[a_idx]
                        shares = info['shares']
                        if info['type'] == 'Long':
                            # Sell Long: receive Cash minus commission minus tax
                            proceeds = shares * exit_price * (1 - self.commission_rate - self.tax_rate)
                            surplus_pool += proceeds
                            pnl = proceeds - info['cost']
                        else:
                            # Buy to Cover Short: pay Cash plus commission
                            cost_to_cover = shares * exit_price * (1 + self.commission_rate)
                            surplus_pool -= cost_to_cover
                            pnl = info['cost'] - cost_to_cover

                        trades_log.append({
                            'Exit Date': self.dates[i],
                            'Asset': self.assets[a_idx],
                            'Type': info['type'],
                            'Entry Price': info['entry_price'],
                            'Exit Price': exit_price,
                            'Reason': exit_reason,
                            'PnL': pnl
                        })
                        slots[s_id] = None

            # C. Entry Signals
            if any(s is None for s in slots.values()):
                for a in range(len(self.assets)):
                    cp = curr_prices[a]
                    support = swing_low_v[i-1, a]
                    resistance = swing_high_v[i-1, a]
                    tol = vol_v[i, a] * self.vol_frac

                    if cp < support - tol: break_low_active[a] = True
                    if cp > resistance + tol: break_high_active[a] = True

                    if any(info and info['asset_idx'] == a for info in slots.values()):
                        continue

                    macd_cross_up = macd_line_v[i, a] > signal_line_v[i, a] and macd_line_v[i-1, a] <= signal_line_v[i-1, a]
                    macd_cross_down = macd_line_v[i, a] < signal_line_v[i, a] and macd_line_v[i-1, a] >= signal_line_v[i-1, a]

                    core_long = macd_cross_up and macd_line_v[i, a] < 0 and cp > ema200_v[i, a]
                    core_short = macd_cross_down and macd_line_v[i, a] > 0 and cp < ema200_v[i, a]

                    if core_long:
                        reclaim = break_low_active[a] and cp > support
                        if reclaim:
                            stop_loss = support - tol
                            if stop_loss < cp:
                                surplus_pool = self.enter_trade(slots, surplus_pool, a, i, 'Long', cp, next_prices[a], stop_loss, trades_log)
                                break_low_active[a] = False
                    elif core_short and self.mode == 'LS':
                        reclaim = break_high_active[a] and cp < resistance
                        if reclaim:
                            stop_loss = resistance + tol
                            if stop_loss > cp:
                                surplus_pool = self.enter_trade(slots, surplus_pool, a, i, 'Short', cp, next_prices[a], stop_loss, trades_log)
                                break_high_active[a] = False

        return pd.DataFrame(equity_curve), pd.DataFrame(trades_log)

    def enter_trade(self, slots, surplus_pool, a_idx, i, t_type, cp, next_p, stop_loss, trades_log):
        for s_id, info in slots.items():
            if info is None:
                budget_per_slot = 3000000.0

                if t_type == 'Long':
                    if surplus_pool < 1000000: return surplus_pool
                    invest_budget = min(surplus_pool, budget_per_slot)
                    shares = (invest_budget // (next_p * (1 + self.commission_rate)) // 1000) * 1000
                    if shares <= 0: return surplus_pool
                    cost = shares * next_p * (1 + self.commission_rate)
                    surplus_pool -= cost
                else:
                    # Short: receive proceeds minus commission minus tax minus short fee
                    shares = (budget_per_slot // (next_p * (1 + self.commission_rate)) // 1000) * 1000
                    if shares <= 0: return surplus_pool
                    # Seller pays tax (0.3%), commission (0.1425%), and short fee (0.1%)
                    cost = shares * next_p * (1 - self.commission_rate - self.tax_rate - self.short_fee_rate)
                    surplus_pool += cost

                risk = abs(next_p - stop_loss)
                take_profit = next_p + self.rr_target * risk if t_type == 'Long' else next_p - self.rr_target * risk

                slots[s_id] = {
                    'asset_idx': a_idx,
                    'shares': shares,
                    'entry_price': next_p,
                    'stop_loss': stop_loss,
                    'take_profit': take_profit,
                    'type': t_type,
                    'cost': cost
                }
                trades_log.append({
                    'Entry Date': self.dates[i],
                    'Asset': self.assets[a_idx],
                    'Type': t_type,
                    'Entry Price': next_p,
                    'Stop Loss': stop_loss,
                    'Take Profit': take_profit,
                    'Shares': shares
                })
                return surplus_pool
        return surplus_pool

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

if __name__ == "__main__":
    prices, volumes, code_to_name = clean_data('資料-1.xlsx')

    print("Running Long-Only Backtest...")
    bt_l = PhantomBacktester(prices, volumes, code_to_name, mode='L')
    equity_l, trades_l = bt_l.run()

    print("Running Long/Short Backtest...")
    bt_ls = PhantomBacktester(prices, volumes, code_to_name, mode='LS')
    equity_ls, trades_ls = bt_ls.run()

    def print_results(equity, trades, title):
        print("\n" + "="*40)
        print(f"RESULTS: {title}")
        print("="*40)
        if not equity.empty:
            cagr, max_dd, calmar, total_ret = calculate_metrics(equity)
            print(f"Final Equity: {equity.iloc[-1]['權益']:,.2f}")
            print(f"Total Return: {total_ret*100:.2f}%")
            print(f"CAGR: {cagr*100:.2f}%")
            print(f"Max Drawdown: {max_dd*100:.2f}%")
            print(f"Calmar Ratio: {calmar:.2f}")

            if not trades.empty:
                comp = trades[trades['Reason'].notna()]
                print(f"Trades Count: {len(comp)}")
                if len(comp) > 0:
                    win_rate = (comp['PnL'] > 0).mean()
                    print(f"Win Rate: {win_rate*100:.2f}%")

    print_results(equity_l, trades_l, "LONG-ONLY (With Costs)")
    print_results(equity_ls, trades_ls, "LONG/SHORT (With Costs)")
