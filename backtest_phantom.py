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

    # ATR (using RMA as in Pine Script)
    tr = (prices_df - prices_df.shift(1)).abs()
    # Pine Script ATR uses RMA: alpha = 1/length
    atr = tr.ewm(alpha=1/14, min_periods=14, adjust=False).mean()

    return ema200, macd_line, signal_line, macd_hist, atr

class PhantomBacktester:
    def __init__(self, prices, volumes, code_to_name, initial_capital=30000000):
        self.prices_df = prices
        self.volumes_df = volumes
        self.prices = prices.values
        self.dates = prices.index
        self.assets = prices.columns
        self.code_to_name = code_to_name
        self.initial_capital = initial_capital

        # Strategy Parameters
        self.ema_len = 200
        self.macd_fast = 12
        self.macd_slow = 26
        self.macd_signal = 9
        self.pivot_left = 3
        self.pivot_right = 3
        self.min_touches = 2
        self.max_level_age = 100
        self.max_break_closes = 2
        self.atr_frac = 0.25
        self.rr_target = 1.5
        self.use_sr_filter = True
        self.use_reclaim = False

    def run(self):
        print("Calculating indicators...")
        ema200, macd_line, signal_line, macd_hist, atr = calculate_indicators(self.prices_df)

        ema200_v = ema200.values
        macd_line_v = macd_line.values
        signal_line_v = signal_line.values
        atr_v = atr.values

        # Pre-detect pivots
        pivots_low = np.full(self.prices.shape, np.nan)
        pivots_high = np.full(self.prices.shape, np.nan)

        print("Detecting pivots...")
        for a in range(len(self.assets)):
            series = self.prices[:, a]
            for i in range(self.pivot_left, len(series) - self.pivot_right):
                window = series[i - self.pivot_left : i + self.pivot_right + 1]
                if series[i] == np.min(window):
                    pivots_low[i, a] = series[i]
                if series[i] == np.max(window):
                    pivots_high[i, a] = series[i]

        surplus_pool = float(self.initial_capital)
        slots = {i: None for i in range(10)}

        equity_curve = []
        trades_log = []

        asset_low_levels = [[] for _ in range(len(self.assets))]
        asset_high_levels = [[] for _ in range(len(self.assets))]

        start_idx = max(self.ema_len, self.macd_slow, 20)

        print(f"Starting backtest from index {start_idx}...")
        peak_equity = float(self.initial_capital)
        for i in range(start_idx, len(self.dates)):
            curr_prices = self.prices[i]

            # A. Equity Calculation
            # Equity = Cash + Long_MV - Short_MV
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

            # B. Update Levels
            pivot_idx = i - self.pivot_right
            if pivot_idx >= 0:
                for a in range(len(self.assets)):
                    if not np.isnan(pivots_low[pivot_idx, a]):
                        asset_low_levels[a].append((pivots_low[pivot_idx, a], pivot_idx))
                    if not np.isnan(pivots_high[pivot_idx, a]):
                        asset_high_levels[a].append((pivots_high[pivot_idx, a], pivot_idx))

            for a in range(len(self.assets)):
                asset_low_levels[a] = [(lvl, idx) for lvl, idx in asset_low_levels[a] if i - idx <= self.max_level_age * 3]
                asset_high_levels[a] = [(lvl, idx) for lvl, idx in asset_high_levels[a] if i - idx <= self.max_level_age * 3]

            # C. Check for Exit (TP/SL)
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
                            # Sell Long: receive Cash minus commission/tax
                            proceeds = shares * exit_price * (1 - 0.001425 - 0.003)
                            surplus_pool += proceeds
                            pnl = proceeds - info['cost']
                        else:
                            # Buy to Cover Short: pay Cash plus commission
                            cost_to_cover = shares * exit_price * (1 + 0.001425)
                            surplus_pool -= cost_to_cover
                            # PnL = Initial Proceeds - Cost to Cover
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

            # D. Entry Signals
            if any(s is None for s in slots.values()):
                for a in range(len(self.assets)):
                    if any(info and info['asset_idx'] == a for info in slots.values()):
                        continue

                    cp = curr_prices[a]
                    macd_cross_up = macd_line_v[i, a] > signal_line_v[i, a] and macd_line_v[i-1, a] <= signal_line_v[i-1, a]
                    macd_cross_down = macd_line_v[i, a] < signal_line_v[i, a] and macd_line_v[i-1, a] >= signal_line_v[i-1, a]

                    core_long = macd_cross_up and macd_line_v[i, a] < 0 and cp > ema200_v[i, a]
                    core_short = macd_cross_down and macd_line_v[i, a] > 0 and cp < ema200_v[i, a]

                    if not (core_long or core_short): continue

                    if self.use_sr_filter:
                        if core_long:
                            valid_lvl = self.find_latest_valid_level(a, i, True, asset_low_levels[a], atr_v[:, a])
                            if valid_lvl:
                                if not self.use_reclaim or self.check_reclaim(a, i, True, valid_lvl, atr_v[:, a]):
                                    stop_loss = self.find_swing_stop(a, i, True, asset_low_levels[a], ema200_v[i, a])
                                    if np.isnan(stop_loss) or stop_loss >= cp:
                                        stop_loss = cp * 0.95
                                    surplus_pool = self.enter_trade(slots, surplus_pool, a, i, 'Long', cp, next_prices[a], stop_loss, trades_log)
                        elif core_short:
                            valid_lvl = self.find_latest_valid_level(a, i, False, asset_high_levels[a], atr_v[:, a])
                            if valid_lvl:
                                if not self.use_reclaim or self.check_reclaim(a, i, False, valid_lvl, atr_v[:, a]):
                                    stop_loss = self.find_swing_stop(a, i, False, asset_high_levels[a], ema200_v[i, a])
                                    if np.isnan(stop_loss) or stop_loss <= cp:
                                        stop_loss = cp * 1.05
                                    surplus_pool = self.enter_trade(slots, surplus_pool, a, i, 'Short', cp, next_prices[a], stop_loss, trades_log)
                    else:
                        if core_long:
                            surplus_pool = self.enter_trade(slots, surplus_pool, a, i, 'Long', cp, next_prices[a], cp * 0.95, trades_log)
                        elif core_short:
                            surplus_pool = self.enter_trade(slots, surplus_pool, a, i, 'Short', cp, next_prices[a], cp * 1.05, trades_log)

        return pd.DataFrame(equity_curve), pd.DataFrame(trades_log)

    def find_latest_valid_level(self, a_idx, curr_i, is_support, levels, atr_col):
        for lvl, pivot_i in reversed(levels):
            bars_since = curr_i - pivot_i
            if bars_since > self.max_level_age: continue

            touches = 0
            prev_touch = False
            for k in range(pivot_i, curr_i + 1):
                tol = atr_col[k] * self.atr_frac
                p = self.prices[k, a_idx]
                touch = (p <= lvl + tol and p >= lvl - tol)
                if touch and not prev_touch:
                    touches += 1
                prev_touch = touch

            if touches >= self.min_touches:
                return (lvl, pivot_i)
        return None

    def check_reclaim(self, a_idx, curr_i, is_long, level_info, atr_col):
        lvl, pivot_i = level_info
        break_active = False
        break_extrema_open = np.nan
        close_beyond_count = 0

        for k in range(pivot_i + 1, curr_i + 1):
            tol = atr_col[k] * self.atr_frac
            p = self.prices[k, a_idx]
            beyond = p < lvl - tol if is_long else p > lvl + tol

            if not break_active:
                if beyond:
                    break_active = True
                    break_extrema_open = p
                    close_beyond_count = 1
            else:
                break_extrema_open = max(break_extrema_open, p) if is_long else min(break_extrema_open, p)
                if beyond:
                    close_beyond_count += 1

                if close_beyond_count > self.max_break_closes:
                    break_active = beyond
                    break_extrema_open = p if beyond else np.nan
                    close_beyond_count = 1 if beyond else 0
                elif (p > break_extrema_open and p > lvl) if is_long else (p < break_extrema_open and p < lvl):
                    if k == curr_i:
                        return True
                    break_active = False
        return False

    def find_swing_stop(self, a_idx, curr_i, is_long, levels, ema_val):
        swing_lookback = 10
        stop_val = np.nan
        for lvl, pivot_i in reversed(levels):
            if curr_i - pivot_i <= swing_lookback:
                if (is_long and lvl < ema_val) or (not is_long and lvl > ema_val):
                    stop_val = lvl
                    break
            else:
                break
        return stop_val

    def enter_trade(self, slots, surplus_pool, a_idx, i, t_type, cp, next_p, stop_loss, trades_log):
        for s_id, info in slots.items():
            if info is None:
                # Taiwan context: 3M per slot
                budget_per_slot = 3000000.0

                if t_type == 'Long':
                    if surplus_pool < 1000000: return surplus_pool
                    invest_budget = min(surplus_pool, budget_per_slot)
                    shares = (invest_budget // (next_p * 1.001425) // 1000) * 1000
                    if shares <= 0: return surplus_pool
                    cost = shares * next_p * 1.001425
                    surplus_pool -= cost
                else:
                    # Short: received cash minus commission.
                    # Note: Simplified, not considering margin requirement for cash pool.
                    # We use budget as a size limit.
                    shares = (budget_per_slot // (next_p * 1.001425) // 1000) * 1000
                    if shares <= 0: return surplus_pool
                    # Cost here is the initial proceeds
                    cost = shares * next_p * (1 - 0.001425)
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
                    'cost': cost # cost/proceeds
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
    bt = PhantomBacktester(prices, volumes, code_to_name)
    equity, trades = bt.run()

    print("\n" + "="*30)
    print("BACKTEST RESULTS (Phantom Strategy)")
    print("="*30)
    if not equity.empty:
        cagr, max_dd, calmar, total_ret = calculate_metrics(equity)
        print(f"Initial Capital: {bt.initial_capital:,.2f}")
        print(f"Final Equity: {equity.iloc[-1]['權益']:,.2f}")
        print(f"Total Return: {total_ret*100:.2f}%")
        print(f"CAGR: {cagr*100:.2f}%")
        print(f"Max Drawdown: {max_dd*100:.2f}%")
        print(f"Calmar Ratio: {calmar:.2f}")

    if not trades.empty:
        completed_trades = trades[trades['Reason'].notna()]
        print(f"Total Completed Trades: {len(completed_trades)}")
        if len(completed_trades) > 0:
            win_rate = (completed_trades['PnL'] > 0).mean()
            print(f"Win Rate: {win_rate*100:.2f}%")
    else:
        print("No trades executed.")
