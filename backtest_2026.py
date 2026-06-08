import pandas as pd
import numpy as np
import warnings
import os
import re

# 忽略 Pandas 的 Slice 複製警告以維持輸出乾淨
warnings.filterwarnings('ignore')

# ==============================================================================
# 1. 使用者參數設定區塊
# ==============================================================================
DATA_FILE = "資料26Q2-2.xlsx"
LAST_DATE = "2026-06-08"

# 每季新增標的池集中管理
NEW_STOCKS_REGISTRY = {
    "2026-04-01": ["3481群創", "6446藥華藥", "2368金像電", "2344華邦電", "3037欣興", "2449京元電", "7769鴻勁"]
}

START_DATE = "2026-01-02"
END_DATE = "2026-06-08"

# ==============================================================================
# 延續部位設定 (Warm-Start)
# ==============================================================================
WARM_START_ENABLED = True
# 數據來源：現金 = 總權益(157,948,201) - 市值總計(30,022,550) = 127,925,651
WARM_START_CASH = 127925651.0

WARM_START_SLOTS = {
    0: {'code': '3211', 'shares': 29000, 'entry_price': 342.50, 'max_price': 337.00,
        'budget': 9_946_653.81, 'entry_date': '2025-12-29'},
    1: {'code': '3152', 'shares': 66000, 'entry_price': 149.50, 'max_price': 147.00,
        'budget': 9_881_060.48, 'entry_date': '2025-12-29'},
    2: {'code': '3260', 'shares': 39000, 'entry_price': 252.06, 'max_price': 270.45,
        'budget': 9_844_348.23, 'entry_date': '2025-12-29'},
}

INITIAL_TRADING_CAPITAL = 30000000
INITIAL_AUTHORIZED_CAPITAL = 150000000

SMA_PERIOD = 303
ROC_PERIOD = 14
REBALANCE_INTERVAL = 9

STOP_LOSS_TYPE = 'fixed'
STOP_LOSS_VAL = 0.0999
VOL_PERIOD = 15
VOL_MULTIPLIER = 2.7
USE_BREADTH_WEIGHT = False

USE_MARKET_FILTER = False
BREADTH_THRESHOLD = 0.42
BREADTH_WINDOW = 290
MKT_SMA_WINDOW = 14

COMMISSION_RATE = 0.001425
TAX_RATE = 0.003
SL_SLIPPAGE = 0.0
FILTER_SLIPPAGE = 0.0

# ==============================================================================
# 2. 輔助函數
# ==============================================================================
def extract_code(stock_str):
    match = re.search(r'\d+', stock_str)
    return match.group() if match else stock_str

def clean_data(filepath):
    df_prices = pd.read_excel(filepath, sheet_name='還原收盤價', header=None)
    df_volume = pd.read_excel(filepath, sheet_name='成交量', header=None)

    stock_codes = df_prices.iloc[0, 1:].astype(str).values
    stock_names = df_prices.iloc[1, 1:].values

    date_strings = df_prices.iloc[2:, 0].astype(str).str.extract(r'(\d+)')[0]
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

def apply_new_stocks_registry(prices, volumes, registry):
    prices_filtered = prices.copy()
    volumes_filtered = volumes.copy()
    for date_str, stocks in registry.items():
        eff_date = pd.to_datetime(date_str)
        pre_mask = prices_filtered.index < eff_date
        if pre_mask.any():
            for stock in stocks:
                code = extract_code(stock)
                if code in prices_filtered.columns:
                    prices_filtered.loc[pre_mask, code] = np.nan
                if code in volumes_filtered.columns:
                    volumes_filtered.loc[pre_mask, code] = np.nan
    return prices_filtered, volumes_filtered

# ==============================================================================
# 3. 回測引擎
# ==============================================================================
class BacktesterVol:
    def __init__(self, prices, volumes, code_to_name, trading_capital=30000000, authorized_capital=150000000,
                 warm_start_slots=None, warm_start_cash=None):
        self.prices_df = prices
        self.volumes_df = volumes
        self.prices = prices.values
        self.volumes = volumes.values
        self.dates = prices.index
        self.assets = prices.columns
        self.code_to_name = code_to_name
        self.trading_capital = float(trading_capital)
        self.authorized_capital = float(authorized_capital)
        self.warm_start_slots = warm_start_slots
        self.warm_start_cash = warm_start_cash
        self.commission_rate = COMMISSION_RATE
        self.tax_rate = TAX_RATE

    def run(self, sma_period, roc_period, stop_loss_type='vol', stop_loss_val=0.0999,
            vol_period=15, vol_multiplier=2.7,
            rebalance_interval=9, use_market_filter=True,
            breadth_threshold=0.42, mkt_sma_window=14, breadth_window=290,
            start_date=None, end_date=None, use_breadth_weight=True,
            sl_slippage=0.0, filter_slippage=0.0):

        sma = self.prices_df.rolling(window=sma_period).mean().values
        roc = self.prices_df.pct_change(periods=roc_period).values
        sma5 = self.prices_df.rolling(window=5).mean().values
        sma10 = self.prices_df.rolling(window=10).mean().values
        sma20 = self.prices_df.rolling(window=20).mean().values
        returns = self.prices_df.pct_change()
        vol = returns.rolling(window=vol_period).std().values
        breadth_sma_all = self.prices_df.rolling(window=breadth_window).mean().values
        breadth = np.nanmean(np.where(np.isnan(self.prices), np.nan, self.prices > breadth_sma_all), axis=1)
        market_avg = self.prices_df.mean(axis=1).values
        market_sma = self.prices_df.mean(axis=1).rolling(window=mkt_sma_window).mean().values
        mkt_filter = (breadth >= breadth_threshold) | (market_avg >= market_sma)

        if start_date:
            mask = self.dates >= pd.to_datetime(start_date)
            first_idx = np.where(mask)[0][0] if any(mask) else 0
        else:
            first_idx = 0
        if end_date:
            mask = self.dates <= pd.to_datetime(end_date)
            last_idx = np.where(mask)[0][-1] if any(mask) else len(self.dates)-1
        else:
            last_idx = len(self.dates)-1

        buffer = max(sma_period, roc_period, breadth_window, mkt_sma_window, vol_period + 1)
        loop_start = max(first_idx, buffer)

        if self.warm_start_cash is not None:
            surplus_pool = float(self.warm_start_cash)
        else:
            surplus_pool = float(self.trading_capital)

        slots = {0: None, 1: None, 2: None}
        if self.warm_start_slots:
            asset_code_to_idx = {str(code): idx for idx, code in enumerate(self.assets)}
            for s_id, ws in self.warm_start_slots.items():
                code_str = str(ws['code'])
                if code_str in asset_code_to_idx:
                    a_idx = asset_code_to_idx[code_str]
                    slots[s_id] = {
                        'asset_idx': a_idx,
                        'shares': ws['shares'],
                        'max_price': ws['max_price'],
                        'budget': ws['budget'],
                        'entry_date': pd.to_datetime(ws['entry_date']),
                        'entry_price': ws['entry_price'],
                        'entry_reason': f"延續部位，entry={ws['entry_price']}"
                    }

        equity_curve_data = []
        trades_log = []
        trades2_log = []
        daily_details = []
        equity_hold_details = []

        if self.warm_start_slots and self.warm_start_cash is not None:
            init_mv = sum(ws['shares'] * self.prices[loop_start][asset_code_to_idx[str(ws['code'])]] for ws in self.warm_start_slots.values())
            peak_equity = float(self.warm_start_cash) + init_mv
        else:
            peak_equity = float(self.trading_capital)

        start_of_year_equity = None
        current_year = self.dates[loop_start].year

        for i in range(loop_start, last_idx + 1):
            date = self.dates[i]
            current_prices = self.prices[i]

            initial_slots = {s_id: (info.copy() if info else None) for s_id, info in slots.items()}
            stock_mv = 0.0
            for s_id, info in initial_slots.items():
                if info and 'asset_idx' in info:
                    a_idx = info['asset_idx']
                    mv = info['shares'] * current_prices[a_idx]
                    stock_mv += mv

            total_equity = surplus_pool + stock_mv
            if start_of_year_equity is None:
                start_of_year_equity = total_equity
            elif date.year != current_year:
                start_of_year_equity = total_equity
                current_year = date.year

            if total_equity > peak_equity:
                peak_equity = total_equity

            drawdown = (total_equity - peak_equity) / peak_equity if peak_equity != 0 else 0
            drawdown_fixed = (total_equity - peak_equity) / self.authorized_capital
            yearly_pnl = total_equity - start_of_year_equity
            yearly_return = yearly_pnl / start_of_year_equity if start_of_year_equity != 0 else 0

            equity_curve_data.append({
                '日期': date, '權益': total_equity, '回撤(Drawdown)': drawdown,
                '固定基準回撤': drawdown_fixed, '市場寬度': breadth[i],
                '年度損益': yearly_pnl, '年度報酬率': yearly_return
            })

            buys_list, sells_list, holds_list = [], [], []
            slot_actions = {s_id: "續抱" for s_id in slots}
            market_filter_triggered = use_market_filter and not mkt_filter[i]

            if market_filter_triggered:
                for s_id, info in initial_slots.items():
                    if info and 'asset_idx' in info:
                        a_idx = info['asset_idx']
                        reason = f"市場濾網：寬度({breadth[i]:.1%})與大盤皆弱"
                        sells_list.append(f"{self.assets[a_idx]}{self.code_to_name[self.assets[a_idx]]}")
                        slot_actions[s_id] = f"賣出 ({reason})"

                if i < last_idx:
                    next_prices = self.prices[i+1]
                    for s_id, info in slots.items():
                        if info and 'asset_idx' in info:
                            a_idx = info['asset_idx']
                            sell_price = next_prices[a_idx] * (1 - filter_slippage)
                            sell_fee = info['shares'] * sell_price * self.commission_rate
                            sell_tax = info['shares'] * sell_price * self.tax_rate
                            proceeds = info['shares'] * sell_price - sell_fee - sell_tax
                            surplus_pool += proceeds
                            trades2_log.append({
                                '買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx],
                                '股票名稱': self.code_to_name[self.assets[a_idx]],
                                'T+1日買進價格': info['entry_price'], '股數': info['shares'],
                                '賣出訊號日期': date, 'T+1日賣出價格': sell_price, '損益': proceeds - info['budget'],
                                '報酬率': (proceeds / info['budget']) - 1, '買進原因': info['entry_reason'],
                                '賣出原因': f"市場濾網：寬度({breadth[i]:.1%})與大盤皆弱"
                            })
                            trades_log.append({
                                '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出',
                                '價格': sell_price, '股數': info['shares'], '動能值': f"{roc[i][a_idx]*100:.2f}%",
                                '標的名稱': self.code_to_name[self.assets[a_idx]], '原因': '市場濾網',
                                '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax,
                                '說明': f"市場濾網賣出：{self.code_to_name[self.assets[a_idx]]}"
                            })
                            slots[s_id] = None
            else:
                exited_slots = set()
                for s_id, info in slots.items():
                    if info and 'asset_idx' in info:
                        a_idx = info['asset_idx']
                        if current_prices[a_idx] > info['max_price']:
                            info['max_price'] = current_prices[a_idx]
                        exit_triggered = False
                        reason_str = ""
                        if stop_loss_type == 'fixed':
                            if current_prices[a_idx] < info['max_price'] * (1 - stop_loss_val):
                                exit_triggered = True
                                reason_str = f"固定停損：最高點回落{stop_loss_val*100:.2f}%"
                        elif stop_loss_type == 'vol':
                            current_vol = vol[i][a_idx]
                            effective_mult = vol_multiplier * (0.8 if use_breadth_weight and breadth[i] < breadth_threshold else 1.0)
                            stop_price = info['max_price'] * (1 - effective_mult * current_vol)
                            if current_prices[a_idx] < stop_price:
                                exit_triggered = True
                                reason_str = f"Vol停損：低於停損價{stop_price:.2f}"

                        if exit_triggered:
                            exited_slots.add(s_id)
                            sells_list.append(f"{self.assets[a_idx]}{self.code_to_name[self.assets[a_idx]]}(停損)")
                            slot_actions[s_id] = f"賣出 ({reason_str})"
                            if i < last_idx:
                                next_prices = self.prices[i+1]
                                sell_price = next_prices[a_idx] * (1 - sl_slippage)
                                sell_fee = info['shares'] * sell_price * self.commission_rate
                                sell_tax = info['shares'] * sell_price * self.tax_rate
                                proceeds = info['shares'] * sell_price - sell_fee - sell_tax
                                surplus_pool += proceeds
                                trades2_log.append({
                                    '買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx],
                                    '股票名稱': self.code_to_name[self.assets[a_idx]],
                                    'T+1日買進價格': info['entry_price'], '股數': info['shares'],
                                    '賣出訊號日期': date, 'T+1日賣出價格': sell_price, '損益': proceeds - info['budget'],
                                    '報酬率': (proceeds / info['budget']) - 1, '買進原因': info['entry_reason'], '賣出原因': reason_str
                                })
                                trades_log.append({
                                    '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出',
                                    '價格': sell_price, '股數': info['shares'], '動能值': f"{roc[i][a_idx]*100:.2f}%",
                                    '標的名稱': self.code_to_name[self.assets[a_idx]], '原因': '停損',
                                    '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax,
                                    '說明': f"停損賣出：{self.code_to_name[self.assets[a_idx]]}"
                                })
                                slots[s_id] = None

                is_rebalance_day = (i - loop_start) % rebalance_interval == 0
                if is_rebalance_day:
                    top_3_signals = []
                    valid_roc = roc[i]
                    sorted_all = np.argsort(valid_roc)[::-1]
                    for idx in sorted_all:
                        if len(top_3_signals) >= 3: break
                        if np.isnan(valid_roc[idx]): continue
                        p, s, r = current_prices[idx], sma[i][idx], roc[i][idx]
                        amount = p * self.volumes[i][idx] * 1000
                        c2 = p > sma5[i][idx] and p > sma10[i][idx] and p > sma20[i][idx]
                        if p > s and r > 0 and amount > 30000000 and c2:
                            top_3_signals.append(idx)

                    signal_to_slot_map = {}
                    for s_id, info in slots.items():
                        if info and 'asset_idx' in info:
                            if info['asset_idx'] in top_3_signals:
                                signal_to_slot_map[info['asset_idx']] = s_id
                                holds_list.append(f"{self.assets[info['asset_idx']]}{self.code_to_name[self.assets[info['asset_idx']]]}")
                                slot_actions[s_id] = "續抱 (動能維持前三名)"
                            else:
                                a_idx = info['asset_idx']
                                sells_list.append(f"{self.assets[a_idx]}{self.code_to_name[self.assets[a_idx]]}(再平衡)")
                                slot_actions[s_id] = "賣出 (再平衡：動能跌出前三名)"
                                if i < last_idx:
                                    next_prices = self.prices[i+1]
                                    sell_price = next_prices[a_idx]
                                    sell_fee = info['shares'] * sell_price * self.commission_rate
                                    sell_tax = info['shares'] * sell_price * self.tax_rate
                                    proceeds = info['shares'] * sell_price - sell_fee - sell_tax
                                    surplus_pool += proceeds
                                    trades2_log.append({
                                        '買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx],
                                        '股票名稱': self.code_to_name[self.assets[a_idx]],
                                        'T+1日買進價格': info['entry_price'], '股數': info['shares'],
                                        '賣出訊號日期': date, 'T+1日賣出價格': sell_price, '損益': proceeds - info['budget'],
                                        '報酬率': (proceeds / info['budget']) - 1, '買進原因': info['entry_reason'],
                                        '賣出原因': f"再平衡賣出：排名外"
                                    })
                                    trades_log.append({
                                        '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出',
                                        '價格': sell_price, '股數': info['shares'], '動能值': f"{roc[i][a_idx]*100:.2f}%",
                                        '標的名稱': self.code_to_name[self.assets[a_idx]], '原因': '再平衡',
                                        '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax,
                                        '說明': f"再平衡賣出：{self.code_to_name[self.assets[a_idx]]}"
                                    })
                                    slots[s_id] = None

                    for s_id in slots:
                        if slots[s_id] is None and top_3_signals:
                            next_sig = None
                            for sig in top_3_signals:
                                if sig not in [slots[sid]['asset_idx'] for sid in slots if slots[sid]]:
                                    next_sig = sig
                                    break
                            if next_sig is not None:
                                buys_list.append(f"{self.assets[next_sig]}{self.code_to_name[self.assets[next_sig]]}")
                                if i < last_idx:
                                    next_prices = self.prices[i+1]
                                    budget = 10000000.0
                                    buy_price_exec = next_prices[next_sig]
                                    shares = (int(budget // (buy_price_exec * (1 + self.commission_rate))) // 1000) * 1000
                                    if shares > 0:
                                        actual_cost = shares * buy_price_exec * (1 + self.commission_rate)
                                        buy_fee = shares * buy_price_exec * self.commission_rate
                                        surplus_pool -= actual_cost
                                        slots[s_id] = {
                                            'asset_idx': next_sig, 'shares': shares, 'max_price': buy_price_exec,
                                            'budget': actual_cost, 'entry_date': date, 'entry_price': buy_price_exec,
                                            'entry_reason': f"符合趨勢與濾網，ROC:{roc[i][next_sig]*100:.2f}%"
                                        }
                                        trades_log.append({
                                            '訊號日期': date, '股票代號': self.assets[next_sig], '狀態': '買進',
                                            '價格': buy_price_exec, '股數': shares, '動能值': f"{roc[i][next_sig]*100:.2f}%",
                                            '標的名稱': self.code_to_name[self.assets[next_sig]], '原因': '符合趨勢',
                                            '買入手續費': buy_fee, '賣出手續費': 0, '賣出交易稅': 0,
                                            '說明': f"買進：{self.code_to_name[self.assets[next_sig]]}"
                                        })
                    if i < last_idx:
                        for sig, s_id in signal_to_slot_map.items():
                            trades_log.append({
                                '訊號日期': date, '股票代號': self.assets[sig], '狀態': '保持',
                                '價格': current_prices[sig], '股數': slots[s_id]['shares'],
                                '動能值': f"{roc[i][sig]*100:.2f}%", '標的名稱': self.code_to_name[self.assets[sig]], '原因': '趨勢持續',
                                '買入手續費': 0, '賣出手續費': 0, '賣出交易稅': 0, '說明': f"續抱：{self.code_to_name[self.assets[sig]]}"
                            })
                else:
                    for s_id, info in slots.items():
                        if info and 'asset_idx' in info and s_id not in exited_slots:
                            holds_list.append(f"{self.assets[info['asset_idx']]}{self.code_to_name[self.assets[info['asset_idx']]]}")
                            slot_actions[s_id] = "續抱"

            action_parts = []
            if buys_list: action_parts.append("【買進】" + "、".join(buys_list))
            if sells_list: action_parts.append("【賣出】" + "、".join(sells_list))
            if holds_list: action_parts.append("【續抱】" + "、".join(holds_list))
            action_memo = " | ".join(action_parts) if action_parts else "無動作"

            for s_id, info in initial_slots.items():
                if info and 'asset_idx' in info:
                    a_idx = info['asset_idx']
                    mv = info['shares'] * current_prices[a_idx]
                    daily_details.append({
                        '日期': date, '股票代號': self.assets[a_idx], '股票名稱': self.code_to_name[self.assets[a_idx]],
                        '持有股數': info['shares'], '本日收盤價': current_prices[a_idx], '市值': mv,
                        '次一日交易動作': slot_actions[s_id]
                    })
            holdings_str = []
            for s_id in range(3):
                info = initial_slots.get(s_id)
                if info and 'asset_idx' in info:
                    holdings_str.append(f"{self.assets[info['asset_idx']]}{self.code_to_name[self.assets[info['asset_idx']]]} ({info['shares']:,}股)")
                else:
                    holdings_str.append("無")
            equity_hold_details.append({
                '日期': date, '持有部位 1': holdings_str[0], '持有部位 2': holdings_str[1],
                '持有部位 3': holdings_str[2], '次一日交易動作': action_memo
            })

        return (pd.DataFrame(equity_curve_data), pd.DataFrame(trades_log), pd.DataFrame(trades2_log),
                pd.DataFrame(daily_details), pd.DataFrame(equity_hold_details))

def calculate_metrics_dual(equity_curve_df, trading_cap, authorized_cap):
    if equity_curve_df.empty: return {}
    equity = equity_curve_df['權益']
    total_gain = equity.iloc[-1] - equity.iloc[0]
    days = (equity_curve_df['日期'].iloc[-1] - equity_curve_df['日期'].iloc[0]).days
    years = days / 365.25
    trading_total_return = total_gain / trading_cap
    trading_cagr = (1 + trading_total_return) ** (1 / years) - 1 if years > 0 else 0
    authorized_total_return = total_gain / authorized_cap
    authorized_cagr = (1 + authorized_total_return) ** (1 / years) - 1 if years > 0 else 0
    max_dd = equity_curve_df['回撤(Drawdown)'].min()
    max_fixed_dd = equity_curve_df['固定基準回撤'].min()
    trading_calmar = trading_cagr / abs(max_dd) if max_dd != 0 else 0
    yearly_perf = equity_curve_df.groupby(equity_curve_df['日期'].dt.year).last()[['年度報酬率', '年度損益']]
    return {
        'Trading CAGR': trading_cagr, 'Authorized CAGR': authorized_cagr, 'Standard MaxDD': max_dd,
        'Fixed Base MaxDD': max_fixed_dd, 'Trading Calmar': trading_calmar, 'Yearly Performance': yearly_perf
    }

def export_to_excel_premium(equity_df, trades_df, trades2_df, daily_df, hold_df, metrics, filename):
    print(f"正在產出高品質 Excel 報表：{filename}...")
    eq_out, t_out, t2_out, d_out, h_out = equity_df.copy(), trades_df.copy(), trades2_df.copy(), daily_df.copy(), hold_df.copy()
    for df in [eq_out, t_out, t2_out, d_out, h_out]:
        if not df.empty:
            date_col = '日期' if '日期' in df.columns else ('訊號日期' if '訊號日期' in df.columns else '買進訊號日期')
            df[date_col] = pd.to_datetime(df[date_col]).dt.strftime('%Y/%m/%d')
            if '賣出訊號日期' in df.columns: df['賣出訊號日期'] = pd.to_datetime(df['賣出訊號日期']).dt.strftime('%Y/%m/%d')

    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        t_out.to_excel(writer, sheet_name='Trades', index=False)
        t2_out.to_excel(writer, sheet_name='Trades2', index=False)
        eq_out.to_excel(writer, sheet_name='Equity_Curve', index=False)
        h_out.to_excel(writer, sheet_name='Equity_Hold', index=False)
        d_out.to_excel(writer, sheet_name='Daily', index=False)
        summary_data = [
            ['策略指標 (全期間)', '數值'],
            ['最初投入資金 (Trading Capital)', INITIAL_TRADING_CAPITAL],
            ['初始授權金額 (Authorized Capital)', INITIAL_AUTHORIZED_CAPITAL],
            ['Trading CAGR (30M)', metrics['Trading CAGR']],
            ['Authorized CAGR (150M)', metrics['Authorized CAGR']],
            ['Standard MaxDD', metrics['Standard MaxDD']],
            ['Fixed Base MaxDD (150M)', metrics['Fixed Base MaxDD']],
            ['Trading Calmar', metrics['Trading Calmar']],
            ['', ''],
            ['年度績效 (實戰模式)', '年度報酬率', '年度損益 (TWD)', '年度MDD (150M基準)']
        ]
        equity_raw = equity_df.copy()
        equity_raw['Year'] = pd.to_datetime(equity_raw['日期']).dt.year
        for year, row in metrics['Yearly Performance'].iterrows():
            year_mdd_fixed = equity_raw[equity_raw['Year'] == year]['固定基準回撤'].min()
            summary_data.append([f"{int(year)} 年度", row['年度報酬率'], row['年度損益'], year_mdd_fixed])
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False, header=False)
        workbook, font_name = writer.book, 'Microsoft JhengHei'
        header_fmt = workbook.add_format({'bold': True, 'font_name': font_name, 'font_color': '#FFFFFF', 'bg_color': '#1B365D', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        num_fmt = workbook.add_format({'num_format': '#,##0', 'font_name': font_name})
        pct_fmt = workbook.add_format({'num_format': '0.00%', 'font_name': font_name})
        text_fmt = workbook.add_format({'font_name': font_name})
        summary_title_fmt = workbook.add_format({'bold': True, 'font_name': font_name, 'font_size': 12, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'left'})
        summary_section_fmt = workbook.add_format({'bold': True, 'font_name': font_name, 'font_size': 11, 'bg_color': '#F2F2F2', 'border': 1, 'align': 'left'})
        summary_sheet = writer.sheets['Summary']
        summary_sheet.set_column('A:A', 35); summary_sheet.set_column('B:D', 20)
        summary_sheet.write('A1', '策略指標 (全期間)', summary_title_fmt); summary_sheet.write('B1', '數值', summary_title_fmt)
        summary_sheet.write('A2', '最初投入資金 (Trading Capital)', text_fmt); summary_sheet.write('B2', INITIAL_TRADING_CAPITAL, num_fmt)
        summary_sheet.write('A3', '初始授權金額 (Authorized Capital)', text_fmt); summary_sheet.write('B3', INITIAL_AUTHORIZED_CAPITAL, num_fmt)
        summary_sheet.write('A4', 'Trading CAGR (30M)', text_fmt); summary_sheet.write('B4', metrics['Trading CAGR'], pct_fmt)
        summary_sheet.write('A5', 'Authorized CAGR (150M)', text_fmt); summary_sheet.write('B5', metrics['Authorized CAGR'], pct_fmt)
        summary_sheet.write('A6', 'Standard MaxDD', text_fmt); summary_sheet.write('B6', metrics['Standard MaxDD'], pct_fmt)
        summary_sheet.write('A7', 'Fixed Base MaxDD (150M)', text_fmt); summary_sheet.write('B7', metrics['Fixed Base MaxDD'], pct_fmt)
        summary_sheet.write('A8', 'Trading Calmar', text_fmt); summary_sheet.write('B8', metrics['Trading Calmar'], workbook.add_format({'num_format': '0.00', 'font_name': font_name}))
        summary_sheet.write('A10', '年度績效 (實戰模式)', summary_section_fmt); summary_sheet.write('B10', '年度報酬率', summary_section_fmt); summary_sheet.write('C10', '年度損益 (TWD)', summary_section_fmt); summary_sheet.write('D10', '年度MDD (150M基準)', summary_section_fmt)
        row_idx = 10
        for year, row in metrics['Yearly Performance'].iterrows():
            year_mdd_fixed = equity_raw[equity_raw['Year'] == year]['固定基準回撤'].min()
            summary_sheet.write(row_idx, 0, f"{int(year)} 年度", text_fmt); summary_sheet.write(row_idx, 1, row['年度報酬率'], pct_fmt); summary_sheet.write(row_idx, 2, row['年度損益'], num_fmt); summary_sheet.write(row_idx, 3, year_mdd_fixed, pct_fmt); row_idx += 1
        sheet_mapping = {'Trades': t_out, 'Trades2': t2_out, 'Equity_Curve': eq_out, 'Equity_Hold': h_out, 'Daily': d_out}
        for sheet_name, df in sheet_mapping.items():
            sheet = writer.sheets[sheet_name]
            for col_num, col_name in enumerate(df.columns): sheet.write(0, col_num, col_name, header_fmt)
            for col_num, col_name in enumerate(df.columns):
                max_len = min(max(max(df[col_name].astype(str).map(len).max(), len(col_name)) + 4, 10), 65)
                if any(x in col_name for x in ['報酬率', '回撤', '寬度', '動能值']): sheet.set_column(col_num, col_num, max_len, pct_fmt)
                elif any(x in col_name for x in ['權益', '損益', '價格', '收盤價', '市值', '股數', '手續費', '交易稅']): sheet.set_column(col_num, col_num, max_len, num_fmt)
                else: sheet.set_column(col_num, col_num, max_len, text_fmt)
        curve_sheet, chart = writer.sheets['Equity_Curve'], workbook.add_chart({'type': 'line'})
        max_row = len(equity_df)
        chart.add_series({'name': '權益 (Trading Capital 30M)', 'categories': ['Equity_Curve', 1, 0, max_row, 0], 'values': ['Equity_Curve', 1, 1, max_row, 1], 'line': {'color': '#1B365D'}})
        chart.set_title({'name': '權益增長趨勢曲線', 'name_font': {'name': font_name, 'size': 14, 'bold': True}}); chart.set_x_axis({'name': '日期'}); chart.set_y_axis({'name': '權益 (TWD)'}); chart.set_legend({'position': 'bottom'}); chart.set_size({'width': 760, 'height': 420}); curve_sheet.insert_chart('H2', chart)

def main():
    if not os.path.exists(DATA_FILE): raise FileNotFoundError(f"找不到資料檔: {DATA_FILE}")
    prices_raw, volumes_raw, code_to_name = clean_data(DATA_FILE)
    prices_filtered, volumes_filtered = apply_new_stocks_registry(prices_raw, volumes_raw, NEW_STOCKS_REGISTRY)
    if LAST_DATE:
        last_date_dt = pd.to_datetime(LAST_DATE)
        prices_filtered = prices_filtered.loc[prices_filtered.index <= last_date_dt]
        volumes_filtered = volumes_filtered.loc[volumes_filtered.index <= last_date_dt]
    bt = BacktesterVol(prices_filtered, volumes_filtered, code_to_name, trading_capital=INITIAL_TRADING_CAPITAL, authorized_capital=INITIAL_AUTHORIZED_CAPITAL,
                       warm_start_slots=WARM_START_SLOTS if WARM_START_ENABLED else None, warm_start_cash=WARM_START_CASH if WARM_START_ENABLED else None)
    eq, t, t2, d, h = bt.run(sma_period=SMA_PERIOD, roc_period=ROC_PERIOD, stop_loss_type=STOP_LOSS_TYPE, stop_loss_val=STOP_LOSS_VAL, vol_period=VOL_PERIOD, vol_multiplier=VOL_MULTIPLIER,
                             rebalance_interval=REBALANCE_INTERVAL, use_market_filter=USE_MARKET_FILTER, breadth_threshold=BREADTH_THRESHOLD, mkt_sma_window=MKT_SMA_WINDOW, breadth_window=BREADTH_WINDOW,
                             start_date=START_DATE, end_date=END_DATE, use_breadth_weight=USE_BREADTH_WEIGHT, sl_slippage=SL_SLIPPAGE, filter_slippage=FILTER_SLIPPAGE)
    metrics = calculate_metrics_dual(eq, INITIAL_TRADING_CAPITAL, INITIAL_AUTHORIZED_CAPITAL)
    date_suffix = pd.to_datetime(LAST_DATE).strftime('%y%m%d')
    output_filename = f"trendstrategy_results_equityV-adj4(舊)-{date_suffix}.xlsx"
    export_to_excel_premium(eq, t, t2, d, h, metrics, output_filename)
    print(f"完成！報表: {output_filename}, CAGR: {metrics['Trading CAGR']:.2%}, MaxDD: {metrics['Standard MaxDD']:.2%}")

if __name__ == "__main__":
    main()
