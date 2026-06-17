import pandas as pd
import numpy as np
import warnings
import os

# 忽略 Pandas 的 Slice 複製警告以維持輸出乾淨
warnings.filterwarnings('ignore')

# ==============================================================================
# 1. 使用者參數設定區塊 (每日/每季更動此處即可)
# ==============================================================================
# 資料來源設定
DATA_FILE = "資料26Q2-1.xlsx"           # 原始資料 Excel 檔名 (驗證時使用資料26Q2-1.xlsx)
LAST_DATE = "2026-06-16"              # 最新日期，資料限制於此日期之前

# 產出報表檔名設定 (可自行修改檔名另存備查)
# 支援 {date_suffix} 佔位符，會自動替換為 LAST_DATE 的 YYMMDD 格式
OUTPUT_FILE = "trendstrategy_results_equityV-adj4-2-{date_suffix}_fixed.xlsx"

# 每季新增標的池集中管理
# 引擎會在生效日之前，將這些標的的價格與成交量排除 (設為 NaN)
NEW_STOCKS_REGISTRY = {
    "2026-04-01": ["3481群創", "6446藥華藥", "2368金像電", "2334華邦電", "3037欣興", "2449京元電", "7769鴻勁"]
}

# 回測期間設定
START_DATE = "2026-01-02"             # 回測紀錄起始日期
END_DATE = None                       # 設為 None 時，會自動以 LAST_DATE 或資料最末日為準

# ==============================================================================
# 延續部位設定 (Warm-Start：從 2025/12/31 持倉延續)
# ==============================================================================
WARM_START_ENABLED = True
WARM_START_CASH = 327_937.48

WARM_START_SLOTS = {
    0: {'code': '3211', 'shares': 29000, 'entry_price': 342.50, 'max_price': 337.00,
        'budget': 9_946_653.81, 'entry_date': '2025-12-29'},
    1: {'code': '3152', 'shares': 66000, 'entry_price': 149.50, 'max_price': 147.00,
        'budget': 9_881_060.48, 'entry_date': '2025-12-29'},
    2: {'code': '3260', 'shares': 39000, 'entry_price': 252.06, 'max_price': 270.45,
        'budget': 9_844_348.23, 'entry_date': '2025-12-29'},
}

# 策略與風險管理參數
INITIAL_TRADING_CAPITAL = 30000000
INITIAL_AUTHORIZED_CAPITAL = 150000000

YEARLY_CAPITAL_REGISTRY = {}

# 指標參數
SMA_PERIOD = 303
ROC_PERIOD = 14
REBALANCE_INTERVAL = 9

# 停損參數
STOP_LOSS_TYPE = 'vol'
STOP_LOSS_VAL = 0.0999
VOL_PERIOD = 15
VOL_MULTIPLIER = 2.7
USE_BREADTH_WEIGHT = True

# 大盤寬度濾網參數
USE_MARKET_FILTER = True
BREADTH_THRESHOLD = 0.42
BREADTH_WINDOW = 290
MKT_SMA_WINDOW = 14

# 交易成本與滑價
COMMISSION_RATE = 0.001425
TAX_RATE = 0.003
SL_SLIPPAGE = 0.0
FILTER_SLIPPAGE = 0.0

# 交易單位參數
MIN_SHARE_UNIT = 1000


# ==============================================================================
# 2. 資料清洗與標的排除輔助函數
# ==============================================================================
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


def apply_new_stocks_registry(prices, volumes, registry):
    """
    根據新標的池註冊表，在生效日期之前排除對應標的（設為 NaN），生效日期起（含）才納入交易池。
    """
    prices_filtered = prices.copy()
    volumes_filtered = volumes.copy()

    for date_str, stocks in registry.items():
        eff_date = pd.to_datetime(date_str)
        # 找出生效日之前的日期遮罩
        pre_mask = prices_filtered.index < eff_date
        if pre_mask.any():
            for stock in stocks:
                # 【錯誤修正】: 原始代碼使用完整的標的名稱(如 "3481群創") 進行比對，
                # 但 Excel 欄位名稱僅包含代碼 (如 "3481")。
                # 修正後將先提取純數字代碼，並同時支援字串與數值型態的欄位標籤。
                stock_code_str = "".join(filter(str.isdigit, str(stock)))

                # 尋找匹配的欄位 (考慮字串或整數型態)
                matched_col = None
                if stock_code_str in prices_filtered.columns:
                    matched_col = stock_code_str
                elif stock_code_str.isdigit() and int(stock_code_str) in prices_filtered.columns:
                    matched_col = int(stock_code_str)

                if matched_col is not None:
                    prices_filtered.loc[pre_mask, matched_col] = np.nan
                    volumes_filtered.loc[pre_mask, matched_col] = np.nan
                else:
                    # 選項：若未找到標的，可記錄日誌或忽略
                    pass

    return prices_filtered, volumes_filtered


# ==============================================================================
# 3. 回測引擎類別 (BacktesterVol)
# ==============================================================================
class BacktesterVol:
    def __init__(self, prices, volumes, code_to_name, trading_capital=30000000, authorized_capital=150000000,
                 warm_start_slots=None, warm_start_cash=None, yearly_capital_registry=None,
                 min_share_unit=1000):
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
        self.yearly_capital_registry = yearly_capital_registry or {}
        self.min_share_unit = min_share_unit
        self.commission_rate = COMMISSION_RATE
        self.tax_rate = TAX_RATE

    def _get_year_capital(self, year):
        if year in self.yearly_capital_registry:
            yr_cfg = self.yearly_capital_registry[year]
            tc = float(yr_cfg.get("trading_capital", self.trading_capital))
            ac = float(yr_cfg.get("authorized_capital", self.authorized_capital))
            return tc, ac
        return self.trading_capital, self.authorized_capital

    def run(self, sma_period, roc_period, stop_loss_type='vol', stop_loss_val=0.0999,
            vol_period=15, vol_multiplier=2.7,
            rebalance_interval=9, use_market_filter=True,
            breadth_threshold=0.42, mkt_sma_window=14, breadth_window=290,
            start_date=None, end_date=None, use_breadth_weight=True,
            sl_slippage=0.0, filter_slippage=0.0):

        min_share_unit = self.min_share_unit
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
                        'asset_idx': a_idx, 'shares': ws['shares'], 'max_price': ws['max_price'],
                        'budget': ws['budget'], 'entry_date': pd.to_datetime(ws['entry_date']),
                        'entry_price': ws['entry_price'], 'entry_reason': "延續部位"
                    }

        equity_curve_data = []
        trades_log = []
        trades2_log = []
        daily_details = []
        equity_hold_details = []

        peak_equity = None
        start_of_year_equity = None
        current_year = self.dates[loop_start].year
        current_trading_cap, current_authorized_cap = self._get_year_capital(current_year)
        current_slot_budget = current_trading_cap / 3.0

        for i in range(loop_start, last_idx + 1):
            date = self.dates[i]
            current_prices = self.prices[i]
            initial_slots = {s_id: (info.copy() if info else None) for s_id, info in slots.items()}

            stock_mv = 0.0
            for s_id, info in initial_slots.items():
                if info:
                    stock_mv += info['shares'] * current_prices[info['asset_idx']]

            total_equity = surplus_pool + stock_mv
            if start_of_year_equity is None: start_of_year_equity = total_equity
            elif date.year != current_year:
                current_year = date.year
                current_trading_cap, current_authorized_cap = self._get_year_capital(current_year)
                current_slot_budget = current_trading_cap / 3.0
                existing_budget_sum = sum(info['budget'] for info in slots.values() if info)
                surplus_pool = current_trading_cap - existing_budget_sum
                total_equity = surplus_pool + stock_mv
                start_of_year_equity = total_equity

            if peak_equity is None or total_equity > peak_equity: peak_equity = total_equity
            drawdown = (total_equity - peak_equity) / peak_equity if peak_equity != 0 else 0
            drawdown_fixed = (total_equity - peak_equity) / current_authorized_cap
            yearly_pnl = total_equity - start_of_year_equity
            yearly_return = yearly_pnl / start_of_year_equity if start_of_year_equity != 0 else 0

            equity_curve_data.append({
                '日期': date, '權益': total_equity, '回撤(Drawdown)': drawdown,
                '固定基準回撤': drawdown_fixed, '市場寬度': breadth[i],
                '年度損益': yearly_pnl, '年度報酬率': yearly_return
            })

            buys_list, sells_list, holds_list = [], [], []
            slot_actions = {s_id: "續抱" for s_id in slots}
            is_rebalance_day = (i - (buffer if self.warm_start_slots else loop_start)) % rebalance_interval == 0
            market_filter_triggered = use_market_filter and not mkt_filter[i]

            if market_filter_triggered:
                for s_id, info in initial_slots.items():
                    if info:
                        a_idx = info['asset_idx']
                        reason = "市場濾網觸發"
                        sells_list.append(f"{self.assets[a_idx]}{self.code_to_name[self.assets[a_idx]]}")
                        slot_actions[s_id] = f"賣出 ({reason})"
                        if i < last_idx:
                            sell_price = self.prices[i+1][a_idx] * (1 - filter_slippage)
                            sell_fee = info['shares'] * sell_price * self.commission_rate
                            sell_tax = info['shares'] * sell_price * self.tax_rate
                            proceeds = info['shares'] * sell_price - sell_fee - sell_tax
                            surplus_pool += proceeds
                            trades2_log.append({
                                '買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx],
                                '股票名稱': self.code_to_name[self.assets[a_idx]],
                                'T+1日買進價格': info['entry_price'], '股數': info['shares'],
                                '賣出訊號日期': date, 'T+1日賣出價格': sell_price, '損益': proceeds - info['budget'],
                                '報酬率': (proceeds / info['budget']) - 1, '買進原因': info['entry_reason'], '賣出原因': reason
                            })
                            trades_log.append({
                                '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出',
                                '價格': sell_price, '股數': info['shares'], '動能值': f"{roc[i][a_idx]*100:.2f}%",
                                '標的名稱': self.code_to_name[self.assets[a_idx]], '原因': '市場濾網',
                                '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax, '說明': f"市場濾網賣出"
                            })
                            slots[s_id] = None
            else:
                exited_slots = set()
                for s_id, info in slots.items():
                    if info:
                        a_idx = info['asset_idx']
                        if current_prices[a_idx] > info['max_price']: info['max_price'] = current_prices[a_idx]
                        exit_triggered, reason_str = False, ""
                        if stop_loss_type == 'fixed':
                            if current_prices[a_idx] < info['max_price'] * (1 - stop_loss_val):
                                exit_triggered, reason_str = True, "固定停損"
                        elif stop_loss_type == 'vol':
                            effective_mult = vol_multiplier * (0.8 if use_breadth_weight and breadth[i] < breadth_threshold else 1.0)
                            stop_price = info['max_price'] * (1 - effective_mult * vol[i][a_idx])
                            if current_prices[a_idx] < stop_price:
                                exit_triggered, reason_str = True, "Vol停損"
                        if exit_triggered:
                            exited_slots.add(s_id)
                            sells_list.append(f"{self.assets[a_idx]}{self.code_to_name[self.assets[a_idx]]}")
                            slot_actions[s_id] = f"賣出 ({reason_str})"
                            if i < last_idx:
                                sell_price = self.prices[i+1][a_idx] * (1 - sl_slippage)
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
                                    '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax, '說明': "停損賣出"
                                })
                                slots[s_id] = None

                if is_rebalance_day:
                    top_3_signals = []
                    sorted_idx = np.argsort(roc[i])[::-1]
                    for idx in sorted_idx:
                        if len(top_3_signals) >= 3: break
                        if np.isnan(roc[i][idx]): continue
                        p, s, r = current_prices[idx], sma[i][idx], roc[i][idx]
                        amount = p * self.volumes[i][idx] * 1000
                        if p > s and r > 0 and amount > 30000000 and p > sma5[i][idx] and p > sma10[i][idx] and p > sma20[i][idx]:
                            top_3_signals.append(idx)

                    for s_id, info in slots.items():
                        if info:
                            if info['asset_idx'] in top_3_signals:
                                holds_list.append(f"{self.assets[info['asset_idx']]}{self.code_to_name[self.assets[info['asset_idx']]]}")
                            else:
                                a_idx = info['asset_idx']
                                sells_list.append(f"{self.assets[a_idx]}{self.code_to_name[self.assets[a_idx]]}")
                                slot_actions[s_id] = "再平衡賣出"
                                if i < last_idx:
                                    sell_price = self.prices[i+1][a_idx]
                                    proceeds = info['shares'] * sell_price * (1 - self.commission_rate - self.tax_rate)
                                    surplus_pool += proceeds
                                    trades2_log.append({
                                        '買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx],
                                        '股票名稱': self.code_to_name[self.assets[a_idx]],
                                        'T+1日買進價格': info['entry_price'], '股數': info['shares'],
                                        '賣出訊號日期': date, 'T+1日賣出價格': sell_price, '損益': proceeds - info['budget'],
                                        '報酬率': (proceeds / info['budget']) - 1, '買進原因': info['entry_reason'], '賣出原因': "再平衡"
                                    })
                                    trades_log.append({
                                        '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出', '原因': '再平衡',
                                        '價格': sell_price, '股數': info['shares'], '動能值': f"{roc[i][a_idx]*100:.2f}%",
                                        '標的名稱': self.code_to_name[self.assets[a_idx]], '說明': "再平衡賣出"
                                    })
                                    slots[s_id] = None

                    for s_id in slots:
                        if slots[s_id] is None and top_3_signals:
                            for sig in top_3_signals:
                                if sig not in [s['asset_idx'] for s in slots.values() if s]:
                                    buys_list.append(f"{self.assets[sig]}{self.code_to_name[self.assets[sig]]}")
                                    if i < last_idx:
                                        buy_price = self.prices[i+1][sig]
                                        shares = (int(current_slot_budget // (buy_price * (1 + self.commission_rate))) // min_share_unit) * min_share_unit
                                        if shares > 0:
                                            cost = shares * buy_price * (1 + self.commission_rate)
                                            surplus_pool -= cost
                                            slots[s_id] = {
                                                'asset_idx': sig, 'shares': shares, 'max_price': buy_price,
                                                'budget': cost, 'entry_date': date, 'entry_price': buy_price,
                                                'entry_reason': f"動能訊號 ROC:{roc[i][sig]*100:.2f}%"
                                            }
                                            trades_log.append({
                                                '訊號日期': date, '股票代號': self.assets[sig], '狀態': '買進', '原因': '符合趨勢',
                                                '價格': buy_price, '股數': shares, '動能值': f"{roc[i][sig]*100:.2f}%",
                                                '標的名稱': self.code_to_name[self.assets[sig]], '說明': "買進"
                                            })
                                    break
                else:
                    for s_id, info in slots.items():
                        if info and s_id not in exited_slots:
                            holds_list.append(f"{self.assets[info['asset_idx']]}{self.code_to_name[self.assets[info['asset_idx']]]}")

            # 記錄 Daily 與 Equity_Hold 明細
            for s_id, info in initial_slots.items():
                if info:
                    a_idx = info['asset_idx']
                    daily_details.append({
                        '日期': date, '股票代號': self.assets[a_idx], '股票名稱': self.code_to_name[self.assets[a_idx]],
                        '持有股數': info['shares'], '本日收盤價': current_prices[a_idx],
                        '市值': info['shares'] * current_prices[a_idx], '次一日交易動作': slot_actions[s_id]
                    })

            h_row = {'日期': date, '持有部位 1': "無", '持有部位 2': "無", '持有部位 3': "無", '次一日交易動作': ""}
            for s_id in range(3):
                if initial_slots.get(s_id):
                    info = initial_slots[s_id]
                    h_row[f'持有部位 {s_id+1}'] = f"{self.assets[info['asset_idx']]}{self.code_to_name[self.assets[info['asset_idx']]]} ({info['shares']:,}股)"

            memo = ""
            if market_filter_triggered: memo = f"本日觸發市場濾網賣出({ '、'.join(sells_list) })"
            elif is_rebalance_day:
                parts = []
                if buys_list: parts.append(f"買進({ '、'.join(buys_list) })")
                if sells_list: parts.append(f"賣出({ '、'.join(sells_list) })")
                if holds_list: parts.append(f"續抱({ '、'.join(holds_list) })")
                if parts: memo = "再平衡日:" + "、".join(parts)
            elif sells_list: memo = f"本日觸發({ '、'.join(sells_list) })停損賣出"
            h_row['次一日交易動作'] = memo
            equity_hold_details.append(h_row)

        return (pd.DataFrame(equity_curve_data), pd.DataFrame(trades_log), pd.DataFrame(trades2_log),
                pd.DataFrame(daily_details), pd.DataFrame(equity_hold_details))


def calculate_metrics_dual(equity_curve_df, trading_cap, authorized_cap):
    if equity_curve_df.empty: return {}
    equity = equity_curve_df['權益']
    days = (equity_curve_df['日期'].iloc[-1] - equity_curve_df['日期'].iloc[0]).days
    years = max(days / 365.25, 0.01)
    trading_cagr = (equity.iloc[-1] / equity.iloc[0]) ** (1 / years) - 1
    authorized_cagr = ((equity.iloc[-1] - equity.iloc[0]) / authorized_cap) / years
    return {
        'Trading CAGR': trading_cagr, 'Authorized CAGR': authorized_cagr,
        'Standard MaxDD': equity_curve_df['回撤(Drawdown)'].min(),
        'Fixed Base MaxDD': equity_curve_df['固定基準回撤'].min(),
        'Trading Calmar': trading_cagr / abs(equity_curve_df['回撤(Drawdown)'].min()) if equity_curve_df['回撤(Drawdown)'].min() != 0 else 0,
        'Yearly Performance': equity_curve_df.groupby(equity_curve_df['日期'].dt.year).last()[['年度報酬率', '年度損益']]
    }


def export_to_excel_premium(eq, t, t2, d, h, metrics, filename):
    print(f"正在產出高品質 Excel 報表：{filename}...")
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        for df, name in [(t, 'Trades'), (t2, 'Trades2'), (eq, 'Equity_Curve'), (h, 'Equity_Hold'), (d, 'Daily')]:
            out = df.copy()
            date_cols = [c for c in out.columns if '日期' in c]
            for c in date_cols: out[c] = pd.to_datetime(out[c]).dt.strftime('%Y/%m/%d')
            out.to_excel(writer, sheet_name=name, index=False)

        summary_data = [
            ['最初投入資金 (Trading Capital)', INITIAL_TRADING_CAPITAL],
            ['初始授權金額 (Authorized Capital)', INITIAL_AUTHORIZED_CAPITAL],
            ['Trading CAGR (30M)', f"{metrics['Trading CAGR']:.2%}"],
            ['Authorized CAGR (150M)', f"{metrics['Authorized CAGR']:.2%}"],
            ['Standard MaxDD', f"{metrics['Standard MaxDD']:.2%}"],
            ['Fixed Base MaxDD (150M)', f"{metrics['Fixed Base MaxDD']:.2%}"],
            ['Trading Calmar', f"{metrics['Trading Calmar']:.2f}"]
        ]
        pd.DataFrame(summary_data, columns=['指標', '數值']).to_excel(writer, sheet_name='Summary', index=False)


def main():
    if not os.path.exists(DATA_FILE): raise FileNotFoundError(f"找不到原始資料檔: {DATA_FILE}")
    prices, volumes, code_to_name = clean_data(DATA_FILE)
    p_f, v_f = apply_new_stocks_registry(prices, volumes, NEW_STOCKS_REGISTRY)

    last_date_dt = pd.to_datetime(LAST_DATE)
    p_f, v_f = p_f.loc[p_f.index <= last_date_dt], v_f.loc[v_f.index <= last_date_dt]

    bt = BacktesterVol(p_f, v_f, code_to_name, INITIAL_TRADING_CAPITAL, INITIAL_AUTHORIZED_CAPITAL,
                       WARM_START_SLOTS if WARM_START_ENABLED else None,
                       WARM_START_CASH if WARM_START_ENABLED else None)
    eq, t, t2, d, h = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_TYPE, STOP_LOSS_VAL, VOL_PERIOD, VOL_MULTIPLIER,
                             REBALANCE_INTERVAL, USE_MARKET_FILTER, BREADTH_THRESHOLD, MKT_SMA_WINDOW, BREADTH_WINDOW,
                             START_DATE, END_DATE, USE_BREADTH_WEIGHT)

    metrics = calculate_metrics_dual(eq, INITIAL_TRADING_CAPITAL, INITIAL_AUTHORIZED_CAPITAL)
    out_name = OUTPUT_FILE.format(date_suffix=pd.to_datetime(LAST_DATE).strftime('%y%m%d'))
    export_to_excel_premium(eq, t, t2, d, h, metrics, out_name)
    print(f"回測完成！報表已儲存至: {out_name}")

if __name__ == "__main__":
    main()
