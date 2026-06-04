import pandas as pd
import numpy as np
import warnings
import os
os.chdir(r"d:\ANTI\B303\修正3")
CLEAN_DATA = "資料26Q2-1.xlsx"

# 忽略 Pandas 的 Slice 複製警告以維持輸出乾淨
warnings.filterwarnings('ignore')

# ==============================================================================
# 1. 使用者參數設定區塊 (每日/每季更動此處即可)
# ==============================================================================
# 資料來源設定
DATA_FILE = "資料26Q2-1.xlsx"           # 原始資料 Excel 檔名 (對應需求 5)
LAST_DATE = "2025-12-31"              # 最新日期，資料限制於此日期之前 (對應需求 3)

# 每季新增標的池集中管理 (對應需求 2)
# 格式為 { "生效日期 YYYY-MM-DD": [ "標的代碼+名稱", ... ] }
# 引擎會在生效日之前，將這些標的的價格與成交量排除 (設為 NaN)
NEW_STOCKS_REGISTRY = {
    "2026-04-01": ["3481群創", "6446藥華藥", "2368金像電", "2334華邦電", "3037欣興", "2449京元電", "7769鴻勁"]
}

# 回測期間設定 (設為 None 則使用整段歷史資料，或依 LAST_DATE 為準)
START_DATE = None                     # 回測紀錄起始日期
END_DATE = None                       # 設為 None 時，會自動以 LAST_DATE 或資料最末日為準

# ==============================================================================
# 延續部位設定 (Warm-Start：從 equityV-adj4.xlsx 2025/12/31 持倉延續)
# ==============================================================================
# 說明：本策略採年初損益歸零、部位延續模式。
#       以下數據來自 equityV-adj4.xlsx 的 2025/12/31 最後狀態。
#       若為全新回測 (非延續)，請將 WARM_START_ENABLED 設為 False。
#
# 數據來源：
#   - 現金：Equity_Curve 2025/12/31 總權益(157,948,201) - 當日市值總計(30,022,550)
#   - 部位：Daily Sheet 2025/12/31 最後持倉，entry_price 來自 Trades 最後買進紀錄，
#           max_price 設為 2025/12/31 收盤價 (作為停損追蹤高點的起始基準)。
WARM_START_ENABLED = False

WARM_START_CASH = 30_000_000.0       # 2025/12/31 現金部位 (surplus_pool)

# 格式：{ slot_id: { 'code': 股票代號字串, 'shares': 持有股數,
#                    'entry_price': 買進成本價, 'max_price': 停損追蹤最高價,
#                    'budget': 買進總成本, 'entry_date': 買進訊號日期字串 } }
WARM_START_SLOTS = {
    0: {'code': '3211', 'shares': 29000, 'entry_price': 342.50, 'max_price': 337.00,
        'budget': 9_946_653.81, 'entry_date': '2025-12-29'},
    1: {'code': '3152', 'shares': 66000, 'entry_price': 149.50, 'max_price': 147.00,
        'budget': 9_881_060.48, 'entry_date': '2025-12-29'},
    2: {'code': '3260', 'shares': 39000, 'entry_price': 252.06, 'max_price': 270.45,
        'budget': 9_844_348.23, 'entry_date': '2025-12-29'},
}

# 策略與風險管理參數 (對應需求 1)
INITIAL_TRADING_CAPITAL = 30000000     # 最初投入資金 (Trading Capital)
INITIAL_AUTHORIZED_CAPITAL = 150000000 # 初始授權金額 (Authorized Capital)

# 指標參數
SMA_PERIOD = 303                       # SMA 均線週期
ROC_PERIOD = 14                        # ROC 動能週期
REBALANCE_INTERVAL = 9                 # 再平衡間隔日數

# 停損參數
STOP_LOSS_TYPE = 'fixed'                 # 停損類型: 'vol' (滾動標準差停損) 或 'fixed' (固定比例停損)
STOP_LOSS_VAL = 0.0999                 # 固定停損比例 (當停損類型為 'fixed' 時生效)
VOL_PERIOD = 15                        # 滾動標準差(波動度)計算週期
VOL_MULTIPLIER = 2.7                   # 停損滾動標準差倍數
USE_BREADTH_WEIGHT = False              # 是否啟用市場寬度加權停損 (寬度低於門檻時倍數乘以 0.8)

# 大盤寬度濾網參數
USE_MARKET_FILTER = False               # 是否啟用大盤濾網
BREADTH_THRESHOLD = 0.42               # 市場寬度門檻
BREADTH_WINDOW = 290                   # 計算市場寬度所需的均線窗格
MKT_SMA_WINDOW = 14                    # 大盤均線窗格

# 交易成本與滑價
COMMISSION_RATE = 0.001425             # 買入手續費率
TAX_RATE = 0.003                       # 賣出交易稅率
SL_SLIPPAGE = 0.0                      # 停損滑價
FILTER_SLIPPAGE = 0.0                  # 濾網出場滑價


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
                if stock in prices_filtered.columns:
                    prices_filtered.loc[pre_mask, stock] = np.nan
                if stock in volumes_filtered.columns:
                    volumes_filtered.loc[pre_mask, stock] = np.nan
                    
    return prices_filtered, volumes_filtered


# ==============================================================================
# 3. 回測引擎類別 (BacktesterVol)
# ==============================================================================
class BacktesterVol:
    """
    回測引擎：加入市場寬度與趨勢雙重確認濾網，並使用 滾動標準差 (Volatility) 動態停損。
    支援雙資本指標：最初投入資金 (Trading Capital) 30M, 初始授權金額 (Authorized Capital) 150M。
    """
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
        self.warm_start_slots = warm_start_slots   # 延續部位設定 (dict 或 None)
        self.warm_start_cash = warm_start_cash     # 延續現金部位 (float 或 None)

        # 交易成本
        self.commission_rate = COMMISSION_RATE
        self.tax_rate = TAX_RATE

    def run(self, sma_period, roc_period, stop_loss_type='vol', stop_loss_val=0.0999,
            vol_period=15, vol_multiplier=2.7,
            rebalance_interval=9, use_market_filter=True,
            breadth_threshold=0.42, mkt_sma_window=14, breadth_window=290,
            start_date=None, end_date=None, use_breadth_weight=True,
            sl_slippage=0.0, filter_slippage=0.0):

        # 1. 指標預計算
        sma = self.prices_df.rolling(window=sma_period).mean().values
        roc = self.prices_df.pct_change(periods=roc_period).values
        sma5 = self.prices_df.rolling(window=5).mean().values
        sma10 = self.prices_df.rolling(window=10).mean().values
        sma20 = self.prices_df.rolling(window=20).mean().values

        # Volatility 計算 (滾動標準差)
        returns = self.prices_df.pct_change()
        vol = returns.rolling(window=vol_period).std().values

        # 市場寬度濾網：使用 np.nanmean 與 np.where 來排除 NaN 標的之影響
        breadth_sma_all = self.prices_df.rolling(window=breadth_window).mean().values
        breadth = np.nanmean(np.where(np.isnan(self.prices), np.nan, self.prices > breadth_sma_all), axis=1)

        market_avg = self.prices_df.mean(axis=1).values
        market_sma = self.prices_df.mean(axis=1).rolling(window=mkt_sma_window).mean().values

        mkt_filter = (breadth >= breadth_threshold) | (market_avg >= market_sma)

        # 2. 確定區間索引
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

        # 考慮指標計算所需的緩衝期
        buffer = max(sma_period, roc_period, breadth_window, mkt_sma_window, vol_period + 1)
        loop_start = max(first_idx, buffer)

        # 3. 帳戶與槽位初始化 (支援 Warm-Start 延續部位)
        if self.warm_start_cash is not None:
            surplus_pool = float(self.warm_start_cash)
        else:
            surplus_pool = float(self.trading_capital)

        # 初始化 slots，若有 warm_start_slots 則從延續部位建立初始持倉
        slots = {0: None, 1: None, 2: None}
        if self.warm_start_slots:
            # 將代碼字串轉換為 asset index，並建立 slot 資料結構
            asset_code_to_idx = {str(code): idx for idx, code in enumerate(self.assets)}
            for s_id, ws in self.warm_start_slots.items():
                code_str = str(ws['code'])
                if code_str in asset_code_to_idx:
                    a_idx = asset_code_to_idx[code_str]
                    entry_reason = f"延續部位(來自 equityV-adj4)，entry={ws['entry_price']}"
                    slots[s_id] = {
                        'asset_idx': a_idx,
                        'shares': ws['shares'],
                        'max_price': ws['max_price'],
                        'budget': ws['budget'],
                        'entry_date': pd.to_datetime(ws['entry_date']),
                        'entry_price': ws['entry_price'],
                        'entry_reason': entry_reason
                    }

        equity_curve_data = []
        trades_log = []
        trades2_log = []
        daily_details = []
        equity_hold_details = [] # 橫向格式明細：一日一行

        # peak_equity 初始值：若有 warm_start，以 cash + 初始市值估計；否則以 trading_capital
        if self.warm_start_slots and self.warm_start_cash is not None:
            init_mv = sum(ws['shares'] * ws['max_price'] for ws in self.warm_start_slots.values())
            peak_equity = float(self.warm_start_cash) + init_mv
        else:
            peak_equity = float(self.trading_capital)

        # 實戰交易模式：每年初損益歸零，持有部位延續
        # start_of_year_equity 以首個回測日的實際總資產為基準
        start_of_year_equity = None  # 第一天進入迴圈後再初始化
        current_year = self.dates[loop_start].year

        for i in range(loop_start, last_idx + 1):
            date = self.dates[i]
            current_prices = self.prices[i]

            # 年度重置檢查 (start_of_year_equity 第一天初始化，每年度換年時重設)
            if start_of_year_equity is None:
                # 第一個交易日：以帳戶實際初始狀態設定年初基準
                init_ws_mv = sum(ws['shares'] * ws['max_price'] for ws in self.warm_start_slots.values()) if self.warm_start_slots else 0.0
                start_of_year_equity = surplus_pool + init_ws_mv
            elif date.year != current_year:
                start_of_year_equity = total_equity if 'total_equity' in locals() else surplus_pool
                current_year = date.year

            # 記錄今天開始時持有的狀態 (以防隨後在當天決定賣出)
            # 因為 slots 會在當天被賣出邏輯更新，所以我們得先拷貝今天開盤時持有部位的狀態
            initial_slots = {s_id: (info.copy() if info else None) for s_id, info in slots.items()}

            stock_mv = 0.0
            for s_id, info in initial_slots.items():
                if info and 'asset_idx' in info:
                    a_idx = info['asset_idx']
                    mv = info['shares'] * current_prices[a_idx]
                    stock_mv += mv

            total_equity = surplus_pool + stock_mv
            if total_equity > peak_equity: 
                peak_equity = total_equity

            # 標準 MDD (相對於最高點)
            drawdown = (total_equity - peak_equity) / peak_equity if peak_equity != 0 else 0
            # 固定基準 MDD (相對於初始授權資金 150M)
            drawdown_fixed = (total_equity - peak_equity) / self.authorized_capital

            # 年度損益 (從年初 0 開始)
            yearly_pnl = total_equity - start_of_year_equity
            yearly_return = yearly_pnl / start_of_year_equity if start_of_year_equity != 0 else 0

            equity_curve_data.append({
                '日期': date,
                '權益': total_equity,
                '回撤(Drawdown)': drawdown,
                '固定基準回撤': drawdown_fixed,
                '市場寬度': breadth[i],
                '年度損益': yearly_pnl,
                '年度報酬率': yearly_return
            })

            # 初始化次一日交易動作收集器 (用於橫向備忘欄位)
            buys_list = []
            sells_list = []
            holds_list = []

            # 暫存次一日個別 slot 的交易動作 (用於 Daily 縱向明細)
            slot_actions = {s_id: "續抱" for s_id in slots}

            # 1. 大盤/市場濾網檢查
            market_filter_triggered = use_market_filter and not mkt_filter[i]

            if market_filter_triggered:
                # 所有持有部位次一日均為賣出
                for s_id, info in initial_slots.items():
                    if info and 'asset_idx' in info:
                        a_idx = info['asset_idx']
                        reason = f"市場濾網：寬度({breadth[i]:.1%})與大盤皆弱"
                        sells_list.append(f"{self.assets[a_idx]}{self.code_to_name[self.assets[a_idx]]}")
                        slot_actions[s_id] = f"賣出 ({reason})"
                
                # 如果不是最後一天，則實際執行交易
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
                    # 最後一天只模擬，不實際更新 slots
                    pass

            else:
                # 2. 每日檢查停損
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
                            effective_mult = vol_multiplier
                            if use_breadth_weight and breadth[i] < breadth_threshold:
                                effective_mult = vol_multiplier * 0.8

                            stop_price = info['max_price'] * (1 - effective_mult * current_vol)
                            if current_prices[a_idx] < stop_price:
                                exit_triggered = True
                                reason_str = f"Vol停損：低於停損價{stop_price:.2f}"

                        if exit_triggered:
                            exited_slots.add(s_id)
                            sells_list.append(f"{self.assets[a_idx]}{self.code_to_name[self.assets[a_idx]]}(停損)")
                            slot_actions[s_id] = f"賣出 ({reason_str})"
                            
                            # 如果不是最後一天，則實際執行交易
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

                # 3. 再平衡
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
                    # 對於此時 slots 中非 None (即沒被停損) 的部位進行排名檢查
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

                                # 如果不是最後一天，則實際執行交易
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

                    # 買進新標的
                    # 我們使用拷貝的 slots 來進行買進動作模擬，以防在最後一天更新 slots 狀態
                    temp_slots = slots.copy()
                    for s_id in temp_slots:
                        if temp_slots[s_id] is None and top_3_signals:
                            next_sig = None
                            for sig in top_3_signals:
                                if sig not in [temp_slots[sid]['asset_idx'] for sid in temp_slots if temp_slots[sid]]:
                                    next_sig = sig
                                    break
                            if next_sig is not None:
                                buys_list.append(f"{self.assets[next_sig]}{self.code_to_name[self.assets[next_sig]]}")
                                
                                # 如果不是最後一天，則實際執行買進並更新 slots
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
                                        temp_slots[s_id] = slots[s_id] # 同時更新 temp_slots 保持同步！
                                        trades_log.append({
                                            '訊號日期': date, '股票代號': self.assets[next_sig], '狀態': '買進',
                                            '價格': buy_price_exec, '股數': shares, '動能值': f"{roc[i][next_sig]*100:.2f}%",
                                            '標的名稱': self.code_to_name[self.assets[next_sig]], '原因': '符合趨勢',
                                            '買入手續費': buy_fee, '賣出手續費': 0, '賣出交易稅': 0,
                                            '說明': f"買進：{self.code_to_name[self.assets[next_sig]]}"
                                        })
                                else:
                                    # 最後一天只模擬，不實際更新 slots
                                    temp_slots[s_id] = {'asset_idx': next_sig}

                    # 保持/續抱 Log (僅在非最後一天寫入 Trades sheet)
                    if i < last_idx:
                        for sig, s_id in signal_to_slot_map.items():
                            trades_log.append({
                                '訊號日期': date, '股票代號': self.assets[sig], '狀態': '保持',
                                '價格': current_prices[sig], '股數': slots[s_id]['shares'],
                                '動能值': f"{roc[i][sig]*100:.2f}%", '標的名稱': self.code_to_name[self.assets[sig]], '原因': '趨勢持續',
                                '買入手續費': 0, '賣出手續費': 0, '賣出交易稅': 0, '說明': f"續抱：{self.code_to_name[self.assets[sig]]}"
                            })
                else:
                    # 非再平衡日，其餘部位皆為續抱
                    for s_id, info in slots.items():
                        if info and 'asset_idx' in info and s_id not in exited_slots:
                            holds_list.append(f"{self.assets[info['asset_idx']]}{self.code_to_name[self.assets[info['asset_idx']]]}")
                            slot_actions[s_id] = "續抱"

            # 4. 彙整「次一日交易動作」綜合備忘字串
            action_parts = []
            if buys_list:
                action_parts.append(f"【買進】" + "、".join(buys_list))
            if sells_list:
                action_parts.append(f"【賣出】" + "、".join(sells_list))
            if holds_list:
                action_parts.append(f"【續抱】" + "、".join(holds_list))
            
            action_memo = " | ".join(action_parts) if action_parts else "無動作"

            # 5. 寫入 Daily (縱向多行) 與 Equity_Hold (橫向一行) 數據
            # A. 縱向明細寫入 daily_details (原 Daily sheet 數據來源)
            for s_id, info in initial_slots.items():
                if info and 'asset_idx' in info:
                    a_idx = info['asset_idx']
                    mv = info['shares'] * current_prices[a_idx]
                    daily_details.append({
                        '日期': date,
                        '股票代號': self.assets[a_idx],
                        '股票名稱': self.code_to_name[self.assets[a_idx]],
                        '持有股數': info['shares'],
                        '本日收盤價': current_prices[a_idx],
                        '市值': mv,
                        '次一日交易動作': slot_actions[s_id]
                    })

            # B. 橫向明細寫入 equity_hold_details (全新 Equity_Hold sheet 數據來源)
            holdings_str = []
            for s_id in range(3):
                info = initial_slots.get(s_id)
                if info and 'asset_idx' in info:
                    a_idx = info['asset_idx']
                    holdings_str.append(f"{self.assets[a_idx]}{self.code_to_name[self.assets[a_idx]]} ({info['shares']:,}股)")
                else:
                    holdings_str.append("無")

            equity_hold_details.append({
                '日期': date,
                '持有部位 1': holdings_str[0],
                '持有部位 2': holdings_str[1],
                '持有部位 3': holdings_str[2],
                '次一日交易動作': action_memo
            })

        return (pd.DataFrame(equity_curve_data), 
                pd.DataFrame(trades_log), 
                pd.DataFrame(trades2_log), 
                pd.DataFrame(daily_details), 
                pd.DataFrame(equity_hold_details))


# ==============================================================================
# 4. 績效指標計算函數
# ==============================================================================
def calculate_metrics_dual(equity_curve_df, trading_cap, authorized_cap):
    if equity_curve_df.empty: return {}
    equity = equity_curve_df['權益']
    total_gain = equity.iloc[-1] - equity.iloc[0]
    days = (equity_curve_df['日期'].iloc[-1] - equity_curve_df['日期'].iloc[0]).days
    years = days / 365.25
    
    trading_total_return = (equity.iloc[-1] / trading_cap) - 1
    trading_cagr = (1 + trading_total_return) ** (1 / years) - 1 if years > 0 else 0
    
    authorized_total_return = total_gain / authorized_cap
    authorized_cagr = (1 + authorized_total_return) ** (1 / years) - 1 if years > 0 else 0
    
    max_dd = equity_curve_df['回撤(Drawdown)'].min()
    max_fixed_dd = equity_curve_df['固定基準回撤'].min()
    
    trading_calmar = trading_cagr / abs(max_dd) if max_dd != 0 else 0
    yearly_perf = equity_curve_df.groupby(equity_curve_df['日期'].dt.year).last()[['年度報酬率', '年度損益']]
    
    return {
        'Trading CAGR': trading_cagr, 
        'Authorized CAGR': authorized_cagr, 
        'Standard MaxDD': max_dd,
        'Fixed Base MaxDD': max_fixed_dd, 
        'Trading Calmar': trading_calmar, 
        'Yearly Performance': yearly_perf
    }


# ==============================================================================
# 5. 高品質 Excel 報表匯出函數 (美化與自動排版)
# ==============================================================================
def export_to_excel_premium(equity_df, trades_df, trades2_df, daily_df, hold_df, metrics, filename):
    print(f"正在產出高品質 Excel 報表：{filename}...")
    
    # 複製 DataFrame 並將日期轉換成 YYYY/MM/DD 字串，僅供 Excel 輸出使用
    eq_out = equity_df.copy()
    t_out = trades_df.copy()
    t2_out = trades2_df.copy()
    d_out = daily_df.copy()
    h_out = hold_df.copy()

    if not eq_out.empty:
        eq_out['日期'] = pd.to_datetime(eq_out['日期']).dt.strftime('%Y/%m/%d')
    if not t_out.empty:
        t_out['訊號日期'] = pd.to_datetime(t_out['訊號日期']).dt.strftime('%Y/%m/%d')
    if not t2_out.empty:
        t2_out['買進訊號日期'] = pd.to_datetime(t2_out['買進訊號日期']).dt.strftime('%Y/%m/%d')
        t2_out['賣出訊號日期'] = pd.to_datetime(t2_out['賣出訊號日期']).dt.strftime('%Y/%m/%d')
    if not d_out.empty:
        d_out['日期'] = pd.to_datetime(d_out['日期']).dt.strftime('%Y/%m/%d')
    if not h_out.empty:
        h_out['日期'] = pd.to_datetime(h_out['日期']).dt.strftime('%Y/%m/%d')

    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        # 寫入各個 Sheet
        t_out.to_excel(writer, sheet_name='Trades', index=False)
        t2_out.to_excel(writer, sheet_name='Trades2', index=False)
        eq_out.to_excel(writer, sheet_name='Equity_Curve', index=False)
        h_out.to_excel(writer, sheet_name='Equity_Hold', index=False)
        d_out.to_excel(writer, sheet_name='Daily', index=False)

        # 彙整 Summary 內容
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

        # 取得 xlsxwriter 的 workbook 與 worksheet
        workbook = writer.book
        font_name = 'Microsoft JhengHei' # 微軟正黑體

        # 定義格式物件
        header_fmt = workbook.add_format({
            'bold': True,
            'font_name': font_name,
            'font_color': '#FFFFFF',
            'bg_color': '#1B365D', # 經典深藍色
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        num_fmt = workbook.add_format({'num_format': '#,##0', 'font_name': font_name})
        pct_fmt = workbook.add_format({'num_format': '0.00%', 'font_name': font_name})
        text_fmt = workbook.add_format({'font_name': font_name})
        
        summary_title_fmt = workbook.add_format({
            'bold': True,
            'font_name': font_name,
            'font_size': 12,
            'bg_color': '#D9E1F2', # 淺藍灰色
            'border': 1,
            'align': 'left'
        })

        summary_section_fmt = workbook.add_format({
            'bold': True,
            'font_name': font_name,
            'font_size': 11,
            'bg_color': '#F2F2F2',
            'border': 1,
            'align': 'left'
        })

        # 格式化 Summary Sheet
        summary_sheet = writer.sheets['Summary']
        summary_sheet.set_column('A:A', 35)
        summary_sheet.set_column('B:D', 20)
        
        summary_sheet.write('A1', '策略指標 (全期間)', summary_title_fmt)
        summary_sheet.write('B1', '數值', summary_title_fmt)
        
        summary_sheet.write('A2', '最初投入資金 (Trading Capital)', text_fmt)
        summary_sheet.write('B2', INITIAL_TRADING_CAPITAL, num_fmt)
        summary_sheet.write('A3', '初始授權金額 (Authorized Capital)', text_fmt)
        summary_sheet.write('B3', INITIAL_AUTHORIZED_CAPITAL, num_fmt)
        
        summary_sheet.write('A4', 'Trading CAGR (30M)', text_fmt)
        summary_sheet.write('B4', metrics['Trading CAGR'], pct_fmt)
        summary_sheet.write('A5', 'Authorized CAGR (150M)', text_fmt)
        summary_sheet.write('B5', metrics['Authorized CAGR'], pct_fmt)
        
        summary_sheet.write('A6', 'Standard MaxDD', text_fmt)
        summary_sheet.write('B6', metrics['Standard MaxDD'], pct_fmt)
        summary_sheet.write('A7', 'Fixed Base MaxDD (150M)', text_fmt)
        summary_sheet.write('B7', metrics['Fixed Base MaxDD'], pct_fmt)
        
        summary_sheet.write('A8', 'Trading Calmar', text_fmt)
        summary_sheet.write('B8', metrics['Trading Calmar'], workbook.add_format({'num_format': '0.00', 'font_name': font_name}))
        
        summary_sheet.write('A10', '年度績效 (實戰模式)', summary_section_fmt)
        summary_sheet.write('B10', '年度報酬率', summary_section_fmt)
        summary_sheet.write('C10', '年度損益 (TWD)', summary_section_fmt)
        summary_sheet.write('D10', '年度MDD (150M基準)', summary_section_fmt)
        
        row_idx = 10
        for year, row in metrics['Yearly Performance'].iterrows():
            year_mdd_fixed = equity_raw[equity_raw['Year'] == year]['固定基準回撤'].min()
            summary_sheet.write(row_idx, 0, f"{int(year)} 年度", text_fmt)
            summary_sheet.write(row_idx, 1, row['年度報酬率'], pct_fmt)
            summary_sheet.write(row_idx, 2, row['年度損益'], num_fmt)
            summary_sheet.write(row_idx, 3, year_mdd_fixed, pct_fmt)
            row_idx += 1

        # 格式化所有數據分頁
        sheet_mapping = {
            'Trades': t_out,
            'Trades2': t2_out,
            'Equity_Curve': eq_out,
            'Equity_Hold': h_out,
            'Daily': d_out
        }

        for sheet_name, df in sheet_mapping.items():
            sheet = writer.sheets[sheet_name]
            
            # 寫入表頭樣式
            for col_num, col_name in enumerate(df.columns):
                sheet.write(0, col_num, col_name, header_fmt)
                
            # 自動調整欄寬
            for col_num, col_name in enumerate(df.columns):
                max_len = max(
                    df[col_name].astype(str).map(len).max(),
                    len(col_name)
                ) + 4
                max_len = min(max(max_len, 10), 65) # 限制寬度範圍 10 ~ 65 字元
                
                # 依據資料型態配置格式
                if any(x in col_name for x in ['報酬率', '回撤', '寬度', '動能值']):
                    sheet.set_column(col_num, col_num, max_len, pct_fmt)
                elif any(x in col_name for x in ['權益', '損益', '價格', '收盤價', '市值', '股數', '手續費', '交易稅']):
                    sheet.set_column(col_num, col_num, max_len, num_fmt)
                else:
                    sheet.set_column(col_num, col_num, max_len, text_fmt)

        # 在 Equity_Curve Sheet 插入權益曲線圖表
        curve_sheet = writer.sheets['Equity_Curve']
        chart = workbook.add_chart({'type': 'line'})
        max_row = len(equity_df)
        chart.add_series({
            'name': '最初投入資金權益 (Trading Capital 30M)',
            'categories': ['Equity_Curve', 1, 0, max_row, 0],
            'values':     ['Equity_Curve', 1, 1, max_row, 1],
            'line':       {'color': '#1B365D'}
        })
        chart.set_title({'name': '權益增長趨勢曲線', 'name_font': {'name': font_name, 'size': 14, 'bold': True}})
        chart.set_x_axis({'name': '日期', 'name_font': {'name': font_name, 'size': 10}})
        chart.set_y_axis({'name': '權益 (TWD)', 'name_font': {'name': font_name, 'size': 10}})
        chart.set_legend({'position': 'bottom'})
        chart.set_size({'width': 760, 'height': 420})
        curve_sheet.insert_chart('H2', chart)


# ==============================================================================
# 6. 主執行程序 (main)
# ==============================================================================
def main():
    print(f"==================================================")
    print(f"交易工程師策略引擎啟動中...")
    print(f"==================================================")
    
    # 檢查原始資料檔是否存在
    if not os.path.exists(DATA_FILE):
        raise FileNotFoundError(f"找不到原始資料檔: {DATA_FILE}，請確認名稱是否正確！")

    print(f"[步驟 1/4] 讀取資料來源: {DATA_FILE}...")
    prices_raw, volumes_raw, code_to_name = clean_data(DATA_FILE)

    print(f"[步驟 2/4] 應用每季標的池註冊表進行排除過濾...")
    # 將每季新增標的生效日前設為 NaN 排除
    prices_filtered, volumes_filtered = apply_new_stocks_registry(prices_raw, volumes_raw, NEW_STOCKS_REGISTRY)

    # 限制最新更新日期至 LAST_DATE
    if LAST_DATE:
        last_date_dt = pd.to_datetime(LAST_DATE)
        prices_filtered = prices_filtered.loc[prices_filtered.index <= last_date_dt]
        volumes_filtered = volumes_filtered.loc[volumes_filtered.index <= last_date_dt]
        print(f"  [OK] 已成功將資料截止日限制於最新日期: {LAST_DATE}")

    # 執行回測
    print(f"[步驟 3/4] 執行回測引擎中...")
    print(f"  策略設定: SMA={SMA_PERIOD}, ROC={ROC_PERIOD}, 再平衡={REBALANCE_INTERVAL}天")
    print(f"  停損設定: 類型={STOP_LOSS_TYPE}, 倍數={VOL_MULTIPLIER} (滾動標準差週期={VOL_PERIOD})")
    
    bt = BacktesterVol(
        prices_filtered, 
        volumes_filtered, 
        code_to_name, 
        trading_capital=INITIAL_TRADING_CAPITAL, 
        authorized_capital=INITIAL_AUTHORIZED_CAPITAL
    )

    eq, t, t2, d, h = bt.run(
        sma_period=SMA_PERIOD,
        roc_period=ROC_PERIOD,
        stop_loss_type=STOP_LOSS_TYPE,
        stop_loss_val=STOP_LOSS_VAL,
        vol_period=VOL_PERIOD,
        vol_multiplier=VOL_MULTIPLIER,
        rebalance_interval=REBALANCE_INTERVAL,
        use_market_filter=USE_MARKET_FILTER,
        breadth_threshold=BREADTH_THRESHOLD,
        mkt_sma_window=MKT_SMA_WINDOW,
        breadth_window=BREADTH_WINDOW,
        start_date=START_DATE,
        end_date=END_DATE,
        use_breadth_weight=USE_BREADTH_WEIGHT,
        sl_slippage=SL_SLIPPAGE,
        filter_slippage=FILTER_SLIPPAGE
    )

    # 計算績效指標 (在計算前保持 datetime 格式，計算後再作展示)
    metrics = calculate_metrics_dual(eq, INITIAL_TRADING_CAPITAL, INITIAL_AUTHORIZED_CAPITAL)

    # 動態產生帶有最新日期後綴的 Excel 檔名
    date_suffix = pd.to_datetime(LAST_DATE).strftime('%y%m%d')
    output_filename = f"trendstrategy_results_equityV-adj4-1-{date_suffix}.xlsx"

    print(f"[步驟 4/4] 匯出報表與美化...")
    export_to_excel_premium(eq, t, t2, d, h, metrics, output_filename)

    print(f"==================================================")
    print(f"[OK] 執行完成！")
    print(f"  * 美化版 Excel 報表: {output_filename}")
    print(f"  * 歷史總報酬 CAGR (30M): {metrics['Trading CAGR']:.2%}")
    print(f"  * 標準最大回撤 MaxDD: {metrics['Standard MaxDD']:.2%}")
    print(f"==================================================")


if __name__ == "__main__":
    main()
