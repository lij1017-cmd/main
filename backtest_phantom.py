import pandas as pd
import numpy as np
import nbformat as nbf
import xlsxwriter
import os

# ==========================================
# 1. 核心回測引擎 (具備完整日誌紀錄與交易成本計算)
# ==========================================
def clean_data(filepath):
    """
    清洗並預處理輸入的 Excel 資料檔。
    讀取「還原收盤價」與「成交量」工作表。
    """
    df_prices = pd.read_excel(filepath, sheet_name='還原收盤價', header=None)
    df_volume = pd.read_excel(filepath, sheet_name='成交量', header=None)

    stock_codes = df_prices.iloc[0, 1:].values
    stock_names = df_prices.iloc[1, 1:].values

    # 處理日期格式 (例如 20190102收盤價)
    date_strings = df_prices.iloc[2:, 0].astype(str).str[:8]
    dates = pd.to_datetime(date_strings, format='%Y%m%d')

    # 提取價格與成交量矩陣
    prices = df_prices.iloc[2:, 1:].astype(float)
    prices.index = dates
    prices.columns = stock_codes

    volumes = df_volume.iloc[2:, 1:].astype(float)
    volumes.index = dates
    volumes.columns = stock_codes

    code_to_name = dict(zip(stock_codes, stock_names))

    # 缺失值處理：價格前向/後向填充，成交量補 0
    prices = prices.ffill().bfill()
    volumes = volumes.fillna(0)

    return prices, volumes, code_to_name

def calculate_indicators(prices_df):
    """
    計算策略所需之各項指標：
    - EMA 200: 長期趨勢過濾
    - MACD (12, 26, 9): 進出場動能訊號
    - Rolling Std (20d): 作為動態容忍度 (替代 ATR)
    - Rolling Max/Min (10d): 作為支撐與壓力代理 (替代 Pivot)
    """
    # 趨勢指標
    ema200 = prices_df.ewm(span=200, adjust=False).mean()

    # 動能指標
    exp1 = prices_df.ewm(span=12, adjust=False).mean()
    exp2 = prices_df.ewm(span=26, adjust=False).mean()
    macd_line = exp1 - exp2
    signal_line = macd_line.ewm(span=9, adjust=False).mean()
    macd_hist = macd_line - signal_line

    # 波動度指標 (作為動態容忍度)
    volatility = prices_df.rolling(window=20).std()

    # 支撐壓力代理 (10 日滾動高低點)
    swing_high = prices_df.rolling(window=10).max()
    swing_low = prices_df.rolling(window=10).min()

    return ema200, macd_line, signal_line, macd_hist, volatility, swing_high, swing_low

class PhantomBacktester:
    """
    回測引擎：支援多部位管理 (Slots)、交易成本計算及多/空方策略。
    """
    def __init__(self, prices, volumes, code_to_name, initial_capital=30000000, mode='LS'):
        self.prices_df = prices
        self.volumes_df = volumes
        self.prices = prices.values
        self.dates = prices.index
        self.assets = prices.columns
        self.code_to_name = code_to_name
        self.initial_capital = initial_capital
        self.mode = mode # 'L': 僅做多, 'LS': 多空皆做

        # 策略參數
        self.vol_frac = 0.5   # 波動度容忍係數
        self.rr_target = 1.5  # 風險報酬比

        # 交易成本設定 (台灣市場)
        self.commission_rate = 0.001425 # 手續費
        self.tax_rate = 0.003           # 交易稅
        self.short_fee_rate = 0.001     # 融券手續費

    def run(self):
        # 預計算所有指標
        ema200, macd_line, signal_line, _, vol, s_high, s_low = calculate_indicators(self.prices_df)

        e_v, m_v, s_v = ema200.values, macd_line.values, signal_line.values
        v_v, sh_v, sl_v = vol.values, s_high.values, s_low.values

        # 帳戶狀態
        surplus_pool = float(self.initial_capital)
        slots = {i: None for i in range(10)} # 限制最多 10 個部位

        # 紀錄用清單
        equity_curve = []
        trades_log = []
        trades2_log = []
        holdings_history = []
        daily_details = []

        # 狀態追蹤：紀錄是否發生過「支撐/壓力突破」
        break_low_active = np.zeros(len(self.assets), dtype=bool)
        break_high_active = np.zeros(len(self.assets), dtype=bool)

        start_idx = 200 # 等待指標穩定
        peak_equity = float(self.initial_capital)

        for i in range(start_idx, len(self.dates)):
            date = self.dates[i]
            curr_prices = self.prices[i]

            # A. 更新每日權益與持股明細
            long_mv, short_mv = 0.0, 0.0
            h_names = []
            for s_id, info in slots.items():
                if info:
                    mv = info['shares'] * curr_prices[info['asset_idx']]
                    if info['type'] == 'Long': long_mv += mv
                    else: short_mv += mv
                    name = self.code_to_name[self.assets[info['asset_idx']]]
                    h_names.append(f"{name}({self.assets[info['asset_idx']]})")
                    daily_details.append({
                        '日期': date, '股票代號': self.assets[info['asset_idx']], '股票名稱': name,
                        '持有股數': info['shares'], '本日收盤價': curr_prices[info['asset_idx']],
                        '市值': mv, '類型': info['type']
                    })

            total_equity = surplus_pool + long_mv - short_mv
            if total_equity > peak_equity: peak_equity = total_equity
            drawdown = (total_equity - peak_equity) / peak_equity if peak_equity != 0 else 0

            equity_curve.append({'日期': date, '權益': total_equity, '回撤(Drawdown)': drawdown})
            holdings_history.append({
                '日期': date, '持股明細': ", ".join(h_names), '持股數': len(h_names),
                '現金': surplus_pool, '多單市值': long_mv, '空單市值': short_mv, '總資產': total_equity
            })

            if i == len(self.dates) - 1: break
            next_prices = self.prices[i+1] # T+1 執行價格

            # B. 檢查出場邏輯 (停損與獲利)
            for s_id, info in slots.items():
                if info:
                    a_idx = info['asset_idx']
                    cp = curr_prices[a_idx]
                    exit_triggered = False
                    reason = ""

                    if info['type'] == 'Long':
                        if cp <= info['stop_loss']: exit_triggered, reason = True, "Stop Loss"
                        elif cp >= info['take_profit']: exit_triggered, reason = True, "Take Profit"
                    else:
                        if cp >= info['stop_loss']: exit_triggered, reason = True, "Stop Loss"
                        elif cp <= info['take_profit']: exit_triggered, reason = True, "Take Profit"

                    if exit_triggered:
                        ex_p = next_prices[a_idx]
                        shares = info['shares']
                        name = self.code_to_name[self.assets[a_idx]]
                        if info['type'] == 'Long':
                            fee = shares * ex_p * self.commission_rate
                            tax = shares * ex_p * self.tax_rate
                            proceeds = shares * ex_p - fee - tax
                            surplus_pool += proceeds
                            pnl = proceeds - info['cost']
                            trades_log.append({
                                '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出', '價格': ex_p,
                                '股數': shares, '標的名稱': name, '原因': reason, '買入手續費': 0, '賣出手續費': fee, '賣出交易稅': tax
                            })
                        else:
                            fee = shares * ex_p * self.commission_rate
                            cost_to_cover = shares * ex_p + fee
                            surplus_pool -= cost_to_cover
                            pnl = info['cost'] - cost_to_cover
                            trades_log.append({
                                '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '補回', '價格': ex_p,
                                '股數': shares, '標的名稱': name, '原因': reason, '買入手續費': fee, '賣出手續費': 0, '賣出交易稅': 0
                            })
                        trades2_log.append({
                            '進場日期': info['entry_date'], '股票代號': self.assets[a_idx], '股票名稱': name,
                            '進場價格': info['entry_price'], '股數': shares, '出場日期': date,
                            '出場價格': ex_p, '損益': pnl, '類型': info['type'], '進場原因': '符合趨勢(Reclaim)', '出場原因': reason
                        })
                        slots[s_id] = None

            # C. 進場訊號檢查
            if any(s is None for s in slots.values()):
                for a in range(len(self.assets)):
                    cp = curr_prices[a]
                    support, resistance = sl_v[i-1, a], sh_v[i-1, a]
                    tol = v_v[i, a] * self.vol_frac

                    # 更新突破狀態
                    if cp < support - tol: break_low_active[a] = True
                    if cp > resistance + tol: break_high_active[a] = True

                    if any(info and info['asset_idx'] == a for info in slots.values()): continue

                    m_cross_up = m_v[i, a] > s_v[i, a] and m_v[i-1, a] <= s_v[i-1, a]
                    m_cross_down = m_v[i, a] < s_v[i, a] and m_v[i-1, a] >= s_v[i-1, a]

                    # 核心過濾條件 (MACD 與 EMA 200)
                    c_long = m_cross_up and m_v[i, a] < 0 and cp > e_v[i, a]
                    c_short = m_cross_down and m_v[i, a] > 0 and cp < e_v[i, a]

                    if c_long:
                        if break_low_active[a] and cp > support:
                            sl = support - tol
                            if sl < cp:
                                surplus_pool = self.enter_trade(slots, surplus_pool, a, i, 'Long', next_prices[a], sl, trades_log)
                                break_low_active[a] = False
                    elif c_short and self.mode == 'LS':
                        if break_high_active[a] and cp < resistance:
                            sl = resistance + tol
                            if sl > cp:
                                surplus_pool = self.enter_trade(slots, surplus_pool, a, i, 'Short', next_prices[a], sl, trades_log)
                                break_high_active[a] = False

        return (pd.DataFrame(equity_curve), pd.DataFrame(trades_log), pd.DataFrame(holdings_history),
                pd.DataFrame(trades2_log), pd.DataFrame(daily_details))

    def enter_trade(self, slots, surplus_pool, a_idx, i, t_type, next_p, stop_loss, trades_log):
        for s_id, info in slots.items():
            if info is None:
                budget = 3000000.0
                name = self.code_to_name[self.assets[a_idx]]
                if t_type == 'Long':
                    if surplus_pool < 1000000: return surplus_pool
                    invest = min(surplus_pool, budget)
                    shares = (invest // (next_p * (1 + self.commission_rate)) // 1000) * 1000
                    if shares <= 0: return surplus_pool
                    cost = shares * next_p * (1 + self.commission_rate)
                    surplus_pool -= cost
                    trades_log.append({
                        '訊號日期': self.dates[i], '股票代號': self.assets[a_idx], '狀態': '買進', '價格': next_p,
                        '股數': shares, '標的名稱': name, '原因': '符合趨勢(Reclaim)',
                        '買入手續費': shares * next_p * self.commission_rate, '賣出手續費': 0, '賣出交易稅': 0
                    })
                else:
                    shares = (budget // (next_p * (1 + self.commission_rate)) // 1000) * 1000
                    if shares <= 0: return surplus_pool
                    fee = shares * next_p * self.commission_rate
                    tax = shares * next_p * self.tax_rate
                    s_fee = shares * next_p * self.short_fee_rate
                    cost = shares * next_p - fee - tax - s_fee
                    surplus_pool += cost
                    trades_log.append({
                        '訊號日期': self.dates[i], '股票代號': self.assets[a_idx], '狀態': '放空', '價格': next_p,
                        '股數': shares, '標的名稱': name, '原因': '符合趨勢(Reclaim)',
                        '買入手續費': 0, '賣出手續費': fee, '賣出交易稅': tax
                    })

                risk = abs(next_p - stop_loss)
                tp = next_p + self.rr_target * risk if t_type == 'Long' else next_p - self.rr_target * risk
                slots[s_id] = {
                    'asset_idx': a_idx, 'shares': shares, 'entry_price': next_p, 'entry_date': self.dates[i],
                    'stop_loss': stop_loss, 'take_profit': tp, 'type': t_type, 'cost': cost
                }
                return surplus_pool
        return surplus_pool

# ==========================================
# 2. 產出成果檔案工具
# ==========================================
def generate_xlsx(filename, eq_df, trades, hold, trades2, daily, summary_df):
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        eq_df.to_excel(writer, sheet_name='Equity_Curve', index=False)
        hold.to_excel(writer, sheet_name='Equity_Hold', index=False)
        trades.to_excel(writer, sheet_name='Trades', index=False)
        trades2.to_excel(writer, sheet_name='Trades2', index=False)
        daily.to_excel(writer, sheet_name='Daily', index=False)

        workbook = writer.book
        curves_sheet = writer.sheets['Equity_Curve']
        chart = workbook.add_chart({'type': 'line'})
        max_row = len(eq_df)
        chart.add_series({
            'name': 'Equity Curve',
            'categories': ['Equity_Curve', 1, 0, max_row, 0],
            'values': ['Equity_Curve', 1, 1, max_row, 1],
        })
        chart.set_title({'name': 'Equity Curve'})
        curves_sheet.insert_chart('E2', chart)

def generate_md(filename, title, summary_dict):
    content = f"""# {title} 策略說明文件 (繁體中文)

## 1. 策略核心概念
本策略改編自 TradingView 的 "Phantom MACD + EMA + Pivot S/R" 策略。
由於原策略依賴的高低價資料受限，本版本採用以下代理指標：

### 指標說明
- **EMA 200**: 用於判斷長期趨勢的方向過濾器。
- **MACD (12, 26, 9)**: 捕捉價格動能的交叉訊號。
- **支撐/壓力代理**: 使用 **10 日滾動收盤最低/最高價**。
- **動態容忍度**: 使用 **20 日收盤價標準差** 代替 ATR。

### 交易邏輯 (Reclaim Mechanism)
- **多方進場**: 價格需先跌破「支撐 - 容忍度」，隨後重新收復支撐線，並搭配 MACD 於零軸下金叉且價格高於 EMA 200。
- **空方進場**: 價格需先突破「壓力 + 容忍度」，隨後重新跌破壓力線，並搭配 MACD 於零軸上死叉且價格低於 EMA 200。

## 2. 交易成本與部位管理
- **最大持股數**: 10 個槽位。
- **單筆投入上限**: 300 萬 TWD。
- **手續費**: 0.1425% (進出皆計)。
- **交易稅**: 0.3% (賣出時計)。
- **融券費**: 0.1% (僅放空進場時計)。

## 3. 績效回測總結 (資料-1.xlsx)
- **年化報酬率 (CAGR)**: {summary_dict['CAGR']}
- **最大回撤 (MaxDD)**: {summary_dict['MaxDD']}
- **卡瑪比率 (Calmar Ratio)**: {summary_dict['Calmar']}
- **完成交易次數**: {summary_dict['Trades']}
- **勝率**: {summary_dict['WinRate']}
"""
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(content)

def generate_ipynb(filename, title):
    nb = nbf.v4.new_notebook()

    # 讀取當前檔案獲取核心代碼
    with open(__file__, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        core_logic = "".join(lines[:330]) # Capture up to the end of functions

    nb.cells = [
        nbf.v4.new_markdown_cell(f"# {title} 策略回測程式碼\n本筆記本包含完整的策略邏輯、指標計算及回測產出。"),
        nbf.v4.new_markdown_cell("## 1. 匯入分析工具\n使用 pandas 處理資料，numpy 進行數值運算。"),
        nbf.v4.new_code_cell("import pandas as pd\nimport numpy as np\nimport matplotlib.pyplot as plt"),
        nbf.v4.new_markdown_cell("## 2. 資料載入與指標計算\n包含 EMA 200, MACD, 10日滾動高低點(支撐壓力), 以及 20日標準差(波動度)。"),
        nbf.v4.new_code_cell(core_logic),
        nbf.v4.new_markdown_cell("## 3. 執行回測\n設定初始資金 3000 萬，並根據策略模式(做多或多空)執行。"),
        nbf.v4.new_code_cell(f"prices, volumes, code_to_name = clean_data('資料-1.xlsx')\nbt = PhantomBacktester(prices, volumes, code_to_name, mode='{'L' if 'LONG' in title else 'LS'}')\neq, tr, hold, tr2, daily = bt.run()"),
        nbf.v4.new_markdown_cell("## 4. 呈現權益曲線"),
        nbf.v4.new_code_cell("plt.figure(figsize=(12, 6))\nplt.plot(eq['日期'], eq['權益'])\nplt.title('Equity Curve')\nplt.show()")
    ]

    with open(filename, 'w', encoding='utf-8') as f:
        nbf.write(nb, f)

if __name__ == "__main__":
    prices, volumes, code_to_name = clean_data('資料-1.xlsx')

    for mode, prefix in [('L', 'MACD-LONG'), ('LS', 'MACD-LS')]:
        print(f"處理中: {prefix}...")
        bt = PhantomBacktester(prices, volumes, code_to_name, mode=mode)
        eq, trades, hold, trades2, daily = bt.run()

        # 指標計算
        def get_metrics(df):
            if df.empty: return 0, 0, 0, 0
            ret = (df['權益'].iloc[-1] / df['權益'].iloc[0]) - 1
            years = (df['日期'].iloc[-1] - df['日期'].iloc[0]).days / 365.25
            cagr = (1 + ret) ** (1/years) - 1 if years > 0 else 0
            mdd = df['回撤(Drawdown)'].min()
            calmar = cagr / abs(mdd) if mdd != 0 else 0
            return cagr, mdd, calmar, ret

        cagr, mdd, calmar, tot = get_metrics(eq)
        wr = (trades2['損益'] > 0).mean() if not trades2.empty else 0

        summary_df = pd.DataFrame([
            {'項目': '年化報酬率 (CAGR)', '數值': f"{cagr:.2%}"},
            {'項目': '最大回撤 (MaxDD)', '數值': f"{mdd:.2%}"},
            {'項目': '卡瑪比率 (Calmar Ratio)', '數值': f"{calmar:.2f}"},
            {'項目': '交易次數', '數值': len(trades2)},
            {'項目': '勝率', '數值': f"{wr:.2%}"},
            {'項目': '總報酬率', '數值': f"{tot:.2%}"},
            {'項目': '初始資金', '數值': "30,000,000"}
        ])

        generate_xlsx(f"{prefix}.xlsx", eq, trades, hold, trades2, daily, summary_df)
        generate_md(f"{prefix}.md", prefix, {
            'CAGR': f"{cagr:.2%}", 'MaxDD': f"{mdd:.2%}", 'Calmar': f"{calmar:.2f}",
            'Trades': len(trades2), 'WinRate': f"{wr:.2%}"
        })
        generate_ipynb(f"{prefix}.ipynb", prefix)

    print("完成！所有檔案 (XLSX, MD, IPYNB) 已根據 LONG 及 LS 分別產出。")
