import pandas as pd
import numpy as np
import warnings
import os

# 忽略 Pandas 的 Slice 複製警告以維持輸出乾淨
warnings.filterwarnings('ignore')

# 預設參數
INITIAL_TRADING_CAPITAL = 30000000     # 策略配置金額 (Trading Capital)
INITIAL_AUTHORIZED_CAPITAL = 150000000 # 初始授權金額 (Authorized Capital)
COMMISSION_RATE = 0.001425             # 買入手續費率
TAX_RATE = 0.003                       # 賣出交易稅率

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
    根據新標的池註冊表，在生效日期之前排除對應標的。
    """
    prices_filtered = prices.copy()
    volumes_filtered = volumes.copy()

    for date_str, stocks in registry.items():
        eff_date = pd.to_datetime(date_str)
        pre_mask = prices_filtered.index < eff_date
        if pre_mask.any():
            for stock in stocks:
                if stock in prices_filtered.columns:
                    prices_filtered.loc[pre_mask, stock] = np.nan
                if stock in volumes_filtered.columns:
                    volumes_filtered.loc[pre_mask, stock] = np.nan

    return prices_filtered, volumes_filtered

class BacktesterVol:
    """
    回測引擎：支援年度報酬率雙軌計算、市場寬度濾網、波動度停損、以及 Warm-Start 銜接。
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
        self.warm_start_slots = warm_start_slots   # 延續部位
        self.warm_start_cash = warm_start_cash     # 延續現金

        self.commission_rate = COMMISSION_RATE
        self.tax_rate = TAX_RATE

    def run(self, sma_period=303, roc_period=14, stop_loss_type='fixed', stop_loss_val=0.0999,
            vol_period=15, vol_multiplier=2.7,
            rebalance_interval=9, use_market_filter=True,
            breadth_threshold=0.42, mkt_sma_window=14, breadth_window=290,
            start_date=None, end_date=None, use_breadth_weight=True,
            sl_slippage=0.0, filter_slippage=0.0):

        # 指標預計算
        sma = self.prices_df.rolling(window=sma_period).mean().values
        roc = self.prices_df.pct_change(periods=roc_period).values
        sma5 = self.prices_df.rolling(window=5).mean().values
        sma10 = self.prices_df.rolling(window=10).mean().values
        sma20 = self.prices_df.rolling(window=20).mean().values
        returns = self.prices_df.pct_change()
        vol = returns.rolling(window=vol_period).std().values

        prices_above_sma = (self.prices_df > self.prices_df.rolling(window=breadth_window).mean()).values
        breadth = np.nanmean(np.where(np.isnan(self.prices), np.nan, prices_above_sma), axis=1)

        market_avg = self.prices_df.mean(axis=1).values
        market_sma = self.prices_df.mean(axis=1).rolling(window=mkt_sma_window).mean().values
        mkt_filter = (breadth >= breadth_threshold) | (market_avg >= market_sma)

        # 確定索引
        if start_date:
            mask = self.dates >= pd.to_datetime(start_date)
            first_idx = np.where(mask)[0][0] if any(mask) else 0
        else:
            first_idx = 0
        last_idx = np.where(self.dates <= pd.to_datetime(end_date))[0][-1] if end_date else len(self.dates)-1

        buffer = max(sma_period, roc_period, breadth_window, mkt_sma_window, vol_period + 1)
        loop_start = first_idx if self.warm_start_slots is not None else max(first_idx, buffer)

        # 帳戶初始化
        surplus_pool = float(self.warm_start_cash) if self.warm_start_cash is not None else float(self.trading_capital)
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

        equity_curve_data, trades_log, trades2_log, daily_details, equity_hold_details = [], [], [], [], []
        peak_equity = surplus_pool + sum(s['shares']*self.prices[loop_start][s['asset_idx']] for s in slots.values() if s)

        # 年度指標重置參考
        start_of_year_total_equity = None
        start_of_year_trading_base = None
        current_year = self.dates[loop_start].year

        for i in range(loop_start, last_idx + 1):
            date, current_prices = self.dates[i], self.prices[i]
            stock_mv = sum(s['shares'] * current_prices[s['asset_idx']] for s in slots.values() if s)
            total_equity = surplus_pool + stock_mv
            peak_equity = max(peak_equity, total_equity)

            if start_of_year_total_equity is None or date.year != current_year:
                start_of_year_total_equity = total_equity
                start_of_year_trading_base = min(total_equity, self.trading_capital)
                current_year = date.year

            yearly_pnl = total_equity - start_of_year_total_equity
            equity_curve_data.append({
                '日期': date, '權益': total_equity,
                '回撤(Drawdown)': (total_equity - peak_equity) / peak_equity if peak_equity != 0 else 0,
                '固定基準回撤': (total_equity - peak_equity) / self.authorized_capital,
                '市場寬度': breadth[i], '年度損益': yearly_pnl,
                '年度報酬率(累積)': yearly_pnl / start_of_year_total_equity if start_of_year_total_equity != 0 else 0,
                '年度報酬率(年初基準)': yearly_pnl / start_of_year_trading_base if start_of_year_trading_base != 0 else 0
            })

            for s_id, info in slots.items():
                if info:
                    a_idx = info['asset_idx']
                    daily_details.append({'日期': date, '股票代號': self.assets[a_idx], '股票名稱': self.code_to_name[self.assets[a_idx]], '持有股數': info['shares'], '買進日期': info['entry_date'], '買進成本': info['entry_price'], '追蹤最高價': info['max_price'], '買入總市值': info['budget'], '本日收盤價': current_prices[a_idx], '市值': info['shares'] * current_prices[a_idx]})

            if i == last_idx: break
            next_prices = self.prices[i+1]

            if use_market_filter and not mkt_filter[i]:
                for s_id, info in slots.items():
                    if info:
                        a_idx = info['asset_idx']
                        sell_price = next_prices[a_idx] * (1 - filter_slippage)
                        proceeds = info['shares'] * sell_price * (1 - self.commission_rate - self.tax_rate)
                        surplus_pool += proceeds
                        trades2_log.append({'買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx], '股票名稱': self.code_to_name[self.assets[a_idx]], 'T+1日買進價格': info['entry_price'], '股數': info['shares'], '賣出訊號日期': date, 'T+1日賣出價格': sell_price, '損益': proceeds - info['budget'], '報酬率': (proceeds / info['budget']) - 1, '買進原因': info['entry_reason'], '賣出原因': "市場濾網"})
                        slots[s_id] = None
                continue

            for s_id, info in slots.items():
                if info:
                    a_idx = info['asset_idx']
                    info['max_price'] = max(info['max_price'], current_prices[a_idx])
                    exit_triggered, reason = False, ""
                    if stop_loss_type == 'fixed':
                        if current_prices[a_idx] < info['max_price'] * (1 - stop_loss_val):
                            exit_triggered, reason = True, f"固定停損 {stop_loss_val:.1%}"
                    elif stop_loss_type == 'vol':
                        mult = vol_multiplier * (0.8 if use_breadth_weight and breadth[i] < breadth_threshold else 1.0)
                        if current_prices[a_idx] < info['max_price'] * (1 - mult * vol[i][a_idx]):
                            exit_triggered, reason = True, "Vol停損"
                    if exit_triggered:
                        sell_price = next_prices[a_idx] * (1 - sl_slippage)
                        proceeds = info['shares'] * sell_price * (1 - self.commission_rate - self.tax_rate)
                        surplus_pool += proceeds
                        trades2_log.append({'買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx], '股票名稱': self.code_to_name[self.assets[a_idx]], 'T+1日買進價格': info['entry_price'], '股數': info['shares'], '賣出訊號日期': date, 'T+1日賣出價格': sell_price, '損益': proceeds - info['budget'], '報酬率': (proceeds / info['budget']) - 1, '買進原因': info['entry_reason'], '賣出原因': reason})
                        slots[s_id] = None

            if (i - loop_start) % rebalance_interval == 0:
                top_3 = []
                sorted_idx = np.argsort(roc[i])[::-1]
                for idx in sorted_idx:
                    if len(top_3) >= 3: break
                    if np.isnan(roc[i][idx]): continue
                    p = current_prices[idx]
                    if p > sma[i][idx] and roc[i][idx] > 0 and (p * self.volumes[i][idx] * 1000) > 30000000 and p > sma5[i][idx] and p > sma10[i][idx] and p > sma20[i][idx]:
                        top_3.append(idx)
                for s_id, info in slots.items():
                    if info and info['asset_idx'] not in top_3:
                        a_idx = info['asset_idx']
                        sell_price = next_prices[a_idx]
                        proceeds = info['shares'] * sell_price * (1 - self.commission_rate - self.tax_rate)
                        surplus_pool += proceeds
                        trades2_log.append({'買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx], '股票名稱': self.code_to_name[self.assets[a_idx]], 'T+1日買進價格': info['entry_price'], '股數': info['shares'], '賣出訊號日期': date, 'T+1日賣出價格': sell_price, '損益': proceeds - info['budget'], '報酬率': (proceeds / info['budget']) - 1, '買進原因': info['entry_reason'], '賣出原因': "再平衡"})
                        slots[s_id] = None
                for s_id in slots:
                    if slots[s_id] is None and top_3:
                        for sig in top_3:
                            if sig not in [s['asset_idx'] for s in slots.values() if s]:
                                buy_p = next_prices[sig]
                                shares = (int(10000000 // (buy_p * (1 + self.commission_rate))) // 1000) * 1000
                                if shares > 0:
                                    cost = shares * buy_p * (1 + self.commission_rate)
                                    surplus_pool -= cost
                                    slots[s_id] = {'asset_idx': sig, 'shares': shares, 'max_price': buy_p, 'budget': cost, 'entry_date': date, 'entry_price': buy_p, 'entry_reason': f"ROC:{roc[i][sig]:.2%}"}
                                    break
            h_row = {'日期': date}
            for s_id in range(3):
                info = slots.get(s_id)
                h_row[f'持有部位 {s_id+1}'] = f"{self.assets[info['asset_idx']]}{self.code_to_name[self.assets[info['asset_idx']]]} ({info['shares']:,}股)" if info else "無"
            equity_hold_details.append(h_row)

        return pd.DataFrame(equity_curve_data), pd.DataFrame(trades_log), pd.DataFrame(trades2_log), pd.DataFrame(daily_details), pd.DataFrame(equity_hold_details)

def calculate_metrics_dual(equity_df, trading_cap, authorized_cap):
    if equity_df.empty: return {}
    eq = equity_df['權益']
    years = (equity_df['日期'].iloc[-1] - equity_df['日期'].iloc[0]).days / 365.25
    t_cagr = (eq.iloc[-1] / trading_cap) ** (1/years) - 1 if years > 0 else 0
    a_cagr = (1 + (eq.iloc[-1] - eq.iloc[0]) / authorized_cap) ** (1/years) - 1 if years > 0 else 0
    yearly_perf = equity_df.groupby(equity_df['日期'].dt.year).last()[['年度報酬率(累積)', '年度報酬率(年初基準)', '年度損益']]
    mdd = equity_df['回撤(Drawdown)'].min()
    return {'Trading CAGR': t_cagr, 'Authorized CAGR': a_cagr, 'MaxDD': mdd, 'Fixed MDD': equity_df['固定基準回撤'].min(), 'Calmar': t_cagr/abs(mdd) if mdd != 0 else 0, 'Yearly Performance': yearly_perf}

def export_to_excel_premium(eq, t, t2, d, h, metrics, filename):
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        t.to_excel(writer, sheet_name='Trades', index=False)
        t2.to_excel(writer, sheet_name='Trades2', index=False)
        eq.to_excel(writer, sheet_name='Equity_Curve', index=False)
        h.to_excel(writer, sheet_name='Equity_Hold', index=False)
        d.to_excel(writer, sheet_name='Daily', index=False)
        summary_data = [['策略指標', '數值'], ['Trading CAGR (30M)', metrics['Trading CAGR']], ['Authorized CAGR (150M)', metrics['Authorized CAGR']], ['Standard MaxDD', metrics['MaxDD']], ['Fixed Base MaxDD (150M)', metrics['Fixed MDD']], ['', ''], ['年度績效', '年度報酬率(累積)', '年度報酬率(年初基準)', '年度損益 (TWD)', '年度MDD (150M基準)']]
        eq_raw = eq.copy()
        eq_raw['Year'] = pd.to_datetime(eq_raw['日期']).dt.year
        for year, row in metrics['Yearly Performance'].iterrows():
            y_mdd = eq_raw[eq_raw['Year'] == year]['固定基準回撤'].min()
            summary_data.append([f"{int(year)} 年度", row['年度報酬率(累積)'], row['年度報酬率(年初基準)'], row['年度損益'], y_mdd])
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False, header=False)
