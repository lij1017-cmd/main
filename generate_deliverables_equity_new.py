import pandas as pd
import numpy as np
import pickle
import nbformat as nbf
import xlsxwriter

# ==========================================
# 1. 核心回測引擎 (基於 run_wfa.py 之邏輯並增強報表)
# ==========================================
def clean_data(filepath):
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

class BacktesterV2:
    def __init__(self, prices, code_to_name, initial_capital=30000000):
        self.prices_df = prices
        self.prices = prices.values
        self.dates = prices.index
        self.assets = prices.columns
        self.code_to_name = code_to_name
        self.initial_capital = initial_capital

    def run(self, sma_period, roc_period, stop_loss_pct, rebalance_interval, start_date, end_date):
        sma = self.prices_df.rolling(window=sma_period).mean().values
        roc = self.prices_df.pct_change(periods=roc_period).values

        mask = (self.dates >= pd.to_datetime(start_date)) & (self.dates <= pd.to_datetime(end_date))
        all_indices = np.where(mask)[0]
        if len(all_indices) == 0: return None, None, None, None, None

        first_idx = all_indices[0]
        last_idx = all_indices[-1]
        start_buffer = max(sma_period, roc_period)
        loop_start = max(first_idx, start_buffer)

        cash = float(self.initial_capital)
        slots = {0: None, 1: None, 2: None} # 槽位化管理

        equity_curve_data = []
        trades_log = []
        trades2_log = []
        holdings_history = []
        daily_details = []

        peak_equity = float(self.initial_capital)
        current_reasons = []

        # 初始權益紀錄
        for i in range(first_idx, loop_start):
            equity_curve_data.append({'日期': self.dates[i], '權益': float(self.initial_capital), '回撤': 0.0})

        for i in range(loop_start, last_idx + 1):
            date = self.dates[i]
            current_prices = self.prices[i]

            # A. 計算今日權益與明細
            stock_mv = 0.0
            h_names = []
            for s_id, info in slots.items():
                if info and 'asset_idx' in info:
                    a_idx = info['asset_idx']
                    mv = info['shares'] * current_prices[a_idx]
                    stock_mv += mv
                    h_names.append(f"{self.code_to_name[self.assets[a_idx]]}({self.assets[a_idx]})")
                    daily_details.append({
                        '日期': date, '股票代號': self.assets[a_idx], '股票名稱': self.code_to_name[self.assets[a_idx]],
                        '持有股數': info['shares'], '本日收盤價': current_prices[a_idx], '市值': mv
                    })

            total_equity = cash + stock_mv
            if total_equity > peak_equity: peak_equity = total_equity
            drawdown = (total_equity - peak_equity) / peak_equity
            equity_curve_data.append({'日期': date, '權益': total_equity, '回撤': drawdown})

            if i == last_idx: break
            next_prices = self.prices[i+1]

            # B. 停損
            triggered_slots = []
            for s_id, info in slots.items():
                if info and 'asset_idx' in info:
                    curr_p = current_prices[info['asset_idx']]
                    if curr_p > info['max_price']: info['max_price'] = curr_p
                    if curr_p < info['max_price'] * (1 - stop_loss_pct): triggered_slots.append(s_id)

            for s_id in triggered_slots:
                info = slots[s_id]
                a_idx = info['asset_idx']
                sell_price = next_prices[a_idx]
                shares = info['shares']
                sell_fee = shares * sell_price * 0.001425
                sell_tax = shares * sell_price * 0.003
                proceeds = shares * sell_price - sell_fee - sell_tax
                cash += proceeds

                name = self.code_to_name[self.assets[a_idx]]
                trades2_log.append({
                    '買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx], '股票名稱': name,
                    'T+1日買進價格': info['entry_price'], '股數': shares, '賣出訊號日期': date,
                    'T+1日賣出價格': sell_price, '損益': proceeds - info['actual_cost'],
                    '報酬率': (proceeds / info['actual_cost']) - 1, '買進原因': '符合趨勢', '賣出原因': '停損機制'
                })
                trades_log.append({
                    '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出', '價格': sell_price,
                    '股數': shares, '動能值': f"{roc[i][a_idx]*100:.2f}%", '標的名稱': name, '原因': '停損',
                    '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax
                })
                slots[s_id] = None

            # C. 再平衡
            is_rebalance_day = (i - loop_start) % rebalance_interval == 0
            if is_rebalance_day:
                current_reasons = []
                top_3_signals = []
                exclusion_reasons = []
                sorted_all = np.argsort(roc[i])[::-1]
                for idx in sorted_all:
                    if len(top_3_signals) >= 3: break
                    p, s, r = current_prices[idx], sma[i][idx], roc[i][idx]
                    if p > s and r > 0: top_3_signals.append(idx)
                    else:
                        name = self.code_to_name[self.assets[idx]]
                        if r <= 0: exclusion_reasons.append(f"{name} ROC < 0")
                        elif p <= s: exclusion_reasons.append(f"{name} 價格 < SMA")

                # 賣出不在名單中的持股
                signal_to_slot_map = {}
                for s_id, info in slots.items():
                    if info and 'asset_idx' in info:
                        if info['asset_idx'] in top_3_signals: signal_to_slot_map[info['asset_idx']] = s_id
                        else:
                            a_idx = info['asset_idx']
                            sell_price = next_prices[a_idx]
                            shares = info['shares']
                            sell_fee = shares * sell_price * 0.001425
                            sell_tax = shares * sell_price * 0.003
                            proceeds = shares * sell_price - sell_fee - sell_tax
                            cash += proceeds
                            name = self.code_to_name[self.assets[a_idx]]
                            trades2_log.append({
                                '買進訊號日期': info['entry_date'], '股票代號': self.assets[a_idx], '股票名稱': name,
                                'T+1日買進價格': info['entry_price'], '股數': shares, '賣出訊號日期': date,
                                'T+1日賣出價格': sell_price, '損益': proceeds - info['actual_cost'],
                                '報酬率': (proceeds / info['actual_cost']) - 1, '買進原因': '符合趨勢', '賣出原因': '再平衡排名外'
                            })
                            trades_log.append({
                                '訊號日期': date, '股票代號': self.assets[a_idx], '狀態': '賣出', '價格': sell_price,
                                '股數': shares, '動能值': f"{roc[i][a_idx]*100:.2f}%", '標的名稱': name, '原因': '再平衡',
                                '買入手續費': 0, '賣出手續費': sell_fee, '賣出交易稅': sell_tax
                            })
                            slots[s_id] = None

                # 買入新訊號
                slot_cap = self.initial_capital / 3
                for sig in top_3_signals:
                    if sig in signal_to_slot_map:
                        name = self.code_to_name[self.assets[sig]]
                        trades_log.append({
                            '訊號日期': date, '股票代號': self.assets[sig], '狀態': '保持', '價格': current_prices[sig],
                            '股數': slots[signal_to_slot_map[sig]]['shares'], '動能值': f"{roc[i][sig]*100:.2f}%",
                            '標的名稱': name, '原因': '趨勢持續', '買入手續費': 0, '賣出手續費': 0, '賣出交易稅': 0
                        })
                        continue

                    available_slot = next((sid for sid, data in slots.items() if data is None), None)
                    if available_slot is not None:
                        buy_price_exec = next_prices[sig]
                        shares = (int(slot_cap // (buy_price_exec * 1.001425)) // 1000) * 1000
                        if shares > 0:
                            cost = shares * buy_price_exec * 1.001425
                            if cash >= cost:
                                cash -= cost
                                name = self.code_to_name[self.assets[sig]]
                                slots[available_slot] = {
                                    'asset_idx': sig, 'shares': shares, 'max_price': buy_price_exec,
                                    'entry_date': date, 'entry_price': buy_price_exec, 'actual_cost': cost
                                }
                                trades_log.append({
                                    '訊號日期': date, '股票代號': self.assets[sig], '狀態': '買進', '價格': buy_price_exec,
                                    '股數': shares, '動能值': f"{roc[i][sig]*100:.2f}%", '標的名稱': name, '原因': '符合趨勢',
                                    '買入手續費': shares * buy_price_exec * 0.001425, '賣出手續費': 0, '賣出交易稅': 0
                                })
                            else: current_reasons.append(f"資金不足無法買入 {self.code_to_name[self.assets[sig]]}")
                        else: current_reasons.append(f"單價過高無法買入整張 {self.code_to_name[self.assets[sig]]}")

                if len(top_3_signals) < 3:
                    rebal_msg = f"當次再平衡僅有 {len(top_3_signals)} 檔符合標準"
                    if exclusion_reasons: rebal_msg += f"；排除原因: {'、'.join(exclusion_reasons[:2])}"
                    current_reasons.append(rebal_msg)

            # D. 持股備註
            count = sum(1 for s in slots.values() if s is not None)
            final_remark = "；".join(list(dict.fromkeys(current_reasons))) if count < 3 else ""
            holdings_history.append({
                'Date': date, 'Holdings': ", ".join(h_names), 'Count': count, '現金': cash, '總資產': total_equity, '補充說明': final_remark
            })

        return pd.DataFrame(equity_curve_data), pd.DataFrame(trades_log), pd.DataFrame(holdings_history), pd.DataFrame(trades2_log), pd.DataFrame(daily_details)

# ==========================================
# 2. 執行回測與產出檔案
# ==========================================
def main():
    DATA_FILE = '個股合-1.xlsx'
    SMA_PERIOD = 54
    ROC_PERIOD = 52
    STOP_LOSS_PCT = 0.09
    REBALANCE = 9
    INITIAL_CAPITAL = 30000000

    prices, code_to_name = clean_data(DATA_FILE)
    bt = BacktesterV2(prices, code_to_name, INITIAL_CAPITAL)

    start_date = '2019-01-02'
    end_date = '2025-12-31'

    print(f"正在執行全期間交付成果產出 (SMA={SMA_PERIOD}, ROC={ROC_PERIOD}, SL={STOP_LOSS_PCT*100:.1f}%, Reb={REBALANCE})...")
    eq_df, trades, hold, trades2, daily = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE, start_date, end_date)

    from run_wfa import calculate_metrics
    cagr, mdd, calmar = calculate_metrics(eq_df)

    # 1. 產出 Excel
    OUTPUT_EXCEL = 'trendstrategy_results_equity-new.xlsx'
    with pd.ExcelWriter(OUTPUT_EXCEL, engine='xlsxwriter') as writer:
        pd.DataFrame([
            {'項目': '年化報酬率 (CAGR)', '數值': f"{cagr:.2%}"},
            {'項目': '最大回撤 (MaxDD)', '數值': f"{mdd:.2%}"},
            {'項目': 'Calmar Ratio', '數值': f"{calmar:.2f}"},
            {'項目': '初始資金', '數值': f"{INITIAL_CAPITAL:,}"},
            {'項目': '期末淨值', '數值': f"{eq_df['權益'].iloc[-1]:,.0f}"},
            {'項目': '版本', '數值': 'Equity-New (SMA54/ROC52/SL9/Reb9)'}
        ]).to_excel(writer, sheet_name='Summary', index=False)

        eq_df.to_excel(writer, sheet_name='Equity_Curve', index=False)
        hold.to_excel(writer, sheet_name='Equity_Hold', index=False)
        trades.to_excel(writer, sheet_name='Trades', index=False)
        trades2.to_excel(writer, sheet_name='Trades2', index=False)
        daily.to_excel(writer, sheet_name='Daily', index=False)

        # 插入 Equity Curve 圖表
        workbook = writer.book
        curves_sheet = writer.sheets['Equity_Curve']
        chart = workbook.add_chart({'type': 'line'})
        max_row = len(eq_df)
        chart.add_series({
            'name': 'Equity Curve',
            'categories': ['Equity_Curve', 1, 0, max_row, 0],
            'values': ['Equity_Curve', 1, 1, max_row, 1],
        })
        chart.set_title({'name': 'Full Period Equity Curve (SMA54/ROC52)'})
        curves_sheet.insert_chart('D2', chart)

    # 2. 產出 Markdown
    OUTPUT_MD = 'reproduce_report_equity-new.md'
    md_content = f"""# Asset Class Trend Following 策略回測報告 (Equity-New)

## 策略參數
- **SMA 週期**: {SMA_PERIOD}
- **ROC 週期**: {ROC_PERIOD}
- **追蹤停損**: {STOP_LOSS_PCT*100:.1f}%
- **再平衡頻率**: {REBALANCE} 交易日

## 全期間績效 (2019-2025)
- **年化報酬率 (CAGR)**: {cagr:.2%}
- **最大回撤 (MaxDD)**: {mdd:.2%}
- **卡瑪比率 (Calmar Ratio)**: {calmar:.2f}
- **期末淨值**: ${eq_df['權益'].iloc[-1]:,.0f}

## 核心邏輯
- 採用 **T 訊號，T+1 執行** 模式。
- 嚴格遵守 **1,000 股整數倍** 交易。
- 槽位化管理：初始資金均分為三等份，每筆投資上限為 1,000 萬。
"""
    with open(OUTPUT_MD, 'w', encoding='utf-8') as f:
        f.write(md_content)

    # 3. 產出 Jupyter Notebook
    nb = nbf.v4.new_notebook()
    nb.cells.append(nbf.v4.new_markdown_cell("# Asset Class Trend Following 策略回測 (Equity-New)"))
    # 讀取當前檔案的代碼作為 Notebook 的回測核心
    with open(__file__, 'r', encoding='utf-8') as f:
        code = f.read()
    nb.cells.append(nbf.v4.new_code_cell(code))
    with open('trendstrategy_equity-new.ipynb', 'w', encoding='utf-8') as f:
        nbf.write(nb, f)

    print(f"所有交付成果檔案 (Excel, MD, IPYNB) 已成功產出。")
    print(f"CAGR: {cagr:.2%}, MDD: {mdd:.2%}, Calmar: {calmar:.2f}")

if __name__ == "__main__":
    main()
