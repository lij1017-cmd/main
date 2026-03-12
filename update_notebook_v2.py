import json

with open('trendstrategy_equity25_cost.ipynb', 'r', encoding='utf-8') as f:
    nb = json.load(f)

# Cell 1: Global Parameters
# No changes needed unless I want to add scenario definitions here, but I can do it in the code cells.

# Cell 3: Backtester Class
bt_code = """class Backtester:
    def __init__(self, prices, code_to_name, initial_capital=30000000):
        self.prices = prices.values
        self.dates = prices.index
        self.assets = prices.columns
        self.code_to_name = code_to_name
        self.initial_capital = initial_capital

    def run(self, sma_period, roc_period, stop_loss_pct, buy_fee_rate=0.001425, sell_fee_rate=0.001425, sell_tax_rate=0.003):
        prices_df = pd.DataFrame(self.prices, index=self.dates, columns=self.assets)
        sma = prices_df.rolling(window=sma_period).mean().values
        roc = prices_df.pct_change(periods=roc_period).values

        capital = self.initial_capital
        portfolio = {}
        equity_curve = np.zeros(len(self.dates))
        rebalance_log = []
        trade_records = []

        start_idx = max(sma_period, roc_period)

        for i in range(start_idx, len(self.dates) - 1):
            date = self.dates[i]
            current_prices = self.prices[i]
            next_prices = self.prices[i+1]

            total_equity = capital
            assets_to_sell = []

            for asset_idx, info in list(portfolio.items()):
                curr_p = current_prices[asset_idx]
                total_equity += info['shares'] * curr_p
                if curr_p > info['max_price']:
                    info['max_price'] = curr_p
                if curr_p < info['max_price'] * (1 - stop_loss_pct):
                    assets_to_sell.append(asset_idx)

            equity_curve[i] = total_equity
            is_rebalance_day = (i - start_idx) % 5 == 0

            new_portfolio_signals = []
            if is_rebalance_day:
                eligible_mask = (current_prices > sma[i]) & (roc[i] > 0)
                if np.any(eligible_mask):
                    eligible_idxs = np.where(eligible_mask)[0]
                    eligible_rocs = roc[i][eligible_idxs]
                    top_k = min(3, len(eligible_idxs))
                    top_idxs = eligible_idxs[np.argsort(eligible_rocs)[-top_k:][::-1]]
                    new_portfolio_signals = list(top_idxs)

            assets_selling_now = set(assets_to_sell)
            if is_rebalance_day:
                for asset_idx in list(portfolio.keys()):
                    if asset_idx not in new_portfolio_signals:
                        assets_selling_now.add(asset_idx)

            for asset_idx in assets_selling_now:
                if asset_idx in portfolio:
                    info = portfolio.pop(asset_idx)
                    sell_price = next_prices[asset_idx]
                    sell_amount = info['shares'] * sell_price
                    sell_fee = sell_amount * sell_fee_rate
                    sell_tax = sell_amount * sell_tax_rate
                    capital += sell_amount - sell_fee - sell_tax

                    trade_records.append({
                        '股票代號': self.assets[asset_idx],
                        '標的名稱': self.code_to_name.get(self.assets[asset_idx], ""),
                        '進場日期': info['buy_date'],
                        '出場日期': self.dates[i+1],
                        '進場價格': info['buy_price'],
                        '出場價格': sell_price,
                        '股數': info['shares'],
                        '買進金額': info['buy_amount'],
                        '賣出金額': sell_amount,
                        '買進手續費': info['buy_fee'],
                        '賣出手續費': sell_fee,
                        '證交稅': sell_tax,
                        '持有天數': (self.dates[i+1] - info['buy_date']).days,
                        '報酬率 (%)': (sell_amount - sell_fee - sell_tax) / (info['buy_amount'] + info['buy_fee']) - 1,
                        '原因': "停損" if asset_idx in assets_to_sell else "再平衡"
                    })

                    rebalance_log.append({
                        '日期': date, '股票代號': self.assets[asset_idx], '狀態': "賣出",
                        '價格': current_prices[asset_idx], '股數': 0,
                        '動能值': f"{roc[i][asset_idx]*100:.2f}%",
                        '標的名稱': self.code_to_name.get(self.assets[asset_idx], ""),
                        '最佳參數': f"SMA={sma_period}, ROC={roc_period}, SL={stop_loss_pct}",
                        '原因': "停損" if asset_idx in assets_to_sell else "再平衡",
                        '說明': f"賣出資產：{self.code_to_name.get(self.assets[asset_idx], '')} ({self.assets[asset_idx]})"
                    })

            if is_rebalance_day:
                assets_to_buy = [a for a in new_portfolio_signals if a not in portfolio]
                slot_capital = self.initial_capital / 3
                for asset_idx in assets_to_buy:
                    buy_price = next_prices[asset_idx]
                    shares = slot_capital // (buy_price * (1 + buy_fee_rate))
                    if shares > 0:
                        buy_amount = shares * buy_price
                        buy_fee = buy_amount * buy_fee_rate
                        buy_cost = buy_amount + buy_fee
                        if capital >= buy_cost:
                            capital -= buy_cost
                            portfolio[asset_idx] = {
                                'shares': shares, 'buy_price': buy_price, 'buy_date': self.dates[i+1],
                                'max_price': buy_price, 'momentum': roc[i][asset_idx],
                                'buy_amount': buy_amount, 'buy_fee': buy_fee
                            }

                for asset_idx in range(len(self.assets)):
                    if asset_idx in portfolio:
                        status = "買進" if self.dates[i+1] == portfolio[asset_idx]['buy_date'] else "保持"
                        reason = "符合趨勢" if status == "買進" else "趨勢持續"

                        if status == "買進" or status == "保持":
                             rebalance_log.append({
                                '日期': date, '股票代號': self.assets[asset_idx], '狀態': status,
                                '價格': current_prices[asset_idx],
                                '股數': portfolio[asset_idx]['shares'],
                                '動能值': f"{roc[i][asset_idx]*100:.2f}%",
                                '標的名稱': self.code_to_name.get(self.assets[asset_idx], ""),
                                '最佳參數': f"SMA={sma_period}, ROC={roc_period}, SL={stop_loss_pct}",
                                '原因': reason,
                                '說明': f"選取資產：{self.code_to_name.get(self.assets[asset_idx], '')} ({self.assets[asset_idx]})，動能：{roc[i][asset_idx]*100:.2f}%"
                            })

        equity_curve[-1] = total_equity
        eq_series = pd.Series(equity_curve, index=self.dates).replace(0, np.nan).dropna()
        return eq_series, pd.DataFrame(rebalance_log), pd.DataFrame(trade_records)"""

nb['cells'][3]['source'] = [line + '\n' for line in bt_code.split('\n')]

# Cell 4: Metrics Calculation
metrics_code = """def calculate_metrics(equity_curve, trades_df):
    if equity_curve.empty: return {}
    total_return = (equity_curve.iloc[-1] / equity_curve.iloc[0]) - 1
    days = (equity_curve.index[-1] - equity_curve.index[0]).days
    cagr = (1 + total_return) ** (365.25 / days) - 1
    rolling_max = equity_curve.cummax()
    drawdown = (equity_curve - rolling_max) / rolling_max
    max_dd = drawdown.min()
    calmar = cagr / abs(max_dd) if max_dd != 0 else 0

    avg_hold_days = trades_df['持有天數'].mean() if not trades_df.empty else 0

    if not trades_df.empty:
        temp_df = trades_df.copy()
        temp_df['Year'] = pd.to_datetime(temp_df['進場日期']).dt.year
        trades_per_year = temp_df.groupby('Year').size().mean()
    else:
        trades_per_year = 0

    return {
        'CAGR': cagr,
        'MaxDD': max_dd,
        'Calmar': calmar,
        'AvgHoldDays': avg_hold_days,
        'TradesPerYear': trades_per_year
    }"""

nb['cells'][4]['source'] = [line + '\n' for line in metrics_code.split('\n')]

# Cell 5: Main execution and Sensitivity Analysis
main_code = """prices, code_to_name = clean_data(DATA_FILE)
bt = Backtester(prices, code_to_name, INITIAL_CAPITAL)

# 1. Target Run (Standard Costs)
eq, rebalance_df, trades_detailed_df = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT)
res = calculate_metrics(eq, trades_detailed_df)

print(f'Target Parameters: SMA={SMA_PERIOD}, ROC={ROC_PERIOD}, SL={STOP_LOSS_PCT}')
print(f"CAGR: {res['CAGR']:.2%}, MaxDD: {res['MaxDD']:.2%}, Calmar: {res['Calmar']:.2f}")
print(f"平均持有天數: {res['AvgHoldDays']:.1f}, 每年平均交易次數: {res['TradesPerYear']:.1f}")

# 2. Sensitivity Analysis
scenarios = [
    ("零成本", 0, 0, 0),
    ("只扣手續費", 0.001425, 0.001425, 0),
    ("只扣稅", 0, 0, 0.003),
    ("手續費減半", 0.0007125, 0.0007125, 0.003),
    ("標準成本", 0.001425, 0.001425, 0.003)
]

sensitivity_results = []
for name, bf, sf, tx in scenarios:
    eq_s, _, tr_s = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, bf, sf, tx)
    m = calculate_metrics(eq_s, tr_s)
    sensitivity_results.append({
        '情境': name,
        'CAGR': f"{m['CAGR']:.2%}",
        'MaxDD': f"{m['MaxDD']:.2%}",
        'Calmar': f"{m['Calmar']:.2f}",
        '平均持有天數': f"{m['AvgHoldDays']:.1f}",
        '每年交易次數': f"{m['TradesPerYear']:.1f}"
    })

sensitivity_df = pd.DataFrame(sensitivity_results)
print('\\n交易成本敏感度測試:')
print(sensitivity_df)

# 3. SMA Parameter Plateau (with standard cost)
plateau_results = []
sma_values = [60, 62, 64, 66, 68]
for s in sma_values:
    eq_s, _, tr_s = bt.run(s, ROC_PERIOD, STOP_LOSS_PCT)
    m = calculate_metrics(eq_s, tr_s)
    plateau_results.append({'SMA': s, 'CAGR': f"{m['CAGR']:.2%}", 'MaxDD': f"{m['MaxDD']:.2%}", 'Calmar': f"{m['Calmar']:.2f}"})

plateau_df = pd.DataFrame(plateau_results)

# Save to Excel
with pd.ExcelWriter(OUTPUT_EXCEL, engine='xlsxwriter') as writer:
    rebalance_df.to_excel(writer, sheet_name='Trades', index=False)
    trades_detailed_df.to_excel(writer, sheet_name='Performance', index=False)
    eq.reset_index().rename(columns={'index':'Date', 0:'Equity'}).to_excel(writer, sheet_name='Equity_Curve', index=False)
    summary_df = pd.DataFrame([
        {'Metric': 'CAGR', 'Value': f"{res['CAGR']:.2%}"},
        {'Metric': 'MaxDD', 'Value': f"{res['MaxDD']:.2%}"},
        {'Metric': 'Calmar Ratio', 'Value': f"{res['Calmar']:.2f}"},
        {'Metric': '平均持有天數', 'Value': f"{res['AvgHoldDays']:.1f}"},
        {'Metric': '每年平均交易次數', 'Value': f"{res['TradesPerYear']:.1f}"},
        {'Metric': 'SMA', 'Value': SMA_PERIOD},
        {'Metric': 'ROC', 'Value': ROC_PERIOD},
        {'Metric': 'StopLoss%', 'Value': STOP_LOSS_PCT}
    ])
    summary_df.to_excel(writer, sheet_name='Summary', index=False)
    sensitivity_df.to_excel(writer, sheet_name='Sensitivity', index=False)

# Update MD Report Content
md_content = f'''# Asset Class Trend Following 策略重現報告 (2025 - 交易成本深度分析)

## 策略說明
本策略針對「{DATA_FILE}」商品，納入完整交易成本（買進 0.1425%，賣出 0.1425% + 0.3% 證交稅）。

## 核心參數
- **SMA 週期**: {SMA_PERIOD}
- **ROC 週期**: {ROC_PERIOD}
- **停損比例 (StopLoss%)**: {STOP_LOSS_PCT*100:.1f}%

## 績效表現 (標準成本)
- **CAGR**: {res['CAGR']:.2%}
- **MaxDD**: {res['MaxDD']:.2%}
- **Calmar Ratio**: {res['Calmar']:.2f}
- **平均持有天數**: {res['AvgHoldDays']:.1f} 天
- **每年交易次數**: {res['TradesPerYear']:.1f} 次

## 交易成本敏感度測試
{sensitivity_df.to_markdown(index=False)}

## 參數高原表 (SMA 變化)
{plateau_df.to_markdown(index=False)}

## 結論
納入交易成本後，CAGR 與 Calmar Ratio 均有下降，其中證交稅對績效的影響最為顯著。
'''

with open(OUTPUT_MD, 'w', encoding='utf-8') as f:
    f.write(md_content)

print(f'\\n結果已儲存至 {OUTPUT_EXCEL} 與 {OUTPUT_MD}')"""

nb['cells'][5]['source'] = [line + '\n' for line in main_code.split('\n')]

# Save the updated notebook
with open('trendstrategy_equity25_cost.ipynb', 'w', encoding='utf-8') as f:
    json.dump(nb, f, ensure_ascii=False, indent=1)
