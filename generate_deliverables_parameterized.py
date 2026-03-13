import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
import nbformat as nbf
from backtest import Backtester, calculate_metrics

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

def calculate_win_rate_v1(trades_df, prices_df):
    if trades_df.empty: return 0.0
    sells = trades_df[trades_df['狀態'] == '賣出'].copy()
    buys = trades_df[trades_df['狀態'] == '買進'].copy()
    win_count = 0
    total_trades = 0
    for code in sells['股票代號'].unique():
        code_sells = sells[sells['股票代號'] == code].sort_values('日期')
        code_buys = buys[buys['股票代號'] == code].sort_values('日期')
        for idx, sell_row in code_sells.iterrows():
            possible_buys = code_buys[code_buys['日期'] < sell_row['日期']]
            if not possible_buys.empty:
                buy_row = possible_buys.iloc[-1]
                code_buys = code_buys.drop(buy_row.name)
                signal_date_buy = buy_row['日期']
                signal_date_sell = sell_row['日期']
                try:
                    buy_idx = prices_df.index.get_loc(signal_date_buy)
                    sell_idx = prices_df.index.get_loc(signal_date_sell)
                    if buy_idx + 1 < len(prices_df) and sell_idx + 1 < len(prices_df):
                        buy_price_exec = prices_df.iloc[buy_idx + 1][code]
                        sell_price_exec = prices_df.iloc[sell_idx + 1][code]
                        ret = (sell_price_exec * 0.995575) / (buy_price_exec * 1.001425) - 1
                        if ret > 0: win_count += 1
                        total_trades += 1
                except: continue
    return win_count / total_trades if total_trades > 0 else 0.0

def generate_deliverables(sma, roc, sl, suffix, data_file):
    prices, code_to_name = clean_data(data_file)
    initial_capital = 30000000
    bt = Backtester(prices, code_to_name, initial_capital)
    eq, trades, hold = bt.run(sma, roc, sl)
    cagr, mdd, calmar, ret = calculate_metrics(eq)
    win_rate = calculate_win_rate_v1(trades, prices)

    # Segmented Testing
    periods = [('2024/10 - 2025/06', '2024-10-01', '2025-06-30'), ('2025/07 - 2025/12', '2025-07-01', '2025-12-31')]
    segment_results = []
    for name, start, end in periods:
        mask = (eq.index >= start) & (eq.index <= end)
        p_eq = eq.loc[mask]
        if p_eq.empty: continue
        p_eq = p_eq / p_eq.iloc[0] * initial_capital
        c, m, cl, r = calculate_metrics(p_eq)
        p_trades = trades[(trades['日期'] >= pd.to_datetime(start)) & (trades['日期'] <= pd.to_datetime(end))]
        w = calculate_win_rate_v1(p_trades, prices)
        segment_results.append({'期間': name, 'CAGR': f"{c:.2%}", 'MaxDD': f"{m:.2%}", 'Calmar': f"{cl:.2f}", 'WinRate': f"{w:.2%}"})
    segment_df = pd.DataFrame(segment_results)

    # Plateau
    plateau_data = []
    for s in [sma-4, sma-2, sma, sma+2, sma+4]:
        eq_s, _, _ = bt.run(s, roc, sl)
        c_s, m_s, cl_s, _ = calculate_metrics(eq_s)
        plateau_data.append({'SMA': s, 'CAGR': f"{c_s:.2%}", 'MaxDD': f"{m_s:.2%}", 'Calmar': f"{cl_s:.2f}"})
    plateau_df = pd.DataFrame(plateau_data)

    # Excel
    with pd.ExcelWriter(f'trendstrategy_results_{suffix}.xlsx', engine='xlsxwriter') as writer:
        pd.DataFrame([{'項目': '年化報酬率 (CAGR)', '數值': f"{cagr:.2%}"}, {'項目': '最大回撤 (MaxDD)', '數值': f"{mdd:.2%}"}, {'項目': 'Calmar Ratio', '數值': f"{calmar:.2f}"}, {'項目': '勝率 (Win Rate)', '數值': f"{win_rate:.2%}"}, {'項目': '最佳參數', '數值': f"SMA={sma}, ROC={roc}, SL={sl*100:.1f}%"}]).to_excel(writer, sheet_name='Summary', index=False)
        segment_df.to_excel(writer, sheet_name='Segment_Performance', index=False)
        eq.reset_index().to_excel(writer, sheet_name='Equity_Curve', index=False)
        hold.to_excel(writer, sheet_name='Equity_Hold', index=False)
        trades['最佳參數'] = f"SMA={sma}, ROC={roc}, SL={sl}"
        trades[['日期', '股票代號', '狀態', '價格', '股數', '動能值', '標的名稱', '最佳參數', '原因', '說明']].to_excel(writer, sheet_name='Trades', index=False)

    # Markdown
    with open(f'reproduce_report_{suffix}.md', 'w', encoding='utf-8') as f:
        f.write(f"# 策略重現報告 ({suffix})\n\n## 核心參數\n- SMA: {sma} | ROC: {roc} | StopLoss: {sl*100:.1f}%\n\n## 績效表現\n- CAGR: {cagr:.2%} | MaxDD: {mdd:.2%} | Calmar: {calmar:.2f} | WinRate: {win_rate:.2%}\n\n## 分段績效\n{segment_df.to_markdown(index=False)}\n\n## 參數高原表\n{plateau_df.to_markdown(index=False)}\n")

    # Notebook
    nb = nbf.v4.new_notebook()
    with open('backtest.py', 'r') as f: blines = f.readlines()
    try: idx = next(i for i, l in enumerate(blines) if 'if __name__ == "__main__":' in l); bcode = "".join(blines[:idx])
    except: bcode = "".join(blines)
    nb.cells.extend([nbf.v4.new_markdown_cell(f"# 策略回測 ({suffix})"), nbf.v4.new_code_cell(f"SMA_PERIOD={sma}\nROC_PERIOD={roc}\nSTOP_LOSS_PCT={sl}\nDATA_FILE='{data_file}'"), nbf.v4.new_code_cell("import pandas as pd\nimport numpy as np\nimport matplotlib.pyplot as plt\ndef clean_data(f):\n    df=pd.read_excel(f,header=None);c=df.iloc[0,2:].values;n=df.iloc[1,2:].values;d=pd.to_datetime(df.iloc[2:,1]);p=df.iloc[2:,2:].astype(float);p.index=d;p.columns=c;p=p.ffill().bfill();return p,dict(zip(c,n))"), nbf.v4.new_code_cell(bcode), nbf.v4.new_code_cell("prices,c2n=clean_data(DATA_FILE);bt=Backtester(prices,c2n);eq,tr,ho=bt.run(SMA_PERIOD,ROC_PERIOD,STOP_LOSS_PCT);plt.figure(figsize=(12,6));plt.plot(eq);plt.show()")])
    with open(f'trendstrategy_equity_{suffix}.ipynb', 'w', encoding='utf-8') as f: nbf.write(nb, f)

if __name__ == "__main__":
    trials = [
        (33, 58, 0.095, 'equity2025(成)V1', '個股2.xlsx'),
        (35, 56, 0.09, 'equity2025(成)V2', '個股2.xlsx')
    ]
    for sma, roc, sl, suffix, data in trials:
        print(f"Generating deliverables for {suffix}...")
        generate_deliverables(sma, roc, sl, suffix, data)
