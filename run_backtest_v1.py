import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
import pickle
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

# Global Parameters
SMA_BEST = 33
ROC_BEST = 58
SL_BEST = 0.095
INITIAL_CAPITAL = 30000000
DATA_FILE = '個股2.xlsx'
SUFFIX = 'equity2025(成)V1'

def run_segmented_backtest(prices, code_to_name, periods):
    results = []
    full_bt = Backtester(prices, code_to_name, INITIAL_CAPITAL)
    full_eq, full_trades, _ = full_bt.run(SMA_BEST, ROC_BEST, SL_BEST)
    for name, start_date, end_date in periods:
        mask = (full_eq.index >= start_date) & (full_eq.index <= end_date)
        period_eq = full_eq.loc[mask]
        if period_eq.empty: continue
        period_eq = period_eq / period_eq.iloc[0] * INITIAL_CAPITAL
        c, m, cl, r = calculate_metrics(period_eq)
        period_trades = full_trades[(full_trades['日期'] >= pd.to_datetime(start_date)) & (full_trades['日期'] <= pd.to_datetime(end_date))]
        w = calculate_win_rate_v1(period_trades, prices)
        results.append({'期間': name, 'CAGR': f"{c:.2%}", 'MaxDD': f"{m:.2%}", 'Calmar': f"{cl:.2f}", 'WinRate': f"{w:.2%}"})
    return pd.DataFrame(results)

if __name__ == "__main__":
    prices, code_to_name = clean_data(DATA_FILE)
    bt = Backtester(prices, code_to_name, INITIAL_CAPITAL)
    eq, trades, hold = bt.run(SMA_BEST, ROC_BEST, SL_BEST)
    cagr, mdd, calmar, ret = calculate_metrics(eq)
    win_rate = calculate_win_rate_v1(trades, prices)

    periods = [('2024/10 - 2025/06', '2024-10-01', '2025-06-30'), ('2025/07 - 2025/12', '2025-07-01', '2025-12-31')]
    segment_df = run_segmented_backtest(prices, code_to_name, periods)

    plateau_data = []
    for s in [SMA_BEST-4, SMA_BEST-2, SMA_BEST, SMA_BEST+2, SMA_BEST+4]:
        eq_s, _, _ = bt.run(s, ROC_BEST, SL_BEST)
        c_s, m_s, cl_s, _ = calculate_metrics(eq_s)
        plateau_data.append({'SMA': s, 'CAGR': f"{c_s:.2%}", 'MaxDD': f"{m_s:.2%}", 'Calmar': f"{cl_s:.2f}"})
    plateau_df = pd.DataFrame(plateau_data)

    OUTPUT_EXCEL = f'trendstrategy_results_{SUFFIX}.xlsx'
    with pd.ExcelWriter(OUTPUT_EXCEL, engine='xlsxwriter') as writer:
        pd.DataFrame([{'項目': '年化報酬率 (CAGR)', '數值': f"{cagr:.2%}"}, {'項目': '最大回撤 (MaxDD)', '數值': f"{mdd:.2%}"}, {'項目': 'Calmar Ratio', '數值': f"{calmar:.2f}"}, {'項目': '勝率 (Win Rate)', '數值': f"{win_rate:.2%}"}, {'項目': '最佳參數', '數值': f"SMA={SMA_BEST}, ROC={ROC_BEST}, SL={SL_BEST*100:.1f}%"}]).to_excel(writer, sheet_name='Summary', index=False)
        segment_df.to_excel(writer, sheet_name='Segment_Performance', index=False)
        eq.reset_index().to_excel(writer, sheet_name='Equity_Curve', index=False)
        hold.to_excel(writer, sheet_name='Equity_Hold', index=False)
        trades['最佳參數'] = f"SMA={SMA_BEST}, ROC={ROC_BEST}, SL={SL_BEST}"
        trades[['日期', '股票代號', '狀態', '價格', '股數', '動能值', '標的名稱', '最佳參數', '原因', '說明']].to_excel(writer, sheet_name='Trades', index=False)

    OUTPUT_MD = f'reproduce_report_{SUFFIX}.md'
    md_content = f"# Asset Class Trend Following 策略重現報告 ({SUFFIX})\n\n## 策略摘要\n此測試應用於 2024/10/01 至 2025/12/31 期間，並劃分不同期間進行績效測試。\n\n## 核心參數\n- **SMA 週期**: {SMA_BEST} | **ROC 週期**: {ROC_BEST} | **停損比例**: {SL_BEST*100:.1f}%\n\n## 績效表現 (總計)\n- **CAGR**: {cagr:.2%} | **MaxDD**: {mdd:.2%} | **Calmar**: {calmar:.2f} | **WinRate**: {win_rate:.2%}\n\n## 不同期間績效表現\n{segment_df.to_markdown(index=False)}\n\n## 參數高原表 (SMA 敏感度分析)\n{plateau_df.to_markdown(index=False)}\n\n## 交易成本與邏輯\n1. **交易成本**: 包含買進 0.1425% 手續費，賣出 0.1425% 手續費與 0.3% 證交稅。\n2. **邏輯**: T 日訊號，T+1 日收盤價執行。最高價回落停損機制。\n"
    with open(OUTPUT_MD, 'w', encoding='utf-8') as f: f.write(md_content)

    nb = nbf.v4.new_notebook()
    with open('backtest.py', 'r') as f: blines = f.readlines()
    try: main_idx = next(i for i, l in enumerate(blines) if 'if __name__ == "__main__":' in l); bcode = "".join(blines[:main_idx])
    except: bcode = "".join(blines)
    nb.cells.extend([nbf.v4.new_markdown_cell(f"# 策略回測 ({SUFFIX})"), nbf.v4.new_code_cell(f"SMA_PERIOD={SMA_BEST}\nROC_PERIOD={ROC_BEST}\nSTOP_LOSS_PCT={SL_BEST}\nDATA_FILE='{DATA_FILE}'"), nbf.v4.new_code_cell("import pandas as pd\nimport numpy as np\nimport matplotlib.pyplot as plt\ndef clean_data(f):\n    df=pd.read_excel(f,header=None);c=df.iloc[0,2:].values;n=df.iloc[1,2:].values;d=pd.to_datetime(df.iloc[2:,1]);p=df.iloc[2:,2:].astype(float);p.index=d;p.columns=c;p=p.ffill().bfill();return p,dict(zip(c,n))"), nbf.v4.new_code_cell(bcode), nbf.v4.new_code_cell("prices,c2n=clean_data(DATA_FILE);bt=Backtester(prices,c2n);eq,tr,ho=bt.run(SMA_PERIOD,ROC_PERIOD,STOP_LOSS_PCT);plt.figure(figsize=(12,6));plt.plot(eq);plt.show()")])
    with open(f'trendstrategy_equity_{SUFFIX}.ipynb', 'w', encoding='utf-8') as f: nbf.write(nb, f)
