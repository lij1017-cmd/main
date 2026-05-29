import pandas as pd
import numpy as np
from backtest_vol import BacktesterVol, calculate_metrics_dual, clean_data

def main():
    filepath = '樣本集-1.xlsx'
    prices, volumes, code_to_name = clean_data(filepath)
    TRADING_CAP = 30000000
    AUTH_CAP = 150000000
    bt = BacktesterVol(prices, volumes, code_to_name, trading_capital=TRADING_CAP, authorized_capital=AUTH_CAP)

    # 目前參數
    SMA = 303
    ROC = 14
    MULT = 2.7
    BREADTH = 290

    # 1. 執行全期間 (Continuous)
    eq_full, _, _, _ = bt.run(SMA, ROC, vol_multiplier=MULT, breadth_window=BREADTH)

    # 2. 模擬 WFA 期間 (Reset)
    # 取一個高 CAGR 的 WFA 區間作為例子：2019-01-02 to 2021-12-31 (原本報告中 Trading CAGR 達 96%)
    wfa_start = '2019-01-02'
    wfa_end = '2021-12-31'
    eq_wfa, _, _, _ = bt.run(SMA, ROC, vol_multiplier=MULT, breadth_window=BREADTH, start_date=wfa_start, end_date=wfa_end)

    m_full = calculate_metrics_dual(eq_full, TRADING_CAP, AUTH_CAP)
    m_wfa = calculate_metrics_dual(eq_wfa, TRADING_CAP, AUTH_CAP)

    print("--- 數學對帳與一致性分析 ---")
    print(f"範例區間: {wfa_start} 至 {wfa_end}")
    print(f"WFA 模式 (Reset to 30M) Trading CAGR: {m_wfa['Trading CAGR']:.2%}")

    # 在全期間數據中定位相同區間
    mask = (eq_full['日期'] >= pd.to_datetime(wfa_start)) & (eq_full['日期'] <= pd.to_datetime(wfa_end))
    eq_in_full = eq_full[mask].copy()

    if not eq_in_full.empty:
        start_val = eq_in_full['權益'].iloc[0]
        end_val = eq_in_full['權益'].iloc[-1]
        days = (eq_in_full['日期'].iloc[-1] - eq_in_full['日期'].iloc[0]).days
        years = days / 365.25
        # 計算在連續模式下的該段 CAGR
        cont_cagr = (end_val / start_val)**(1/years) - 1
        print(f"全期間模式 (Continuous) 同區間 CAGR: {cont_cagr:.2%}")
        print(f"區間起始權益: {start_val:,.0f}")
        print(f"區間結束權益: {end_val:,.0f}")
        print(f"區間絕對獲利: {end_val - start_val:,.0f}")

        # 分析閒置資金
        # 在連續模式下，若權益已增長至 1 億，但買入上限仍是 3 檔各 1,000 萬 (共 3,000 萬)
        # 則閒置資金高達 7,000 萬。這會劇烈稀釋 CAGR。
        avg_equity = eq_in_full['權益'].mean()
        max_investment = 30000000.0
        idle_ratio = max(0, (avg_equity - max_investment) / avg_equity)
        print(f"平均權益: {avg_equity:,.0f}")
        print(f"最大投入上限: {max_investment:,.0f}")
        print(f"閒置資金佔比 (估計): {idle_ratio:.2%}")
        print("\n結論：WFA 模式每次重置為 30M，因此沒有閒置資金稀釋問題，導致 CAGR 極高。")
        print("全期間模式中，由於『固定 1,000 萬/slot』的限制，後期大量獲利無法再投入，導致 CAGR 被嚴重稀釋。")

if __name__ == "__main__":
    main()
