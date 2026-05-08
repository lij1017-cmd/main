import pandas as pd
import numpy as np
from backtest_breadth import clean_data, BacktesterBreadth, calculate_metrics

def main():
    DATA_FILE = '樣本集-1.xlsx'
    INITIAL_CAPITAL = 30000000

    # 策略參數 (equityV)
    SMA_PERIOD = 303
    ROC_PERIOD = 14
    STOP_LOSS_PCT = 0.0999
    REBALANCE = 9

    print(f"Loading data from {DATA_FILE}...")
    prices, volumes, code_to_name = clean_data(DATA_FILE)

    bt = BacktesterBreadth(prices, volumes, code_to_name, initial_capital=INITIAL_CAPITAL)

    print("Running baseline backtest (without breadth filter)...")
    eq_base, trades_base, hold_base, trades2_base, daily_base = bt.run(
        SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE, use_market_filter=False
    )

    print("Running optimized backtest (with breadth filter)...")
    eq_opt, trades_opt, hold_opt, trades2_opt, daily_opt = bt.run(
        SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE, use_market_filter=True
    )

    # 計算各年度績效
    years = [2019, 2020, 2021, 2022, 2023, 2024, 2025]
    annual_metrics = []

    def get_annual_metrics(eq_df, year):
        mask = (eq_df['日期'] >= f"{year}-01-01") & (eq_df['日期'] <= f"{year}-12-31")
        subset = eq_df[mask].copy()
        if subset.empty:
            return 0, 0, 0

        # 重新計算該年度的 MDD (相對於該年度內的最高點)
        subset['Annual_Max'] = subset['權益'].cummax()
        subset['Annual_DD'] = (subset['權益'] - subset['Annual_Max']) / subset['Annual_Max']

        cagr, _, _, _ = calculate_metrics(subset)
        mdd = subset['Annual_DD'].min()
        calmar = cagr / abs(mdd) if mdd != 0 else 0
        return cagr, mdd, calmar

    # 全期間績效 (2019-2025)
    mask_full = (eq_base['日期'] >= '2019-01-01') & (eq_base['日期'] <= '2025-12-31')
    res_base_full = eq_base[mask_full]
    c_base_f, m_base_f, cal_base_f, _ = calculate_metrics(res_base_full)

    mask_full_opt = (eq_opt['日期'] >= '2019-01-01') & (eq_opt['日期'] <= '2025-12-31')
    res_opt_full = eq_opt[mask_full_opt]
    c_opt_f, m_opt_f, cal_opt_f, _ = calculate_metrics(res_opt_full)

    comparison_data = []
    comparison_data.append({
        '期間': '全期間 (2019-2025)',
        '優化前 CAGR': f"{c_base_f:.2%}", '優化前 MDD': f"{m_base_f:.2%}", '優化前 Calmar': f"{cal_base_f:.2f}",
        '優化後 CAGR': f"{c_opt_f:.2%}", '優化後 MDD': f"{m_opt_f:.2%}", '優化後 Calmar': f"{cal_opt_f:.2f}"
    })

    for y in years:
        c_b, m_b, cal_b = get_annual_metrics(eq_base, y)
        c_o, m_o, cal_o = get_annual_metrics(eq_opt, y)
        comparison_data.append({
            '期間': f"{y}年",
            '優化前 CAGR': f"{c_b:.2%}", '優化前 MDD': f"{m_b:.2%}", '優化前 Calmar': f"{cal_b:.2f}",
            '優化後 CAGR': f"{c_o:.2%}", '優化後 MDD': f"{m_o:.2%}", '優化後 Calmar': f"{cal_o:.2f}"
        })

    df_comp = pd.DataFrame(comparison_data)

    # 產出 Excel
    OUTPUT_EXCEL = 'equityV-breadth.xlsx'
    with pd.ExcelWriter(OUTPUT_EXCEL, engine='xlsxwriter') as writer:
        df_comp.to_excel(writer, sheet_name='Summary', index=False)
        eq_opt.to_excel(writer, sheet_name='Equity_Curve', index=False)
        hold_opt.to_excel(writer, sheet_name='Equity_Hold', index=False)
        trades_opt.to_excel(writer, sheet_name='Trades', index=False)
        trades2_opt.to_excel(writer, sheet_name='Trades2', index=False)
        daily_opt.to_excel(writer, sheet_name='Daily', index=False)

        # 加入自動化圖表
        workbook = writer.book
        curves_sheet = writer.sheets['Equity_Curve']
        chart = workbook.add_chart({'type': 'line'})
        max_row = len(eq_opt)
        chart.add_series({
            'name': 'Equity Curve (Optimized)',
            'categories': ['Equity_Curve', 1, 0, max_row, 0],
            'values': ['Equity_Curve', 1, 1, max_row, 1],
        })
        chart.set_title({'name': 'Equity Curve with Market Breadth Filter'})
        chart.set_x_axis({'name': 'Date'})
        chart.set_y_axis({'name': 'Equity'})
        curves_sheet.insert_chart('G2', chart)

    # 產出 MD
    OUTPUT_MD = 'reproduce_equityV_breadth.md'
    md_content = f"""# Asset Class Trend Following 策略優化報告 (Market Breadth Filter)

## 1. 策略優化說明
本次優化在原有的 equityV 策略 (SMA 303, ROC 14, SL 9.99%, Reb 9) 基礎上，加入了**市場寬度濾網 (Market Breadth Filter)**。

### 市場寬度濾網規則：
- **定義**：計算全市場 131 檔標的中，收盤價高於其各自 SMA(200) 的標的佔比。
- **觸發條件**：當市場寬度 **< 35%** 時，視為市場環境轉弱。
- **執行動作**：觸發時進行**全清倉行為**，且在寬度恢復至 35% 以上前不持有任何股票。

---

## 2. 優化前後績效對比 (2019 - 2025)

| 期間 | 優化前 CAGR | 優化後 CAGR | 優化前 MDD | 優化後 MDD | 優化前 Calmar | 優化後 Calmar |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
"""
    for row in comparison_data:
        md_content += f"| {row['期間']} | {row['優化前 CAGR']} | {row['優化後 CAGR']} | {row['優化前 MDD']} | {row['優化後 MDD']} | {row['優化前 Calmar']} | {row['優化後 Calmar']} |\n"

    md_content += """
---

## 3. 結果分析
透過加入市場寬度濾網，策略在市場系統性風險較高（寬度低於 35%）的時期能及時空倉避險。

---

## 4. 相關檔案
- `equityV-breadth.xlsx`：詳細回測數據與優化對比。
- `backtest_breadth.py`：包含市場寬度邏輯的回測引擎。
- `run_equityV_breadth.py`：產出此報告的執行程式碼。
"""
    with open(OUTPUT_MD, 'w', encoding='utf-8') as f:
        f.write(md_content)

    print(f"Successfully generated {OUTPUT_EXCEL} and {OUTPUT_MD}")

if __name__ == "__main__":
    main()
