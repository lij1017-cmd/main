import pandas as pd
import numpy as np
import nbformat as nbf
from backtest_equityV2 import clean_data, BacktesterV2, calculate_metrics

def main():
    """
    主程式：執行 equityV2 策略回測並產出各類績效明細。
    本腳本包含詳盡的繁體中文註解，確保每一步邏輯透明清晰。
    """

    # --- 第一步：設定資料路徑與載入清洗後的資料 ---
    # 使用已經過前處理(補值、除錯)的 26Q2 資料集，確保數據品質
    CLEAN_DATA = '資料26Q2-1.xlsx'

    # --- 第二步：設定策略核心參數 ---
    # 這些參數沿用自 equityV 版本，確保回測邏輯的延續性與比較基準
    # SMA_PERIOD 303：選用較長的移動平均週期作為市場趨勢的主濾網
    # ROC_PERIOD 14：計算 14 日的價格變動率，捕捉中短期動能
    # STOP_LOSS_PCT 9.99%：當價格自波段高點回落約 10% 時觸發停損
    # REBALANCE 9：設定每 9 個交易日執行一次投資組合再平衡
    # INITIAL_CAPITAL：設定初始投資資金為 3,000 萬 TWD
    SMA_PERIOD = 303
    ROC_PERIOD = 14
    STOP_LOSS_PCT = 0.0999
    REBALANCE = 9
    INITIAL_CAPITAL = 30000000

    print(f"正在從 {CLEAN_DATA} 讀取資料並進行基礎預處理...")
    # 呼叫 backtest_equityV2 中的 clean_data 函數，提取價格、成交量與股票名稱對應表
    prices, volumes, code_to_name = clean_data(CLEAN_DATA)

    print(f"開始執行 equityV2 策略回測 (參數設定：SMA={SMA_PERIOD}, ROC={ROC_PERIOD}, 停損={STOP_LOSS_PCT*100:.2f}%, 再平衡={REBALANCE}天)...")
    # 初始化回測引擎實例
    bt = BacktesterV2(prices, volumes, code_to_name, initial_capital=INITIAL_CAPITAL)

    # 執行完整期間回測
    # 核心邏輯：在 2026/04/01 之前選股池為 131 檔；2026/04/01 之後自動擴展至 138 檔
    # 此版本新增了「備註」功能，會紀錄停損與再平衡指令
    eq_df, trades, hold, trades2, daily = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE, 'peak', 10)

    # --- 第三步：數據篩選與績效指標計算 ---
    # 依據使用者需求，績效明細從 2026/01/02 開始記錄，以便與先前成果進行對接驗證
    mask_record = (eq_df['日期'] >= '2026-01-02')
    res_record = eq_df[mask_record]

    # 計算記錄區間內的年化報酬率 (CAGR)、最大回撤 (MaxDD) 與卡瑪比率 (Calmar Ratio)
    cagr_rec, mdd_rec, calmar_rec, total_ret_rec = calculate_metrics(res_record)

    # 篩選並統計 2026/01/02 之後發生的實際買賣交易筆數
    trades_rec = trades[(trades['訊號日期'] >= '2026-01-02')]
    trade_count_rec = len(trades_rec[trades_rec['狀態'].isin(['買進', '賣出'])])

    # --- 第四步：產出 Excel 績效明細報表 ---
    OUTPUT_EXCEL = 'trendstrategy_results_equityV2.xlsx'
    print(f"正在產出 Excel 報表：{OUTPUT_EXCEL}...")
    with pd.ExcelWriter(OUTPUT_EXCEL, engine='xlsxwriter') as writer:
        # 1. Summary：呈現本階段回測的核心績效統計
        pd.DataFrame([
            {'項目': '記錄起始日期', '數值': '2026-01-02'},
            {'項目': '年化報酬率 (CAGR)', '數值': f"{cagr_rec:.2%}"},
            {'項目': '最大回撤 (MaxDD)', '數值': f"{mdd_rec:.2%}"},
            {'項目': 'Calmar Ratio', '數值': f"{calmar_rec:.2f}"},
            {'項目': '總報酬率', '數值': f"{total_ret_rec:.2%}"},
            {'項目': '2026起交易筆數 (不含保持)', '數值': trade_count_rec}
        ]).to_excel(writer, sheet_name='Summary', index=False)

        # 2. Equity_Curve：記錄每日帳戶淨值與回撤變動
        res_record.to_excel(writer, sheet_name='Equity_Curve', index=False)

        # 3. Equity_Hold：記錄每日持股摘要、可用現金及重要操作「備註」
        hold[hold['Date'] >= '2026-01-02'].to_excel(writer, sheet_name='Equity_Hold', index=False)

        # 4. Trades：詳細列出所有訊號 (買入、賣出、持倉保持) 的詳細資訊
        trades_rec.to_excel(writer, sheet_name='Trades', index=False)

        # 5. Trades2：成對交易紀錄，展示每筆進出場的關聯、損益與原因
        trades2[trades2['賣出訊號日期'] >= '2026-01-02'].to_excel(writer, sheet_name='Trades2', index=False)

        # 6. Daily：逐日記錄每檔持股的股數、收盤價與市值，便於稽核
        daily[daily['日期'] >= '2026-01-02'].to_excel(writer, sheet_name='Daily', index=False)

        # 在 Equity_Curve 工作表中自動插入動態折線圖，直觀呈現淨值走勢
        workbook = writer.book
        curves_sheet = writer.sheets['Equity_Curve']
        chart = workbook.add_chart({'type': 'line'})
        max_row = len(res_record)
        chart.add_series({
            'name': '權益曲線 (equityV2)',
            'categories': ['Equity_Curve', 1, 0, max_row, 0],
            'values': ['Equity_Curve', 1, 1, max_row, 1],
        })
        chart.set_title({'name': '2026 淨值走勢圖 (起始日: 2026/01/02)'})
        curves_sheet.insert_chart('E2', chart)

    # --- 第五步：產出 Markdown 回測總結報告 ---
    OUTPUT_MD = 'reproduce_equityV2.md'
    md_content = f"""# Asset Class Trend Following 策略回測報告 (equityV2)

## 1. 策略說明
本報告針對 **equityV2** 版本進行總結，該版本在維持原有核心邏輯的基礎上，新增了 2026 Q2 的標的擴充規則，並優化了執行指令備註。

- **參數配置**：
    * **SMA 週期**：{SMA_PERIOD}
    * **ROC 週期**：{ROC_PERIOD}
    * **停損機制**：波段最高點回落 {STOP_LOSS_PCT*100:.2f}%
    * **再平衡頻率**：每 {REBALANCE} 個交易日
- **選股池異動**：
    * **2026/04/01 之前**：選股池規模為 131 檔。
    * **2026/04/01 起**：正式納入 7 檔新標的，選股池擴大至 138 檔。
- **功能優化**：
    * **Equity_Hold 備註欄**：新增停損提醒與再平衡交易指令。
    * **最後一日訊號**：2026/04/10 已產出完整再平衡指令，供次日執行參考。

---

## 2. 績效表現摘要 (區間：2026.01.02 – 至今)
- **年化報酬率 (CAGR)**：**{cagr_rec:.2%}**
- **最大回撤 (MaxDD)**：**{mdd_rec:.2%}**
- **卡瑪比率 (Calmar Ratio)**：**{calmar_rec:.2f}**
- **總報酬率**：**{total_ret_rec:.2%}**
- **2026起交易筆數**：**{trade_count_rec}**

---

## 3. 交付檔案說明
- `trendstrategy_results_equityV2.xlsx`：包含完整交易、持股明細與「備註」指令的 Excel 檔案。
- `trendstrategy_equityV2.ipynb`：可於 Jupyter 環境執行的互動式回測實作。
- `backtest_equityV2.py`：封裝好的回測引擎模組。
"""
    with open(OUTPUT_MD, 'w', encoding='utf-8') as f:
        f.write(md_content)

    # --- 第六步：產出 Jupyter Notebook 實作檔 (.ipynb) ---
    nb = nbf.v4.new_notebook()
    nb.cells.append(nbf.v4.new_markdown_cell(f"# Asset Class Trend Following 策略實作 (equityV2)\\n\\n本筆記本詳細記錄了從 **2026-01-02** 開始的策略執行邏輯，包含 **備註指令欄位** 以利實務執行。"))

    code_block = f"""import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from backtest_equityV2 import clean_data, BacktesterV2, calculate_metrics

# --- 1. 資料讀取與載入 ---
prices, volumes, code_to_name = clean_data('{CLEAN_DATA}')

# --- 2. 策略參數定義 ---
sma_p, roc_p, sl_p, reb_p = {SMA_PERIOD}, {ROC_PERIOD}, {STOP_LOSS_PCT}, {REBALANCE}

# --- 3. 執行回測運算 ---
# BacktesterV2 已內建「備註」欄位產生邏輯與動態選股池規則。
bt = BacktesterV2(prices, volumes, code_to_name)
eq, trades, hold, trades2, daily = bt.run(sma_p, roc_p, sl_p, reb_p, 'peak', 10)

# --- 4. 數據篩選 (2026/01/02 起) ---
mask = (eq['日期'] >= '2026-01-02')
res_p = eq[mask]

# --- 5. 視覺化：權益曲線 ---
plt.figure(figsize=(12, 6))
plt.plot(res_p['日期'], res_p['權益'], color='darkgreen', linewidth=2)
plt.title('Equity Curve (equityV2 - With Remarks)', fontsize=14)
plt.grid(True, linestyle='--', alpha=0.5)
plt.show()

# --- 6. 最後一日 (2026/04/10) 執行指令查看 ---
last_day_remark = hold.iloc[-1]['備註']
print(f"2026/04/10 操作備註：\\n{{last_day_remark}}")

# --- 7. 績效統計 ---
cagr, mdd, calmar, total_ret = calculate_metrics(res_p)
print(f"\\n年化報酬率 (CAGR): {{cagr:.2%}}")
print(f"最大回撤 (MaxDD): {{mdd:.2%}}")
print(f"卡瑪比率 (Calmar Ratio): {{calmar:.2f}}")
"""
    nb.cells.append(nbf.v4.new_code_cell(code_block))

    with open('trendstrategy_equityV2.ipynb', 'w', encoding='utf-8') as f:
        nbf.write(nb, f)

    print(f"✅ 交付文件產出完成：Excel (含備註欄), Markdown, Notebook 均已就緒。")

if __name__ == "__main__":
    main()
