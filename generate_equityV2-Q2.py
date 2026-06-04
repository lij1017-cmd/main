import pandas as pd
import numpy as np
import nbformat as nbf
from backtest_equityV2 import clean_data, BacktesterV2, calculate_metrics

# --- 使用者設定區塊 (每日/季更新) ---
LAST_DATE = "2026-06-03"        # 每日更新最後一天日期
OUTPUT_SUFFIX = "260603"        # 每日更新檔名後綴

# 每季新增標的 (集中管理)
NEW_STOCKS_20260401 = ["3481群創", "6446藥華藥", "2368金像電",
                       "2334華邦電", "3037欣興", "2449京元電", "7769鴻勁"]

# 臨時指定回測期間 (可選) → 若不需要，保持 None
START_DATE = None
END_DATE = None

def main():
    CLEAN_DATA = "資料26Q2-1.xlsx"

    # 策略參數
    SMA_PERIOD = 303
    ROC_PERIOD = 14
    STOP_LOSS_PCT = 0.0999
    REBALANCE = 9
    INITIAL_CAPITAL = 30000000

    print(f"正在從 {CLEAN_DATA} 讀取資料並進行基礎預處理...")
    prices, volumes, code_to_name = clean_data(CLEAN_DATA)

    # --- 樣本群切換邏輯 ---
    mask_pre = (prices.index <= "2026-03-31")
    mask_post = (prices.index >= "2026-04-01")

    prices_pre = prices.loc[mask_pre].drop(columns=NEW_STOCKS_20260401, errors="ignore")
    volumes_pre = volumes.loc[mask_pre].drop(columns=NEW_STOCKS_20260401, errors="ignore")

    prices_post = prices.loc[mask_post]
    volumes_post = volumes.loc[mask_post]

    prices = pd.concat([prices_pre, prices_post])
    volumes = pd.concat([volumes_pre, volumes_post])

    print(f"開始執行回測 (SMA={SMA_PERIOD}, ROC={ROC_PERIOD}, 停損={STOP_LOSS_PCT*100:.2f}%, 再平衡={REBALANCE}天)...")
    bt = BacktesterV2(prices, volumes, code_to_name, initial_capital=INITIAL_CAPITAL)
    eq_df, trades, hold, trades2, daily = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE, "peak", 10)

    # --- 篩選回測期間 ---
    if START_DATE and END_DATE:
        mask_record = (eq_df["日期"] >= START_DATE) & (eq_df["日期"] <= END_DATE)
    else:
        mask_record = (eq_df["日期"] >= "2026-01-02") & (eq_df["日期"] <= LAST_DATE)

    res_record = eq_df[mask_record]
    trades_rec = trades[(trades["訊號日期"] >= res_record["日期"].min()) & (trades["訊號日期"] <= res_record["日期"].max())]
    hold_rec = hold[(hold["Date"] >= res_record["日期"].min()) & (hold["Date"] <= res_record["日期"].max())]
    trades2_rec = trades2[(trades2["賣出訊號日期"] >= res_record["日期"].min()) & (trades2["賣出訊號日期"] <= res_record["日期"].max())]
    daily_rec = daily[(daily["日期"] >= res_record["日期"].min()) & (daily["日期"] <= res_record["日期"].max())]

    # --- 計算績效指標 (保留 datetime 格式) ---
    cagr_rec, mdd_rec, calmar_rec, total_ret_rec = calculate_metrics(res_record)
    trade_count_rec = len(trades_rec[trades_rec["狀態"].isin(["買進", "賣出"])])

    # --- 日期欄位轉換成簡短格式 (僅用於輸出) ---
    for df, col in [(res_record, "日期"), (hold_rec, "Date"),
                    (trades_rec, "訊號日期"), (trades2_rec, "賣出訊號日期"),
                    (daily_rec, "日期")]:
        df[col] = pd.to_datetime(df[col]).dt.strftime("%Y/%m/%d")

    # --- 產出 Excel 報表 ---
    OUTPUT_EXCEL = f"trendstrategy_results_equityV2-{OUTPUT_SUFFIX}.xlsx"
    print(f"正在產出 Excel 報表：{OUTPUT_EXCEL}...")
    with pd.ExcelWriter(OUTPUT_EXCEL, engine="xlsxwriter") as writer:
        pd.DataFrame([
            {"項目": "記錄起始日期", "數值": res_record["日期"].min()},
            {"項目": "年化報酬率 (CAGR)", "數值": f"{cagr_rec:.2%}"},
            {"項目": "最大回撤 (MaxDD)", "數值": f"{mdd_rec:.2%}"},
            {"項目": "Calmar Ratio", "數值": f"{calmar_rec:.2f}"},
            {"項目": "總報酬率", "數值": f"{total_ret_rec:.2%}"},
            {"項目": "交易筆數", "數值": trade_count_rec}
        ]).to_excel(writer, sheet_name="Summary", index=False)

        res_record.to_excel(writer, sheet_name="Equity_Curve", index=False)
        hold_rec.to_excel(writer, sheet_name="Equity_Hold", index=False)
        trades_rec.to_excel(writer, sheet_name="Trades", index=False)
        trades2_rec.to_excel(writer, sheet_name="Trades2", index=False)
        daily_rec.to_excel(writer, sheet_name="Daily", index=False)

    # --- 產出 Markdown 報告 (自動更新最後一天日期) ---
    OUTPUT_MD = "reproduce_equityV2.md"
    md_content = f"""# Asset Class Trend Following 策略回測報告 (equityV2)

## 1. 策略說明
- **參數配置**：
    * SMA 週期：{SMA_PERIOD}
    * ROC 週期：{ROC_PERIOD}
    * 停損機制：波段最高點回落 {STOP_LOSS_PCT*100:.2f}%
    * 再平衡頻率：每 {REBALANCE} 個交易日
- **選股池異動**：
    * 2026/04/01 之前：選股池規模為 131 檔。
    * 2026/04/01 起：正式納入 7 檔新標的，選股池擴大至 138 檔。
- **最後一日訊號**：{LAST_DATE} 已產出完整再平衡指令

## 2. 績效表現摘要 (區間：{res_record['日期'].min()} – {res_record['日期'].max()})
- 年化報酬率 (CAGR)：{cagr_rec:.2%}
- 最大回撤 (MaxDD)：{mdd_rec:.2%}
- 卡瑪比率 (Calmar Ratio)：{calmar_rec:.2f}
- 總報酬率：{total_ret_rec:.2%}
- 交易筆數：{trade_count_rec}
"""
    with open(OUTPUT_MD, "w", encoding="utf-8") as f:
        f.write(md_content)

    print("✅ Excel、Markdown 均已完成！")

if __name__ == "__main__":
    main()
