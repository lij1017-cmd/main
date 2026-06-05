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

    # --- Warm-Start 設定 (承接 2025/12/31 狀態) ---
    WARM_START_ENABLED = True
    WARM_START_CASH = 127925651.0
    WARM_START_SLOTS = {
        0: {'code': '3211', 'shares': 29000, 'entry_price': 342.50, 'max_price': 337.00,
            'budget': 9946653.81, 'entry_date': '2025-12-29'},
        1: {'code': '3152', 'shares': 66000, 'entry_price': 149.50, 'max_price': 147.00,
            'budget': 9881060.48, 'entry_date': '2025-12-29'},
        2: {'code': '3260', 'shares': 39000, 'entry_price': 252.06, 'max_price': 270.45,
            'budget': 9844348.23, 'entry_date': '2025-12-29'},
    }

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
    bt = BacktesterV2(prices, volumes, code_to_name,
                      initial_capital=INITIAL_CAPITAL,
                      warm_start_slots=WARM_START_SLOTS if WARM_START_ENABLED else None,
                      warm_start_cash=WARM_START_CASH if WARM_START_ENABLED else None)

    eq_df, trades, hold, trades2, daily = bt.run(SMA_PERIOD, ROC_PERIOD, STOP_LOSS_PCT, REBALANCE, "peak", 10)

    # --- 篩選回測期間 ---
    mask_record = (eq_df["日期"] >= "2026-01-02") & (eq_df["日期"] <= LAST_DATE)
    res_record = eq_df[mask_record]

    # 年度績效歸零計算
    eq_df["年度基準"] = eq_df.groupby(eq_df["日期"].str[:4])["權益"].transform("first")
    eq_df["年度報酬率"] = (eq_df["權益"] - eq_df["年度基準"]) / eq_df["年度基準"]

    # 後續 Excel / Markdown 輸出保持不變
