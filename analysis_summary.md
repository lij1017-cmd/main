# B303 策略回測差異分析報告

## 1. 核心問題說明
針對 `trendstrategy_results_equityV-adj4-1-251231.xlsx` (以下簡稱 Excel) 與 `report_equityV-adj4.md` (以下簡稱 MD 報告) 數據不一致的問題，經交叉驗證診斷，主因可歸納為以下兩點：

### A. 資料來源檔案差異 (Data Source Discrepancy)
- **MD 報告基準值**：使用的是 `樣本集-1.xlsx`。在此資料集下，新引擎的 Baseline CAGR 為 **32.69%**，MaxDD 為 **-11.22%**。
- **Excel 產出**：使用的是 `資料26Q2-1.xlsx`。在此資料集下，同樣的邏輯產出的 CAGR 為 **32.83%**，MaxDD 為 **-12.63%**。
- **結論**：MaxDD 的顯著落差 (-11.22% vs -12.63%) 主要是因為 `資料26Q2-1.xlsx` 包含了不同的價格波動或除權息修正數據。

### B. 引擎校正效應 (Engine Calibration Effect)
- **舊版引擎 (`BacktesterV2`)**：在原始紀錄中，CAGR 為 **31.98%**，MaxDD 為 **-10.94%**。
- **新版引擎 (`BacktesterVol`)**：在相同資料 (`樣本集-1.xlsx`) 下，校正後的 CAGR 提升至 **32.69%** (+0.71%)。
- **差異原因**：新引擎優化了「槽位預算回收邏輯」(Slot Budget Recycling) 與 T+1 執行細節，減少了資金閒置產生的「現金拖累」(Cash Drag)，因此在多頭期間績效更強，但也因持倉更飽和而使回撤略微放大。

---

## 2. 數據對照表 (2019-2025)

| 指標 | 原始歷史紀錄 (Legacy) | MD 報告基準 (New Engine) | Excel 產出結果 (New Engine) |
| :--- | :--- | :--- | :--- |
| **使用資料檔** | 樣本集-1.xlsx | **樣本集-1.xlsx** | **資料26Q2-1.xlsx** |
| **CAGR (30M)** | 31.98% | 32.69% | 32.83% |
| **MaxDD** | -10.94% | -11.22% | -12.63% |
| **Calmar Ratio** | 2.92 | 2.91 | 2.60 |

---

## 3. 邏輯正確性驗證
1. **市場寬度邏輯**：經代碼檢查，Excel 採用的 `np.nanmean` 邏輯能正確排除未上市股票，與 MD 報告邏輯一致。
2. **停損與濾網**：Excel 目前產出的版本為 **Baseline (固定停損 9.99% + 無濾網)**。若要達到 MD 報告中「方案 C」的優異指標 (CAGR 33.66% / MaxDD -10.28%)，需手動於 `backtest_adj4(舊).py` 開啟 `USE_MARKET_FILTER` 與 `STOP_LOSS_TYPE = 'vol'`。
3. **雙資本指標**：Excel 已成功整合 30M 實戰資金與 150M 授權資金的雙軌計算，符合最新報表標準。

## 4. 建議行動
若需產生與 MD 報告完全一致的交易明細，建議：
1. 將 `backtest_adj4(舊).py` 的資料來源改回 `樣本集-1.xlsx`。
2. 根據需求切換 `USE_MARKET_FILTER` (True) 與 `STOP_LOSS_TYPE` ('vol') 以對應「方案 C」的最優績效版本。
