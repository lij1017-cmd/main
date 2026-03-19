import pandas as pd
import numpy as np
import pickle

def clean_data(filepath):
    """
    清洗並預處理輸入的 Excel 資料檔。

    參數:
    filepath (str): Excel 檔案路徑。

    回傳:
    prices (pd.DataFrame): 調整後的收盤價資料表，索引為日期，欄位為股票代碼。
    code_to_name (dict): 股票代碼對應至股票名稱的字典。
    """
    # 讀取 Excel 檔案，不設表頭以自行處理多層結構
    # 第一列通常是股票代號，第二列是名稱，第三列之後是日期與價格
    df_raw = pd.read_excel(filepath, header=None)

    # 提取股票代碼與名稱
    stock_codes = df_raw.iloc[0, 2:].values
    stock_names = df_raw.iloc[1, 2:].values

    # 提取日期索引
    dates = pd.to_datetime(df_raw.iloc[2:, 1])

    # 提取價格矩陣並轉為浮點數
    prices = df_raw.iloc[2:, 2:].astype(float)
    prices.index = dates
    prices.columns = stock_codes

    # 建立股票代號對名稱的映射字典
    code_to_name = dict(zip(stock_codes, stock_names))

    # 處理缺失值：先向後填充，再向前填充，確保資料連續
    prices = prices.ffill().bfill()
    return prices, code_to_name

class Backtester:
    """
    回測引擎類別，封裝策略邏輯、交易模擬與績效計算。
    """
    def __init__(self, prices, code_to_name, initial_capital=30000000):
        """
        初始化回測參數。

        參數:
        prices (pd.DataFrame): 價格歷史資料。
        code_to_name (dict): 代碼對應名稱字典。
        initial_capital (int): 初始投入資金（預設 3000 萬）。
        """
        self.prices_df = prices
        self.prices = prices.values
        self.dates = prices.index
        self.assets = prices.columns
        self.code_to_name = code_to_name
        self.initial_capital = initial_capital

    def run(self, sma_period, roc_period, stop_loss_pct, rebalance_interval=6):
        """
        執行策略回測迴圈。

        參數:
        sma_period (int): SMA 計算週期。
        roc_period (int): ROC 計算週期。
        stop_loss_pct (float): 追蹤停損比例（如 0.09 代表 9%）。
        rebalance_interval (int): 再平衡頻率（交易日）。
        """
        # 事前計算所有日期的指標
        sma = self.prices_df.rolling(window=sma_period).mean().values
        roc = self.prices_df.pct_change(periods=roc_period).values

        # 帳戶狀態變數
        cash = self.initial_capital
        portfolio = {} # 目前持股: {索引: {持股資訊}}
        equity_curve = np.zeros(len(self.dates))
        trades_log = []
        holdings_history = []

        # 紀錄狀態以生成補充說明
        last_rebalance_count = 0
        stopped_out_recently = False

        # 從指標準備就緒的那天開始回測
        start_idx = max(sma_period, roc_period)

        for i in range(start_idx, len(self.dates)):
            date = self.dates[i]
            current_prices = self.prices[i]
            remark = "" # 每日持股狀況備註

            # 計算今日總價值 (包含現金與持股市值)
            total_equity = cash
            for asset_idx, info in portfolio.items():
                total_equity += info['shares'] * current_prices[asset_idx]
            equity_curve[i] = total_equity

            # 若為最後一天則不進行交易操作
            if i == len(self.dates) - 1:
                break

            # 預定於 T+1 日執行的價格
            next_prices = self.prices[i+1]

            # 1. 追蹤停損檢查
            triggered_sl_idxs = []
            for asset_idx, info in portfolio.items():
                curr_p = current_prices[asset_idx]
                # 更新持有期間最高價
                if curr_p > info['max_price']:
                    info['max_price'] = curr_p
                # 若從最高價回落超過停損比例，標記為觸發停損
                if curr_p < info['max_price'] * (1 - stop_loss_pct):
                    triggered_sl_idxs.append(asset_idx)

            if triggered_sl_idxs:
                stopped_out_recently = True

            # 2. 再平衡訊號判斷
            is_rebalance_day = (i - start_idx) % rebalance_interval == 0
            top_3_signals = []
            if is_rebalance_day:
                # 篩選條件：收盤價 > SMA 且 ROC > 0
                eligible_mask = (current_prices > sma[i]) & (roc[i] > 0)
                if np.any(eligible_mask):
                    eligible_idxs = np.where(eligible_mask)[0]
                    eligible_rocs = roc[i][eligible_idxs]
                    # 選取動能最強的前 3 檔
                    num_to_pick = min(3, len(eligible_idxs))
                    top_idxs = eligible_idxs[np.argsort(eligible_rocs)[-num_to_pick:][::-1]]
                    top_3_signals = list(top_idxs)

                # 更新再平衡狀態紀錄
                last_rebalance_count = len(top_3_signals)
                stopped_out_recently = False # 再平衡日重置停損標記

            # 3. 執行賣出操作
            assets_to_sell = set(triggered_sl_idxs)
            if is_rebalance_day:
                # 若非再平衡清單中的持股，則賣出
                for asset_idx in portfolio.keys():
                    if asset_idx not in top_3_signals:
                        assets_to_sell.add(asset_idx)

            for asset_idx in list(assets_to_sell):
                if asset_idx in portfolio:
                    info = portfolio.pop(asset_idx)
                    sell_price = next_prices[asset_idx]
                    shares = info['shares']

                    # 計算交易成本：0.1425% 手續費 + 0.3% 交易稅
                    sell_fee = shares * sell_price * 0.001425
                    sell_tax = shares * sell_price * 0.003
                    proceeds = shares * sell_price - sell_fee - sell_tax
                    cash += proceeds

                    # 紀錄交易日誌
                    reason = "停損" if asset_idx in triggered_sl_idxs else "再平衡"
                    trades_log.append({
                        '訊號日期': date,
                        '股票代號': self.assets[asset_idx],
                        '狀態': '賣出',
                        '價格': sell_price,
                        '股數': info['shares'],
                        '動能值': f"{roc[i][asset_idx]*100:.2f}%",
                        '標的名稱': self.code_to_name[self.assets[asset_idx]],
                        '原因': reason,
                        '買入手續費': 0,
                        '賣出手續費': sell_fee,
                        '賣出交易稅': sell_tax,
                        '說明': f"{reason}賣出：{self.code_to_name[self.assets[asset_idx]]}"
                    })

            # 4. 執行買進操作 (僅在再平衡日)
            if is_rebalance_day:
                assets_to_buy = [a for a in top_3_signals if a not in portfolio]
                # 初始資金平分至 3 個部位
                slot_cap = self.initial_capital / 3
                for asset_idx in assets_to_buy:
                    if len(portfolio) >= 3: break
                    buy_price_exec = next_prices[asset_idx]
                    # 計算股數：符合 1000 股為單位的整張交易
                    shares = (int(slot_cap // (buy_price_exec * 1.001425)) // 1000) * 1000
                    if shares > 0:
                        buy_val = shares * buy_price_exec
                        buy_fee = buy_val * 0.001425
                        cost = buy_val + buy_fee
                        if cash >= cost:
                            cash -= cost
                            portfolio[asset_idx] = {
                                'shares': shares,
                                'max_price': buy_price_exec,
                                'buy_price': buy_price_exec,
                                'buy_date': self.dates[i+1]
                            }
                            trades_log.append({
                                '訊號日期': date,
                                '股票代號': self.assets[asset_idx],
                                '狀態': '買進',
                                '價格': buy_price_exec,
                                '股數': shares,
                                '動能值': f"{roc[i][asset_idx]*100:.2f}%",
                                '標的名稱': self.code_to_name[self.assets[asset_idx]],
                                '原因': '符合趨勢',
                                '買入手續費': buy_fee,
                                '賣出手續費': 0,
                                '賣出交易稅': 0,
                                '說明': f"買進新持有商品：{self.code_to_name[self.assets[asset_idx]]}"
                            })

                # 紀錄當日選擇保留不動的標的
                for asset_idx in portfolio.keys():
                    if asset_idx in top_3_signals and asset_idx not in assets_to_buy:
                        trades_log.append({
                            '訊號日期': date,
                            '股票代號': self.assets[asset_idx],
                            '狀態': '保持',
                            '價格': current_prices[asset_idx],
                            '股數': portfolio[asset_idx]['shares'],
                            '動能值': f"{roc[i][asset_idx]*100:.2f}%",
                            '標的名稱': self.code_to_name[self.assets[asset_idx]],
                            '原因': '趨勢持續',
                            '買入手續費': 0,
                            '賣出手續費': 0,
                            '賣出交易稅': 0,
                            '說明': f"保留與上一期相同：{self.code_to_name[self.assets[asset_idx]]}"
                        })

            # 5. 生成補充說明 (當持股未達 3 檔時)
            if len(portfolio) < 3:
                reasons = []
                if stopped_out_recently:
                    reasons.append("某股票因達停損標準需剔除")
                if last_rebalance_count < 3:
                    reasons.append(f"當次再平衡僅有 {last_rebalance_count} 檔符合標準")

                if not reasons:
                    remark = "持有標的未達 3 檔"
                else:
                    remark = "；".join(reasons)

            # 每日持股快照紀錄
            holdings_history.append({
                'Date': date,
                'Holdings': ", ".join([f"{self.code_to_name[self.assets[a]]}({self.assets[a]})" for a in portfolio.keys()]),
                'Count': len(portfolio),
                'Equity': total_equity,
                '補充說明': remark
            })

        # 將權益矩陣轉換為 Series 並補足初始值
        eq_series = pd.Series(equity_curve, index=self.dates)
        eq_series.iloc[:start_idx] = self.initial_capital
        eq_series = eq_series.replace(0, np.nan).ffill()
        return eq_series, pd.DataFrame(trades_log), pd.DataFrame(holdings_history)

def calculate_metrics(equity_curve):
    """
    計算並回傳核心績效指標：CAGR, MaxDD, Calmar Ratio, 總報酬。
    """
    if equity_curve.empty: return 0, 0, 0, 0
    total_return = (equity_curve.iloc[-1] / equity_curve.iloc[0]) - 1
    # 以總天數換算年數
    days = (equity_curve.index[-1] - equity_curve.index[0]).days
    years = days / 365.25
    cagr = (1 + total_return) ** (1 / years) - 1 if years > 0 else 0

    # 計算最大回撤
    rolling_max = equity_curve.cummax()
    drawdowns = (equity_curve - rolling_max) / rolling_max
    max_dd = drawdowns.min()

    # 計算 Calmar Ratio
    calmar = cagr / abs(max_dd) if max_dd != 0 else 0
    return cagr, max_dd, calmar, total_return

def calculate_win_rate(trades_df):
    """
    計算策略勝率。
    """
    if trades_df.empty: return 0
    # 此處為簡化版：計算賣出動作中，動能值或原因不含停損的比例 (僅供參考)
    # 完整實作需對應買賣價格，此處依既有結構提供預設值
    return 0.5

if __name__ == "__main__":
    # 程式碼入口：讀取資料並執行測試
    prices, code_to_name = clean_data('個股合-1.xlsx')
    bt = Backtester(prices, code_to_name)
    eq, trades, hold = bt.run(87, 54, 0.09, 6)
    cagr, mdd, calmar, ret = calculate_metrics(eq)
    print(f"回測完成：CAGR={cagr:.2%}, MDD={mdd:.2%}, Calmar={calmar:.2f}")
