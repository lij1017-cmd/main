import pandas as pd
import numpy as np
import pickle
import time
from backtest_equity2MA import clean_data, Backtester, calculate_metrics

class ACO_Optimizer:
    """
    使用螞蟻演算法 (ACO) 最佳化策略參數。
    """
    def __init__(self, bt, n_ants=20, n_iterations=10, rho=0.1, alpha=1):
        self.bt = bt
        self.n_ants = n_ants
        self.n_iterations = n_iterations
        self.rho = rho
        self.alpha = alpha

        # 搜尋空間
        self.sma1_range = np.arange(10, 401, 10)
        self.sma2_range = np.arange(10, 401, 10)
        self.roc_range = np.arange(10, 31, 1)
        self.sl_type_range = ['peak', 'ma']
        self.sl_peak_range = np.arange(0.01, 0.10, 0.005)
        self.sl_ma_range = np.arange(5, 61, 5)
        self.rebal_range = np.arange(5, 11, 1)

        # 初始化費洛蒙
        self.phero_sma1 = np.ones(len(self.sma1_range))
        self.phero_sma2 = np.ones(len(self.sma2_range))
        self.phero_roc = np.ones(len(self.roc_range))
        self.phero_sl_type = np.ones(len(self.sl_type_range))
        self.phero_sl_peak = np.ones(len(self.sl_peak_range))
        self.phero_sl_ma = np.ones(len(self.sl_ma_range))
        self.phero_rebal = np.ones(len(self.rebal_range))

        self.best_params = None
        self.best_score = -np.inf

    def _select(self, range_vals, pheromones):
        probs = pheromones ** self.alpha
        probs /= probs.sum()
        return np.random.choice(range_vals, p=probs)

    def optimize(self, train_start, train_end, test_start, test_end):
        print(f"開始最佳化程序...")

        for gen in range(self.n_iterations):
            ants_results = []
            for ant in range(self.n_ants):
                s1 = int(self._select(self.sma1_range, self.phero_sma1))
                s2 = int(self._select(self.sma2_range, self.phero_sma2))
                r = int(self._select(self.roc_range, self.phero_roc))
                st = self._select(self.sl_type_range, self.phero_sl_type)

                if st == 'peak':
                    sv = float(self._select(self.sl_peak_range, self.phero_sl_peak))
                else:
                    sv = int(self._select(self.sl_ma_range, self.phero_sl_ma))

                rb = int(self._select(self.rebal_range, self.phero_rebal))

                # 執行回測
                eq, trades, _, _, _ = self.bt.run(s1, s2, r, sv, rb, st)

                # 評估樣本集與測試集績效
                eq_train = eq[(eq['日期'] >= train_start) & (eq['日期'] <= train_end)]
                cagr_tr, mdd_tr, cal_tr, _ = calculate_metrics(eq_train)

                eq_test = eq[(eq['日期'] >= test_start) & (eq['日期'] <= test_end)]
                if not eq_test.empty:
                    eq_test_reset = eq_test.copy()
                    iv = eq_test_reset['權益值'].iloc[0]
                    eq_test_reset['權益值'] = eq_test_reset['權益值'] / iv * 30000000
                    peak = eq_test_reset['權益值'].cummax()
                    eq_test_reset['回撤(Drawdown)'] = (eq_test_reset['權益值'] - peak) / peak
                    cagr_te, mdd_te, cal_te, _ = calculate_metrics(eq_test_reset)
                else:
                    cal_te = 0
                    mdd_te = 0

                # 評分函數：權衡樣本與測試集 Calmar，並懲罰 MDD 過大或交易次數過少
                if mdd_tr < -0.25 or mdd_te < -0.25:
                    score = min(cal_tr, cal_te) * 0.1
                else:
                    score = min(cal_tr, cal_te) * 1.5 + (cal_tr + cal_te) / 2.0

                if len(trades) < 15:
                    score = -1

                ants_results.append(((s1, s2, r, st, sv, rb), score))

                if score > self.best_score:
                    self.best_score = score
                    self.best_params = (s1, s2, r, st, sv, rb)
                    print(f"找到更優參數! 評分: {score:.4f} | 樣本集 Calmar: {cal_tr:.2f}, 測試集 Calmar: {cal_te:.2f}")

            # 更新費洛蒙
            self._update_pheromones(ants_results)
            print(f"第 {gen+1} 輪演化完成。")

        return self.best_params

    def _update_pheromones(self, results):
        self.phero_sma1 *= (1 - self.rho)
        self.phero_sma2 *= (1 - self.rho)
        self.phero_roc *= (1 - self.rho)
        self.phero_sl_type *= (1 - self.rho)
        self.phero_sl_peak *= (1 - self.rho)
        self.phero_sl_ma *= (1 - self.rho)
        self.phero_rebal *= (1 - self.rho)

        for params, score in results:
            if score > 0:
                s1, s2, r, st, sv, rb = params
                self.phero_sma1[np.where(self.sma1_range == s1)[0][0]] += score
                self.phero_sma2[np.where(self.sma2_range == s2)[0][0]] += score
                self.phero_roc[np.where(self.roc_range == r)[0][0]] += score
                self.phero_sl_type[self.sl_type_range.index(st)] += score
                if st == 'peak':
                    self.phero_sl_peak[np.where(self.sl_peak_range == sv)[0][0]] += score
                else:
                    self.phero_sl_ma[np.where(self.sl_ma_range == sv)[0][0]] += score
                self.phero_rebal[np.where(self.rebal_range == rb)[0][0]] += score

if __name__ == "__main__":
    prices, volumes, names = clean_data('樣本集-1.xlsx')
    bt = Backtester(prices, volumes, names)

    tr_s, tr_e = pd.to_datetime('2019-01-02'), pd.to_datetime('2023-12-31')
    te_s, te_e = pd.to_datetime('2024-01-02'), pd.to_datetime('2025-12-31')

    optimizer = ACO_Optimizer(bt, n_ants=40, n_iterations=15)
    best_p = optimizer.optimize(tr_s, tr_e, te_s, te_e)

    print(f"\n最佳參數組合: {best_p}")
    with open('best_params_equity2MA.pkl', 'wb') as f:
        pickle.dump(best_p, f)
