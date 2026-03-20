import pandas as pd
import numpy as np
from run_wfa import Backtester, clean_data, calculate_metrics

class ACO_Optimizer_2022_2025:
    def __init__(self, prices, code_to_name, n_ants=50, n_iterations=30, rho=0.1, alpha=1):
        self.prices = prices
        self.code_to_name = code_to_name
        self.n_ants = n_ants
        self.n_iterations = n_iterations
        self.rho = rho
        self.alpha = alpha

        # 搜尋空間 (受限版：SMA <= 90, ROC <= 90, SL <= 10%)
        self.sma_range = np.arange(10, 91, 1)
        self.roc_range = np.arange(10, 91, 1)
        self.sl_range = np.arange(0.05, 0.101, 0.005)

        # 初始化信息素
        self.sma_pheromones = np.ones(len(self.sma_range))
        self.roc_pheromones = np.ones(len(self.roc_range))
        self.sl_pheromones = np.ones(len(self.sl_range))

        # 初始偏好 21, 37, 0.095 (之前找到過 > 2 的一組)
        # self.sma_pheromones[np.where(self.sma_range == 21)[0][0]] = 2.0
        # self.roc_pheromones[np.where(self.roc_range == 37)[0][0]] = 2.0
        # self.sl_pheromones[np.where(self.sl_range == 0.095)[0][0]] = 2.0

        self.best_params = None
        self.best_score = -np.inf

    def _select_param(self, range_vals, pheromones):
        probs = pheromones ** self.alpha
        probs /= probs.sum()
        return np.random.choice(range_vals, p=probs)

    def optimize(self, start_date, end_date):
        bt = Backtester(self.prices, self.code_to_name)

        for gen in range(self.n_iterations):
            ants_results = []
            for _ in range(self.n_ants):
                sma = self._select_param(self.sma_range, self.sma_pheromones)
                roc = self._select_param(self.roc_range, self.roc_pheromones)
                sl = self._select_param(self.sl_range, self.sl_pheromones)

                # 使用 6 天再平衡，依照 2022-2025 期間的最佳化需求
                eq, trades = bt.run(int(sma), int(roc), float(sl), 6, start_date, end_date)
                mask = (eq['日期'] >= pd.to_datetime(start_date)) & (eq['日期'] <= pd.to_datetime(end_date))
                eq_period = eq[mask]

                if eq_period.empty or len(eq_period) < 2:
                    score = -1
                else:
                    cagr, mdd, calmar = calculate_metrics(eq_period)
                    # 目標：最大化 Calmar Ratio，同時給予一定的 MDD 罰則避免極端
                    # 給予 Calmar > 2 的極高權重
                    score = calmar if mdd < -0.05 and trades > 20 else -1
                    if calmar > 2: score *= 2

                ants_results.append(((sma, roc, sl), score))

                if score > self.best_score:
                    self.best_score = score
                    self.best_params = (sma, roc, sl)

            # 更新信息素
            self.sma_pheromones *= (1 - self.rho)
            self.roc_pheromones *= (1 - self.rho)
            self.sl_pheromones *= (1 - self.rho)

            for (sma, roc, sl), score in ants_results:
                if score > 0:
                    self.sma_pheromones[np.where(self.sma_range == sma)[0][0]] += score
                    self.roc_pheromones[np.where(self.roc_range == roc)[0][0]] += score
                    self.sl_pheromones[np.where(self.sl_range == sl)[0][0]] += score

            print(f"Gen {gen+1}: Best Calmar {self.best_score:.4f} (SMA={self.best_params[0]}, ROC={self.best_params[1]}, SL={self.best_params[2]:.3f})")

        return self.best_params, self.best_score

if __name__ == "__main__":
    prices, code_to_name = clean_data('個股合-1.xlsx')
    optimizer = ACO_Optimizer_2022_2025(prices, code_to_name)

    start_date = '2022-01-02'
    end_date = '2025-12-31'

    print(f"開始為 {start_date} 至 {end_date} 進行參數最佳化...")
    best_p, best_s = optimizer.optimize(start_date, end_date)

    # 計算最佳參數下的最終指標
    bt = Backtester(prices, code_to_name)
    eq, trades = bt.run(int(best_p[0]), int(best_p[1]), float(best_p[2]), 6, start_date, end_date)
    cagr, mdd, calmar = calculate_metrics(eq)

    print("\n=== 最佳化結果 (2022-2025) ===")
    print(f"SMA: {best_p[0]}")
    print(f"ROC: {best_p[1]}")
    print(f"Stop Loss: {best_p[2]*100:.1f}%")
    print(f"CAGR: {cagr:.2%}")
    print(f"MaxDD: {mdd:.2%}")
    print(f"Calmar Ratio: {calmar:.2f}")
    print(f"交易次數: {trades}")
