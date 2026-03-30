import pandas as pd
import numpy as np
import pickle
from backtest_v2 import clean_data, BacktesterV2, calculate_metrics

class ACO_Optimizer_Fixed:
    def __init__(self, backtester, n_ants=20, n_iterations=10, rho=0.1, alpha=1):
        self.bt = backtester
        self.n_ants = n_ants
        self.n_iterations = n_iterations
        self.rho = rho
        self.alpha = alpha

        # Search space expanded to include the robust area discovered (SMA up to 400)
        self.sma_range = np.arange(10, 401, 5)
        self.roc_range = np.arange(1, 151, 5)
        self.sl_range = np.arange(0.01, 0.10, 0.01)
        self.reb_range = np.arange(5, 21, 2)
        self.sl_type_range = ['peak', 'ma']
        self.ma_stop_range = [5, 10, 20, 60]

        self.sma_pheromones = np.ones(len(self.sma_range))
        self.roc_pheromones = np.ones(len(self.roc_range))
        self.sl_pheromones = np.ones(len(self.sl_range))
        self.reb_pheromones = np.ones(len(self.reb_range))
        self.sl_type_pheromones = np.ones(len(self.sl_type_range))
        self.ma_stop_pheromones = np.ones(len(self.ma_stop_range))

        self.best_params = None
        self.best_score = -np.inf

    def _select_param(self, range_vals, pheromones):
        probs = pheromones ** self.alpha
        probs /= probs.sum()
        return np.random.choice(range_vals, p=probs)

    def optimize(self, p1_s, p1_e, p2_s, p2_e):
        for gen in range(self.n_iterations):
            ants_results = []
            for ant in range(self.n_ants):
                sma = self._select_param(self.sma_range, self.sma_pheromones)
                roc = self._select_param(self.roc_range, self.roc_pheromones)
                sl = self._select_param(self.sl_range, self.sl_pheromones)
                reb = self._select_param(self.reb_range, self.reb_pheromones)
                sl_type = self._select_param(self.sl_type_range, self.sl_type_pheromones)
                ma_stop = self._select_param(self.ma_stop_range, self.ma_stop_pheromones)

                eq, trades, _, _, _ = self.bt.run(int(sma), int(roc), float(sl), int(reb), sl_type, int(ma_stop))

                res1 = calculate_metrics(eq[(eq['日期'] >= p1_s) & (eq['日期'] <= p1_e)])
                res2 = calculate_metrics(eq[(eq['日期'] >= p2_s) & (eq['日期'] <= p2_e)])

                score = min(res1[2], res2[2])
                if res1[1] < -0.25 or res2[1] < -0.25: score = 0
                if len(trades) < 20: score = 0

                ants_results.append(((sma, roc, sl, reb, sl_type, ma_stop), score))
                if score > self.best_score:
                    self.best_score = score
                    self.best_params = (sma, roc, sl, reb, sl_type, ma_stop)

            # Update pheromones
            self.sma_pheromones *= (1 - self.rho)
            self.roc_pheromones *= (1 - self.rho)
            for (sma, roc, sl, reb, sl_type, ma_stop), score in ants_results:
                if score > 0:
                    self.sma_pheromones[np.where(self.sma_range == sma)[0][0]] += score
                    self.roc_pheromones[np.where(self.roc_range == roc)[0][0]] += score

if __name__ == "__main__":
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterV2(prices, volumes, code_to_name)
    optimizer = ACO_Optimizer_Fixed(bt)
    print("ACO Optimizer with correct range prepared.")
