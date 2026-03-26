import pandas as pd
import numpy as np
import pickle
from backtest_v2 import clean_data, BacktesterV2, calculate_metrics

class ACO_Optimizer_V6:
    def __init__(self, backtester, n_ants=30, n_iterations=30, rho=0.1, alpha=1):
        self.bt = backtester
        self.n_ants = n_ants
        self.n_iterations = n_iterations
        self.rho = rho
        self.alpha = alpha

        # Search space
        self.sma_range = np.arange(30, 121, 1)
        self.roc_range = np.arange(30, 121, 1)
        self.sl_range = np.arange(0.01, 0.101, 0.005)
        self.reb_range = np.arange(5, 11, 1)
        self.sl_type_range = ['peak', 'ma']
        self.ma_stop_range = [5, 10, 20]

        # Initialize pheromones
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

    def optimize(self, start_date, end_date):
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
                mask = (eq['日期'] >= start_date) & (eq['日期'] <= end_date)
                eq_period = eq[mask]

                if eq_period.empty or len(trades) < 5:
                    score = -1
                else:
                    cagr, mdd, calmar, _ = calculate_metrics(eq_period)
                    score = calmar
                    if mdd < -0.25: score *= 0.1

                ants_results.append(((sma, roc, sl, reb, sl_type, ma_stop), score))

                if score > self.best_score:
                    self.best_score = score
                    self.best_params = (sma, roc, sl, reb, sl_type, ma_stop)

            # Update pheromones
            self.sma_pheromones *= (1 - self.rho)
            self.roc_pheromones *= (1 - self.rho)
            self.sl_pheromones *= (1 - self.rho)
            self.reb_pheromones *= (1 - self.rho)
            self.sl_type_pheromones *= (1 - self.rho)
            self.ma_stop_pheromones *= (1 - self.rho)

            for (sma, roc, sl, reb, sl_type, ma_stop), score in ants_results:
                if score > 0:
                    self.sma_pheromones[np.where(self.sma_range == sma)[0][0]] += score
                    self.roc_pheromones[np.where(self.roc_range == roc)[0][0]] += score
                    self.sl_pheromones[np.where(self.sl_range == sl)[0][0]] += score
                    self.reb_pheromones[np.where(self.reb_range == reb)[0][0]] += score
                    self.sl_type_pheromones[self.sl_type_range.index(sl_type)] += score
                    self.ma_stop_pheromones[self.ma_stop_range.index(ma_stop)] += score

            if gen % 5 == 0:
                print(f"Gen {gen+1}: Best Score {self.best_score:.4f} {self.best_params}")

        return self.best_params, self.best_score

if __name__ == "__main__":
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')
    bt = BacktesterV2(prices, volumes, code_to_name)
    start_date = pd.to_datetime('2019-01-01')
    end_date = pd.to_datetime('2023-12-31')
    optimizer = ACO_Optimizer_V6(bt, n_ants=40, n_iterations=40)
    best_p, best_s = optimizer.optimize(start_date, end_date)
    print(f"\nFinal Best Params: {best_p}")
    with open('best_params_final_v6.pkl', 'wb') as f:
        pickle.dump(best_p, f)
