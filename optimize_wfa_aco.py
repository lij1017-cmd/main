import pandas as pd
import numpy as np
import pickle
from run_backtest_equity2025新_1 import clean_data, Backtester, calculate_metrics

class ACO_Optimizer_WFA:
    def __init__(self, backtester, n_ants=10, n_iterations=10, rho=0.1, alpha=1):
        self.bt = backtester
        self.n_ants = n_ants
        self.n_iterations = n_iterations
        self.rho = rho # Evaporation rate
        self.alpha = alpha # Pheromone influence

        # Search space
        self.sma_range = np.arange(10, 101, 1)
        self.roc_range = np.arange(10, 101, 1)
        self.sl_range = np.arange(0.05, 0.151, 0.005)

        # Initialize pheromones
        self.sma_pheromones = np.ones(len(self.sma_range))
        self.roc_pheromones = np.ones(len(self.roc_range))
        self.sl_pheromones = np.ones(len(self.sl_range))

        self.best_params = None
        self.best_score = -np.inf
        self.history = []

        # Fixed rebalance cycle for optimization
        self.rebalance_cycle = 7

    def _select_param(self, range_vals, pheromones):
        probs = pheromones ** self.alpha
        probs /= probs.sum()
        return np.random.choice(range_vals, p=probs)

    def optimize(self, start_date, end_date):
        # Patch the Backtester to use 8-day rebalance

        def run_fixed(bt, sma, roc, sl):
            # We use the full period for optimization
            # run_backtest_equity2025新_1.Backtester.run definition:
            # run(self, sma_period, roc_period, stop_loss_pct, rebalance_interval=6)
            eq, trades, _ = bt.run(int(sma), int(roc), float(sl), rebalance_interval=7)
            mask = (eq.index >= start_date) & (eq.index <= end_date)
            return eq[mask], trades, None

        for gen in range(self.n_iterations):
            gen_best_params = None
            gen_best_score = -np.inf

            ants_results = []
            for ant in range(self.n_ants):
                sma = self._select_param(self.sma_range, self.sma_pheromones)
                roc = self._select_param(self.roc_range, self.roc_pheromones)
                sl = self._select_param(self.sl_range, self.sl_pheromones)

                eq, trades, costs = run_fixed(self.bt, sma, roc, sl)
                cagr, mdd, calmar, _ = calculate_metrics(eq)

                # Objective: Maximize Calmar Ratio
                score = calmar
                if mdd > -0.05 or len(trades) < 10: # Penalty for unrealistic results
                     score = -1

                ants_results.append(((sma, roc, sl), score))

                if score > gen_best_score:
                    gen_best_score = score
                    gen_best_params = (sma, roc, sl)

                if score > self.best_score:
                    self.best_score = score
                    self.best_params = (sma, roc, sl)

            # Update pheromones
            self.sma_pheromones *= (1 - self.rho)
            self.roc_pheromones *= (1 - self.rho)
            self.sl_pheromones *= (1 - self.rho)

            for (sma, roc, sl), score in ants_results:
                if score > 0:
                    self.sma_pheromones[np.where(self.sma_range == sma)[0][0]] += score
                    self.roc_pheromones[np.where(self.roc_range == roc)[0][0]] += score
                    self.sl_pheromones[np.where(self.sl_range == sl)[0][0]] += score

            print(f"Gen {gen+1}: Best Score {gen_best_score:.4f} (SMA={gen_best_params[0]}, ROC={gen_best_params[1]}, SL={gen_best_params[2]:.3f})")
            self.history.append((gen, gen_best_params, gen_best_score))

        return self.best_params, self.best_score

if __name__ == "__main__":
    data_file = '個股合-1.xlsx'
    prices, code_to_name = clean_data(data_file)
    bt = Backtester(prices, code_to_name)

    # Optimize over a large representative period
    start_date = pd.to_datetime('2019/1/2')
    end_date = pd.to_datetime('2025/12/31')

    optimizer = ACO_Optimizer_WFA(bt, n_ants=20, n_iterations=10)
    best_p, best_s = optimizer.optimize(start_date, end_date)

    print(f"\nFinal Best Params: SMA={best_p[0]}, ROC={best_p[1]}, SL={best_p[2]:.3f}")
    print(f"Final Best Score (Calmar): {best_s:.4f}")

    with open('best_params_aco_wfa.pkl', 'wb') as f:
        pickle.dump(best_p, f)
