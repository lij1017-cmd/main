import pandas as pd
import numpy as np
import random
import pickle
from backtester import Backtester, calculate_metrics
from data_prep import clean_data

def optimize():
    prices, code_to_name = clean_data('個股1.xlsx')
    bt = Backtester(prices)

    num_ants = 15
    num_iterations = 15
    evaporation_rate = 0.3
    q = 1.0

    sma_range = list(range(2, 81))
    roc_range = list(range(2, 81))
    sl_range = [i/100 for i in range(1, 10)]

    sma_pheromones = np.ones(len(sma_range))
    roc_pheromones = np.ones(len(roc_range))
    sl_pheromones = np.ones(len(sl_range))

    best_params = None
    best_calmar = -1
    memo = {}

    for iter in range(num_iterations):
        iter_results = []
        for ant in range(num_ants):
            def select(options, pheromones):
                p = pheromones + 1e-6
                probs = p / p.sum()
                return random.choices(options, weights=probs)[0]

            sma_p = select(sma_range, sma_pheromones)
            roc_p = select(roc_range, roc_pheromones)
            sl_p = select(sl_range, sl_pheromones)

            params = (sma_p, roc_p, sl_p)
            if params in memo:
                calmar = memo[params]
            else:
                eq, trades, holdings, log = bt.run(sma_p, roc_p, sl_p)
                cagr, mdd, calmar, win_rate = calculate_metrics(eq, trades)
                memo[params] = calmar

            iter_results.append((params, calmar))
            if calmar > best_calmar:
                best_calmar = calmar
                best_params = params

        sma_pheromones *= (1 - evaporation_rate)
        roc_pheromones *= (1 - evaporation_rate)
        sl_pheromones *= (1 - evaporation_rate)

        for params, calmar in iter_results:
            if calmar > 0:
                sma_pheromones[sma_range.index(params[0])] += q * calmar
                roc_pheromones[roc_range.index(params[1])] += q * calmar
                sl_pheromones[sl_range.index(params[2])] += q * calmar

    print(f"Best: {best_params}, Calmar: {best_calmar:.2f}")

if __name__ == "__main__":
    optimize()
