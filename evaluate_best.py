import pandas as pd
from backtest_equity2MA import clean_data, Backtester, calculate_metrics

prices, volumes, names = clean_data('樣本集-1.xlsx')
bt = Backtester(prices, volumes, names)

# Best from training: (170, 290, 14, 'ma', 50, 10)
s1, s2, r, st, sv, rb = 170, 290, 14, 'ma', 50, 10

eq, trades, hold, trades2, daily = bt.run(s1, s2, r, sv, rb, st)

# 2019-2023
eq_train = eq[(eq['日期'] >= '2019-01-02') & (eq['日期'] <= '2023-12-31')]
cagr_tr, mdd_tr, cal_tr, ret_tr = calculate_metrics(eq_train)
print(f"Train (2019-2023): CAGR={cagr_tr:.2%}, MDD={mdd_tr:.2%}, Calmar={cal_tr:.2f}")

# 2024-2025
eq_test = eq[(eq['日期'] >= '2024-01-02') & (eq['日期'] <= '2025-12-31')]
# For test set metrics, we need to reset the initial equity to 1.0 or the start value
if not eq_test.empty:
    eq_test = eq_test.copy()
    initial_val = eq_test['權益值'].iloc[0]
    eq_test['權益值'] = eq_test['權益值'] / initial_val * 30000000
    # Recalculate drawdown
    peak = eq_test['權益值'].cummax()
    eq_test['回撤(Drawdown)'] = (eq_test['權益值'] - peak) / peak

    cagr_te, mdd_te, cal_te, ret_te = calculate_metrics(eq_test)
    print(f"Test (2024-2025): CAGR={cagr_te:.2%}, MDD={mdd_te:.2%}, Calmar={cal_te:.2f}")

# Full
cagr_f, mdd_f, cal_f, ret_f = calculate_metrics(eq)
print(f"Full (2019-2025): CAGR={cagr_f:.2%}, MDD={mdd_f:.2%}, Calmar={cal_f:.2f}")
