import nbformat as nbf
def create_notebook():
    nb = nbf.v4.new_notebook()
    cells = []
    cells.append(nbf.v4.new_markdown_cell("# Asset Class Trend Following 策略回測與最佳化"))
    with open('data_prep.py', 'r') as f: data_prep_code = f.read()
    cells.append(nbf.v4.new_code_cell(data_prep_code))
    with open('backtester.py', 'r') as f: bt_code = f.read()
    cells.append(nbf.v4.new_code_cell(bt_code))
    res_code = """
import pickle
with open('best_params.pkl', 'rb') as f: best_params = pickle.load(f)
prices = pd.read_pickle('prices_cleaned.pkl')
bt = Backtester(prices)
eq, trades, holdings, rebalance_log = bt.run(*best_params)
cagr, mdd, calmar, win_rate = calculate_metrics(eq, trades)
print(f"SMA={best_params[0]}, ROC={best_params[1]}, SL={best_params[2]}")
print(f"CAGR: {cagr:.2%}, MaxDD: {mdd:.2%}, Calmar: {calmar:.2f}")
"""
    cells.append(nbf.v4.new_code_cell(res_code))
    nb['cells'] = cells
    with open('trendstrategy_equity2024.ipynb', 'w') as f: nbf.write(nb, f)
if __name__ == "__main__":
    create_notebook()
