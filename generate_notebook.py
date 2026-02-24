import nbformat as nbf
def create_notebook():
    nb = nbf.v4.new_notebook()
    cells = [
        nbf.v4.new_markdown_cell("# Asset Class Trend Following 策略回測與最佳化"),
        nbf.v4.new_code_cell("import pandas as pd\nimport numpy as np\nimport matplotlib.pyplot as plt"),
        nbf.v4.new_markdown_cell("## 1. 資料清理"),
        nbf.v4.new_code_cell(open('data_prep.py').read()),
        nbf.v4.new_markdown_cell("## 2. 策略引擎"),
        nbf.v4.new_code_cell(open('backtester.py').read()),
        nbf.v4.new_markdown_cell("## 3. 執行回測"),
        nbf.v4.new_code_cell("best_params = (69, 23, 0.09)\nprices, code_to_name = clean_data('個股1.xlsx')\nbt = Backtester(prices)\neq, trades, holdings, action_log = bt.run(*best_params)\ncagr, mdd, calmar, win_rate = calculate_metrics(eq, trades)\nprint(f'SMA={best_params[0]}, ROC={best_params[1]}, SL={best_params[2]}')\nprint(f'CAGR: {cagr:.2%}, MaxDD: {mdd:.2%}, Calmar: {calmar:.2f}')\nplt.figure(figsize=(12, 6))\nplt.plot(eq)\nplt.title('Equity Curve')\nplt.grid(True)\nplt.show()")
    ]
    nb['cells'] = cells
    with open('trendstrategy_equity2024.ipynb', 'w') as f: nbf.write(nb, f)
if __name__ == "__main__":
    create_notebook()
