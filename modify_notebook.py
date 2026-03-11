import json

with open('trendstrategy_equity25.ipynb', 'r', encoding='utf-8') as f:
    nb = json.load(f)

def replace_in_source(source, replacements):
    new_source = []
    for line in source:
        for old, new in replacements.items():
            line = line.replace(old, new)
        new_source.append(line)
    return new_source

replacements_params = {
    "OUTPUT_EXCEL = 'trendstrategy_results_equity25.xlsx'": "OUTPUT_EXCEL = 'trendstrategy_results_equity25(成).xlsx'",
    "OUTPUT_MD = 'reproduce_report25.md'": "OUTPUT_MD = 'reproduce_report25(成).md'"
}

replacements_bt = {
    "capital += info['shares'] * sell_price": "capital += info['shares'] * sell_price * (1 - 0.001425 - 0.003) # 扣除手續費與證交稅",
    "shares = slot_capital // buy_price": "shares = slot_capital // (buy_price * 1.001425) # 考慮買進手續費",
    "if shares > 0 and capital >= shares * buy_price:": "buy_cost = shares * buy_price * 1.001425\n                    if shares > 0 and capital >= buy_cost:",
    "capital -= shares * buy_price": "capital -= buy_cost"
}

replacements_report = {
    "# Asset Class Trend Following 策略重現報告 (2025)": "# Asset Class Trend Following 策略重現報告 (2025 - 納入交易成本)",
    "本策略採用 Asset Class Trend Following 方法": "本策略採用 Asset Class Trend Following 方法，並納入交易成本（買進 0.1425%，賣出 0.1425% + 0.3% 證交稅）。"
}

for cell in nb['cells']:
    if cell['cell_type'] == 'code':
        cell['source'] = replace_in_source(cell['source'], replacements_params)
        cell['source'] = replace_in_source(cell['source'], replacements_bt)
        cell['source'] = replace_in_source(cell['source'], replacements_report)

with open('trendstrategy_equity25_cost.ipynb', 'w', encoding='utf-8') as f:
    json.dump(nb, f, ensure_ascii=False, indent=1)
