import json

with open('trendstrategy_equity25_cost.ipynb', 'r', encoding='utf-8') as f:
    nb = json.load(f)

for cell in nb['cells']:
    if cell['cell_type'] == 'code':
        new_source = []
        for line in cell['source']:
            line = line.replace("（買進 0.1425%，賣出 0.1425% + 0.3% 證交稅）。，針對", "（買進 0.1425%，賣出 0.1425% + 0.3% 證交稅），針對")
            new_source.append(line)
        cell['source'] = new_source

with open('trendstrategy_equity25_cost.ipynb', 'w', encoding='utf-8') as f:
    json.dump(nb, f, ensure_ascii=False, indent=1)
