import json
import os

# Update the notebook to use the new filename
with open('trendstrategy_equity25_cost.ipynb', 'r', encoding='utf-8') as f:
    nb = json.load(f)

for cell in nb['cells']:
    if cell['cell_type'] == 'code':
        new_source = []
        for line in cell['source']:
            line = line.replace("OUTPUT_EXCEL = 'trendstrategy_results_equity25(成).xlsx'", "OUTPUT_EXCEL = 'equity(成-1).xlsx'")
            new_source.append(line)
        cell['source'] = new_source

with open('trendstrategy_equity25_cost.ipynb', 'w', encoding='utf-8') as f:
    json.dump(nb, f, ensure_ascii=False, indent=1)

# Run the notebook again to generate the file with the new name
# (Or just rename the existing one, but running ensures consistency)
