import json
import pandas as pd
import numpy as np
import os

with open('trendstrategy_equity25_cost.ipynb', 'r', encoding='utf-8') as f:
    nb = json.load(f)

for cell in nb['cells']:
    if cell['cell_type'] == 'code':
        code = "".join(cell['source'])
        exec(code, globals())
