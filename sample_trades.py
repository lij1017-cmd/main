import pandas as pd
import numpy as np

# Load the performance sheet
df = pd.read_excel('equity(成-1).xlsx', sheet_name='Performance')

# Sample 5 round-trip trades
sampled_trades = df.sample(5, random_state=42)

output_rows = []
for _, row in sampled_trades.iterrows():
    # 1. Buy Record
    output_rows.append({
        '日期': row['進場日期'].strftime('%Y-%m-%d'),
        '股票代碼': str(row['股票代號']),
        '交易類型': '買進',
        '股數': int(row['股數']),
        '單價': round(row['進場價格'], 2),
        '交易金額': round(row['買進金額'], 0),
        '買進手續費': round(row['買進手續費'], 0),
        '賣出手續費': 0,
        '證交稅': 0,
        '持有天數': '-',
        '毛利潤': '-',
        '淨損益': '-'
    })
    # 2. Sell Record
    gross_profit = row['賣出金額'] - row['買進金額']
    net_profit = (row['賣出金額'] - row['賣出手續費'] - row['證交稅']) - (row['買進金額'] + row['買進手續費'])
    output_rows.append({
        '日期': row['出場日期'].strftime('%Y-%m-%d'),
        '股票代碼': str(row['股票代號']),
        '交易類型': '賣出',
        '股數': int(row['股數']),
        '單價': round(row['出場價格'], 2),
        '交易金額': round(row['賣出金額'], 0),
        '買進手續費': 0,
        '賣出手續費': round(row['賣出手續費'], 0),
        '證交稅': round(row['證交稅'], 0),
        '持有天數': int(row['持有天數']),
        '毛利潤': round(gross_profit, 0),
        '淨損益': round(net_profit, 0)
    })

final_df = pd.DataFrame(output_rows)
print(final_df.to_markdown(index=False))
