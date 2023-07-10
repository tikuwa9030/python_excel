import json

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

wb = Workbook()
ws = wb.active
ws.column_dimensions['A'].width = 30
ws.column_dimensions['B'].width = 30

with open('search.json', encoding='utf-8') as f:
    data = json.load(f)

items = data['items']
df = pd.json_normalize(items)
select_df = df[['title', 'link', 'snippet']]

for row in dataframe_to_rows(select_df, index=None, header=True):
    ws.append(row)

wb.save('検索結果.xlsx')
