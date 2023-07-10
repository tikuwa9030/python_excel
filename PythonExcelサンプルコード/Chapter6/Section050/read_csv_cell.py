import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

wb = Workbook()
ws = wb.active

df = pd.read_csv('uriage.csv', encoding='utf-8')

for row in dataframe_to_rows(df, index=None, header=True):
    ws.append(row)

wb.save('売上高.xlsx')
