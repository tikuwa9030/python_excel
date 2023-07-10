import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

wb = Workbook()
ws = wb.active

df = pd.read_csv('uriage_without_header.csv', encoding='utf-8',
                 header=None,
                 names=['部門', '小分類', '当期売上', '前期売上'])

for row in dataframe_to_rows(df, index=None, header=True):
    ws.append(row)

wb.save('売上高.xlsx')
