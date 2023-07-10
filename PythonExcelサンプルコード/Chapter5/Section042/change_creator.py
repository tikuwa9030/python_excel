from openpyxl import load_workbook

wb = load_workbook('資料.xlsx')

name = '鈴木太郎'
properties = wb.properties
properties.creator = name
properties.lastModifiedBy = name

wb.save('資料_変更.xlsx')
