from openpyxl import Workbook

count = input('全シート数: ')

wb = Workbook()
ws = wb.active
ws.title = '概要_1'

for i in range(2, int(count) + 1):
    wb.create_sheet(title=f'概要_{i}')

wb.save('資料.xlsx')
