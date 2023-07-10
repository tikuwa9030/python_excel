from openpyxl import Workbook

wb = Workbook()
ws = wb.active

for row_count in range(1, 5):
    cell_no = f'A{row_count}'
    ws[cell_no] = 'Hello'

wb.save('test.xlsx')
