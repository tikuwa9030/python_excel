from openpyxl import Workbook

wb = Workbook()
ws = wb.active
wb.save('sample.xlsx')
