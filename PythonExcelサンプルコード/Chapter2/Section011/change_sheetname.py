from openpyxl import load_workbook

wb = load_workbook('集計.xlsx')

for i, ws in enumerate(wb.worksheets):
    ws.title = 'ID_' + ws.title
    if (i + 1) % 10 == 0:
        ws.sheet_properties.tabColor = '0000FF'

wb.save('集計_変更後.xlsx')
