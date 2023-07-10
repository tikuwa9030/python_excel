from openpyxl import load_workbook

wb = load_workbook('売上実績.xlsx')
ws = wb.active
column_width = {'B': 6, 'C': 30, 'D': 30}

for col, width in column_width.items():
    ws.column_dimensions[col].width = width

wb.save('売上実績_変更後.xlsx')
