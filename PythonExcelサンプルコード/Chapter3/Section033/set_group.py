from openpyxl import load_workbook

wb = load_workbook('作業時間.xlsx')
ws = wb.active
for row_no in [(5, 20), (22, 27), (29, 30)]:
    ws.row_dimensions.group(*row_no, outline_level=1, hidden=True)

ws.column_dimensions.group('D', outline_level=1, hidden=True)
wb.save('作業時間_変更後.xlsx')
