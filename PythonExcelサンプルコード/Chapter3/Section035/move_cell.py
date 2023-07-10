from openpyxl import load_workbook

wb = load_workbook('テーブル定義書.xlsx')
ws = wb.active

ws.move_range('A3:F7', rows=2, cols=1)

wb.save('テーブル定義書_変更後.xlsx')
