from openpyxl import load_workbook

wb = load_workbook('チェックリスト.xlsx')

for ws in wb.worksheets:
    ws.sheet_view.tabSelected = None

ws_matome = wb['まとめ']
wb.move_sheet(ws_matome, offset=-wb.index(ws_matome))

wb.active = 0
wb.save('チェックリスト_変更後.xlsx')
