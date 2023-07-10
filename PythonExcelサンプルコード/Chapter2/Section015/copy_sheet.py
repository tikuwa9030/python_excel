from datetime import date

from openpyxl import load_workbook

wb = load_workbook('議事録.xlsx')
for ws in wb.worksheets:
    ws.sheet_view.tabSelected = None

ws_template = wb['template']
ws_copy = wb.copy_worksheet(ws_template)

today = date.today()
ws_copy.title = f'{today:%Y-%m-%d}'

wb.move_sheet(ws_copy, offset=-wb.index(ws_copy))

wb.active = 0
wb.save('議事録_変更後.xlsx')
