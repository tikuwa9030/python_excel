from openpyxl import load_workbook

wb = load_workbook('チェックリスト.xlsx')

wb.move_sheet('まとめ', offset=1)

wb.save('チェックリスト_変更後.xlsx')
