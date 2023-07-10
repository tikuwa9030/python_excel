from openpyxl import load_workbook

print_area = 'A1:D50'
print_title_rows = '1:5'
header_text = '&F'
footer_text = '&P / &Nページ'

wb = load_workbook('チェックリスト.xlsx')
for ws in wb.worksheets:
    ws.print_area = print_area
    ws.print_title_rows = print_title_rows
    ws.oddHeader.center.text = header_text
    ws.oddFooter.center.text = footer_text

    wps = ws.page_setup
    wps.orientation = ws.ORIENTATION_LANDSCAPE
    wps.fitToWidth = 1
    wps.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wps.paperSize = ws.PAPERSIZE_A3

wb.save('チェックリスト_変更後.xlsx')
