from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

wb = load_workbook('課題一覧.xlsx')
ws = wb.active

validation = DataValidation(type='list',
                            formula1='"対応待ち,対応中,完了"',
                            allow_blank=False)

validation.add('D3:D10')

ws.add_data_validation(validation)
wb.save('課題一覧_変更後.xlsx')
