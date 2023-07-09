from openpyxl import Workbook

count = input("作成するブック数：")

# 入力値のブック数分のファイルを作成
for i in range(int(count)):
    # エクセルファイルを作成
    wb = Workbook()
    # 選択されているシートを活性
    ws = wb.active
    # シート名を挿入
    ws.title = "概要"
    # エクセルファイルを保存
    wb.save(f'資料_{i + 1}.xlsx')
