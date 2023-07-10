from openpyxl import Workbook

# 作成するシート数を入力
count = input("全シート数は：")

# ブックを作成
wb = Workbook()

# アクティブなシートを選択
ws = wb.active

# ワークシートを作成
ws.title = "概要_1"

# 入力したシート数分を複製して、シート名をつける
for i in range(2, int(count) + 1):

    # ワークシートを作成
    wb.create_sheet(title=f"概要_{i}")


# ブックを保存
wb.save("資料.xlsx")
