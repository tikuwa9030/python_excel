import logging
import sys

from openpyxl import Workbook

logging.basicConfig(filename='create_book.log',
                    level=logging.INFO,
                    format='%(asctime)s: [%(levelname)s] %(message)s')

logging.info('処理を開始しました')
try:
    count = sys.argv[1]
    for i in range(int(count)):
        wb = Workbook()
        ws = wb.active
        ws.title = '概要'

        file_name = f'資料_{i + 1}.xlsx'
        wb.save(file_name)
        logging.info('ブックを作成しました: %s', file_name)

except Exception:
    logging.exception('例外が発生しました')

logging.info('処理が終了しました')
