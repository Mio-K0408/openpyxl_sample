from consts import const

import os
import openpyxl as excel
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# 現在位置
current_row = const.START_ROW
current_column = const.START_COLUMN

# エクセルの保存先
create_file = os.path.join(os.getcwd(),'create','test.xlsx')
print(create_file)

# ブックの新規作成
# 既にブックが存在する場合は、削除してから新規作成
if os.path.exists(create_file):
    os.remove(create_file)

wb = excel.Workbook()
wb.save(create_file)

# 作成したブックの読み込み
wb = excel.load_workbook(create_file)

# 操作対象のシートを指定
ws = wb.worksheets[0]
# シート名の変更
ws.title = '動作確認用'


# フォントの指定（シート全体）
font = Font(name=const.FONT_NAME)
for row in ws:
    for cell in row:
        ws[cell.coordinate].font = font

# 上書き保存
wb.save(create_file)