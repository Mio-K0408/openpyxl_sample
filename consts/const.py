from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
# レイアウト指定用の変数
INPUT_LAYPUT = [1,2,1,2,2,3]

# 1項目の行数
ITEM_ROW = 10

# 開始位置
START_ROW = 1
START_COLUMN = 1

# セルの塗りつぶし色（水色的な）
CELL_FILL_COLOR = PatternFill(fgColor='B2DDF0', fill_type='solid')

# 罫線（黒）
black_thin = Side(color='000000', border_style='thin')

# フォント名
FONT_NAME = 'Meiryo UI'