from gspread_formatting import (
    CellFormat, Color, TextFormat,
    set_data_validation_for_cell_range, DataValidationRule, BooleanCondition
)

# ステータス列と担当列のリスト候補
STATUS_LIST = ["", "仕掛中", "完了"]
TANTO_LIST = ["未設定", "小島", "小林", "北裏", "岩﨑", "小野"]

def add_dropdown(ws, cell_range, options):
    rule = DataValidationRule(
        condition=BooleanCondition('ONE_OF_LIST', options),
        showCustomUi=True
    )
    set_data_validation_for_cell_range(ws, cell_range, rule)

def apply_template_style(ws, start_row):
    # 背景・太字・中央寄せスタイル
    green_bold = CellFormat(
        backgroundColor=Color(0.85, 0.93, 0.81),
        textFormat=TextFormat(bold=True),
        horizontalAlignment='CENTER'
    )
    blue_bold = CellFormat(
        backgroundColor=Color(0.85, 0.92, 0.98),
        textFormat=TextFormat(bold=True),
        horizontalAlignment='CENTER'
    )

    # ステータス（A列）・担当（O列）の範囲設定（ブロック内10行分）
    status_range = f"A{start_row + 2}:A{start_row + 10}"
    tanto_range = f"O{start_row + 2}:O{start_row + 10}"

    add_dropdown(ws, status_range, STATUS_LIST)
    add_dropdown(ws, tanto_range, TANTO_LIST)
