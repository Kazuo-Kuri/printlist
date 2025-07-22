import gspread
from gspread_formatting import (
    DataValidationRule,
    BooleanCondition,
    set_data_validation_for_cell_range
)

# ドロップダウン候補（要件に従って列挙）
STATUS_LIST = ["", "仕掛中", "完了"]
TANTOU_LIST = ["未設定", "小島", "小林", "北裏", "岩﨑", "小野"]

def set_dropdown_validation(worksheet, col_letter, row_start, row_end, values):
    """
    指定した列の範囲に対して、リスト形式のデータ検証（ドロップダウン）を設定。
    """
    cell_range = f"{col_letter}{row_start}:{col_letter}{row_end}"
    rule = DataValidationRule(
        condition=BooleanCondition('ONE_OF_LIST', values),
        showCustomUi=True,
        strict=True
    )
    set_data_validation_for_cell_range(worksheet, cell_range, rule)

def apply_validations(sheet):
    """
    スプレッドシートの対象列に対して、ドロップダウンを設定する。
    """
    max_rows = 1000  # 想定される最大データ行数（必要に応じて変更）

    # ステータス列（A列） → A3:A1000 に設定
    set_dropdown_validation(sheet, "A", 3, max_rows, STATUS_LIST)

    # 担当列（O列） → O3:O1000 に設定（OはA〜Z→AA〜列の15番目）
    set_dropdown_validation(sheet, "O", 3, max_rows, TANTOU_LIST)
