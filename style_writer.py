import gspread
from gspread_formatting import (
    DataValidationRule,
    BooleanCondition,
    set_data_validation_for_cell_range
)

# ステータスと担当のリスト項目
STATUS_LIST = ["", "仕掛中", "完了"]
TANTOU_LIST = ["未設定", "小島", "小林", "北裏", "岩﨑", "小野"]

def set_dropdown_validation(worksheet, col_letter, row_start, row_end, values):
    """
    指定した列の範囲にリスト形式のドロップダウンを設定する。
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
    スプレッドシート全体に必要なバリデーション（ドロップダウン）を適用する関数。
    """
    max_rows = 100  # 適用する最大行数

    # A列 → ステータス用ドロップダウン（A3〜A100）
    set_dropdown_validation(sheet, "A", 3, max_rows, STATUS_LIST)

    # O列 → 担当者用ドロップダウン（O3〜O100）
    set_dropdown_validation(sheet, "O", 3, max_rows, TANTOU_LIST)
