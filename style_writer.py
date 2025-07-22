import gspread
from gspread_formatting import (
    DataValidationRule,
    BooleanCondition,
    set_data_validation_for_cell_range
)

# ドロップダウン候補
STATUS_LIST = ["", "仕掛中", "完了"]
TANTOU_LIST = ["未設定", "小島", "小林", "北裏", "岩﨑", "小野"]

def set_dropdown_validation(worksheet, cell, values):
    """
    指定したセルにドロップダウンを設定する。
    """
    rule = DataValidationRule(
        condition=BooleanCondition('ONE_OF_LIST', values),
        showCustomUi=True,
        strict=True
    )
    set_data_validation_for_cell_range(worksheet, cell, rule)

def apply_validations(sheet, block_start_row):
    """
    ドロップダウンを特定の1ブロック（A1:N10）の相対位置に適用する。
    block_start_row: ブロックの開始行（1ブロック目なら1、2ブロック目なら11、以降 +10）
    """
    # ブロック内の相対位置 A6 → ステータス, O6 → 担当
    status_cell = f"A{block_start_row + 5}"  # A6相当
    tantou_cell = f"O{block_start_row + 5}"  # O6相当

    set_dropdown_validation(sheet, status_cell, STATUS_LIST)
    set_dropdown_validation(sheet, tantou_cell, TANTOU_LIST)
