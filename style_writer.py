import gspread
from gspread_formatting import (
    DataValidationRule,
    BooleanCondition,
    set_data_validation_for_cell_range
)

# ドロップダウンリストの定義
STATUS_LIST = ["", "仕掛中", "完了"]
TANTOU_LIST = ["未設定", "小島", "小林", "北裏", "岩﨑", "小野"]

def set_dropdown_validation(worksheet, col_letter, row_start, row_end, values):
    """
    指定した列の範囲にドロップダウンリストを設定します。
    """
    cell_range = f"{col_letter}{row_start}:{col_letter}{row_end}"
    rule = DataValidationRule(
        condition=BooleanCondition('ONE_OF_LIST', values),
        showCustomUi=True,
        strict=True
    )
    set_data_validation_for_cell_range(worksheet, cell_range, rule)

def apply_template_style(sheet, start_row):
    """
    テンプレートのスタイルを特定のブロックに適用する関数。
    現時点では、ドロップダウンリスト適用が中心。
    """
    # 書き込むブロックはA3〜O100まで（必要に応じて調整可能）
    max_rows = max(start_row + 10, 100)

    # A列 → ステータスドロップダウン
    set_dropdown_validation(sheet, "A", 3, max_rows, STATUS_LIST)

    # O列 → 担当者ドロップダウン（注意：A列追加によりO列に移動済み）
    set_dropdown_validation(sheet, "O", 3, max_rows, TANTOU_LIST)
