from gspread_formatting import *

# チェックボックスの挿入（TRUE/FALSE 書式）
def add_checkboxes(ws, cell_ranges):
    if isinstance(cell_ranges, str):
        cell_ranges = [cell_ranges]
    for rng in cell_ranges:
        set_data_validation_for_cell_range(
            ws, rng,
            DataValidationRule(
                BooleanCondition('BOOLEAN'),
                showCustomUi=True
            )
        )

# ドロップダウンリストの挿入（担当列）
def add_dropdown(ws, cell_range, options):
    set_data_validation_for_cell_range(
        ws, cell_range,
        DataValidationRule(
            BooleanCondition('ONE_OF_LIST', options),
            showCustomUi=True
        )
    )

# スタイル定義（色・太字・中央寄せなど）
def apply_template_style(ws, start_row):
    # === セル背景色 + 太字 ===
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

    # タイトル背景（例：A2:N2 → start_row+1）
    format_cell_range(ws, f"A{start_row+1}:N{start_row+1}", green_bold)

    # C列（製造番号・印刷番号）背景
    format_cell_range(ws, f"C{start_row+3}:C{start_row+4}", blue_bold)

    # === チェックボックス挿入 ===
    checkbox_ranges = [
        f"G{start_row+3}", f"G{start_row+6}", f"G{start_row+9}",  # フック紙・フィルム色・リボン色
        f"H{start_row+3}",  # 整合性
        f"I{start_row+3}", f"I{start_row+6}",  # フィルム継ぎ有無
        f"K{start_row+3}", f"K{start_row+6}",  # フィルム継ぎ確認
        f"M{start_row+3}"  # 数量用
    ]
    add_checkboxes(ws, checkbox_ranges)

    # === 担当列ドロップダウン ===
    add_dropdown(ws, f"N{start_row+3}", ["未設定", "山田", "田中", "佐藤"])

    # === その他：必要に応じて罫線や結合セル追加（高度処理）
    # ここでは gspread-formatting では未対応なため割愛
    # 結合セルや罫線の細部は Sheets UI 側テンプレート維持で対応を推奨
