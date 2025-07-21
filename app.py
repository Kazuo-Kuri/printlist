from flask import Flask, render_template, request, send_file
import os
import io
import json
import re
from openpyxl import load_workbook
import gspread
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)

# --- Google認証と設定 ---
CREDENTIAL_FILE_PATH = "/etc/secrets/credentials.json"

def get_credentials():
    with open(CREDENTIAL_FILE_PATH, "r", encoding="utf-8") as f:
        credentials_dict = json.load(f)
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    return ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)

# gspread認証
credentials = get_credentials()
gc = gspread.authorize(credentials)

SPREADSHEET_ID = "1fKN1EDZTYOlU4OvImQZuifr2owM8MIGgQIr0tu_rX0E"
TEMPLATE_SHEET_NAME = "Sheet1"     # テンプレート用
OUTPUT_SHEET_NAME = "printlist"    # 書き出し先
EXCEL_TEMPLATE_PATH = "printlist_form.xlsx"

# --- アップロード受付エンドポイント ---
@app.route("/upload", methods=["POST"])
def upload():
    try:
        data = request.form.to_dict()

        # Excel出力
        excel_bytes = generate_excel(data)

        # スプレッドシート出力
        write_to_spreadsheet(data)

        return send_file(
            io.BytesIO(excel_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="printlist_output.xlsx"
        )

    except Exception as e:
        return jsonify({"error": str(e)})

# --- Excelファイル生成処理 ---
def generate_excel(data):
    wb = load_workbook(EXCEL_TEMPLATE_PATH)
    ws = wb.active

    mapping = {
        "B2": data.get("製造日", ""),
        "E1": data.get("製造番号", ""),
        "E2": data.get("印刷番号", ""),
        "B3": data.get("会社名", ""),
        "B4": data.get("製品名", "")
    }

    for cell, value in mapping.items():
        ws[cell] = value

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.read()

# --- スプレッドシート書き出し処理 ---
def write_to_spreadsheet(data_dict):
    ss = gc.open_by_key(SPREADSHEET_ID)
    template_ws = ss.worksheet(TEMPLATE_SHEET_NAME)
    output_ws = ss.worksheet(OUTPUT_SHEET_NAME)

    # 次の書き込み位置（10行ブロック単位）
    existing_rows = len(output_ws.get_all_values())
    block_index = max((existing_rows - 2) // 10, 0)
    start_row = block_index * 10 + 1

    # テンプレート（A1:N10）をコピーして出力シートへ貼り付け
    template_range = template_ws.get_values('A1:N10')
    for i, row in enumerate(template_range):
        output_ws.update(f"A{start_row + i}:N{start_row + i}", [row])

    # 指定セルにデータ書き込み
    cell_map = {
        "A3": data_dict.get("印刷データ", ""),
        "B3": data_dict.get("ファイル名", ""),
        "C3": data_dict.get("製造番号", ""),
        "C7": data_dict.get("印刷番号", ""),
        "D3": data_dict.get("製造日", ""),
        "E3": data_dict.get("会社名", ""),
        "E5": data_dict.get("製品名", ""),
        "G3": data_dict.get("製品種類", ""),
        "G6": data_dict.get("外装包材", ""),
        "G9": data_dict.get("表面印刷", ""),
        "L3": data_dict.get("製造個数", "")
    }

    for cell_a1, value in cell_map.items():
        row = int(cell_a1[1:])
        col = ord(cell_a1[0].upper()) - 65 + 1
        output_ws.update_cell(start_row + (row - 1), col, value)

# --- Flask起動 ---
if __name__ == "__main__":
    app.run(debug=True)
