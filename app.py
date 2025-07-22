import os
import json
import gspread
from flask import Flask, request, jsonify
from oauth2client.service_account import ServiceAccountCredentials
from style_writer import apply_template_style

# Google認証ファイルのパス
CREDENTIAL_FILE_PATH = "/etc/secrets/credentials.json"

# Googleスプレッドシート設定
SPREADSHEET_ID = "1fKN1EDZTYOlU4OvImQZuifr2owM8MIGgQIr0tu_rX0E"
SHEET_NAME = "printlist"

# Flask初期化
app = Flask(__name__)

# Google Sheets クライアント取得
def get_gspread_client():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIAL_FILE_PATH, scope)
    client = gspread.authorize(creds)
    return client

# テキストからデータを抽出
def parse_text(text):
    result = {
        "製造番号": "",
        "印刷番号": "",
        "製造日": "",
        "会社名": "",
        "製品名": "",
        "製品種類": "",
        "外装包材": "",
        "表面印刷": "",
        "製造個数": ""
    }

    for line in text.splitlines():
        if "製造番号" in line:
            result["製造番号"] = line.split("：")[-1].strip()
        elif "印刷番号" in line:
            result["印刷番号"] = line.split("：")[-1].strip()
        elif "製造日" in line:
            result["製造日"] = line.split("：")[-1].strip()
        elif "会社名" in line:
            result["会社名"] = line.split("：")[-1].strip()
        elif "製品名" in line:
            result["製品名"] = line.split("：")[-1].strip()
        elif "製品種類" in line:
            result["製品種類"] = line.split("：")[-1].strip()
        elif "外装包材" in line:
            result["外装包材"] = line.split("：")[-1].strip()
        elif "表面印刷：" in line:
            result["表面印刷"] = line.split("：")[-1].strip()
        elif "製造個数" in line:
            result["製造個数"] = line.split("：")[-1].strip()

    return result

# Google Sheets に書き込む
def write_to_sheet(data):
    client = get_gspread_client()
    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)

    # A列（ステータス列）を基に次の空きブロック行を探す
    values = sheet.col_values(1)
    start_row = len(values) + 1 if values else 3
    if start_row < 3:
        start_row = 3

    # スタイル適用（テンプレート反映）
    apply_template_style(sheet, start_row)

    # 各セルにデータ反映
    sheet.update(f"A{start_row}", [[""]])  # ステータス空欄
    sheet.update(f"B{start_row}", [["ファイル名"]])
    sheet.update(f"C{start_row}", [[data["製造番号"]]])
    sheet.update(f"C{start_row + 4}", [[data["印刷番号"]]])
    sheet.update(f"D{start_row}", [[data["製造日"]]])
    sheet.update(f"E{start_row}", [[data["会社名"]]])
    sheet.update(f"E{start_row + 2}", [[data["製品名"]]])
    sheet.update(f"G{start_row}", [[data["製品種類"]]])
    sheet.update(f"G{start_row + 3}", [[data["外装包材"]]])
    sheet.update(f"G{start_row + 6}", [[data["表面印刷"]]])
    sheet.update(f"L{start_row}", [[data["製造個数"]]])
    sheet.update(f"O{start_row}", [["未設定"]])  # 担当

# ヘルスチェック
@app.route("/", methods=["GET"])
def health_check():
    return "OK", 200

# テキスト入力エンドポイント
@app.route("/submit", methods=["POST"])
def submit_text():
    if not request.is_json:
        return jsonify({"error": "JSON形式で送信してください"}), 400

    data = request.get_json()
    text = data.get("text", "")

    if not text.strip():
        return jsonify({"error": "テキストが空です"}), 400

    try:
        parsed_data = parse_text(text)
        write_to_sheet(parsed_data)
        return jsonify({"message": "データを登録しました", "data": parsed_data}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# 実行
if __name__ == "__main__":
    app.run(debug=True, port=10000)
