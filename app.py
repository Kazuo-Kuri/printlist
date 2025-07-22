import os
import json
import gspread
from flask import Flask, request, jsonify, render_template_string
from oauth2client.service_account import ServiceAccountCredentials
from style_writer import apply_validations

# Google認証ファイルのパス
CREDENTIAL_FILE_PATH = "/etc/secrets/credentials.json"

# Googleスプレッドシート設定
SPREADSHEET_ID = "1fKN1EDZTYOlU4OvImQZuifr2owM8MIGgQIr0tu_rX0E"
SHEET_NAME = "printlist"

# Flask初期化
app = Flask(__name__)

# Google Sheets 認証とクライアント取得
def get_gspread_client():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIAL_FILE_PATH, scope)
    client = gspread.authorize(creds)
    return client

# データ抽出ロジック（例：シンプルに項目名ベース）
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

    # 次の空ブロックの開始行を探す（A列の空白） ※A列がステータス列
    values = sheet.col_values(1)
    start_row = len(values) + 1 if values else 3
    if start_row < 3:
        start_row = 3

    apply_validations(sheet)

    # 各セルにデータを反映（A列〜O列）
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

@app.route("/", methods=["GET"])
def index():
    return render_template_string("""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>テキスト抽出アプリ</title>
        <script>
            function clearText() {
                document.getElementById("inputText").value = "";
            }

            function clearData() {
                if (confirm("スプレッドシートの全データを削除してもよろしいですか？")) {
                    fetch("/clear", {
                        method: "POST"
                    })
                    .then(response => response.text())
                    .then(data => alert("結果: " + data))
                    .catch(error => alert("エラー: " + error));
                }
            }
        </script>
    </head>
    <body>
        <h1>テキストから情報抽出</h1>
        <form method="POST" action="/upload_text">
            <textarea id="inputText" name="text" rows="20" cols="100"></textarea><br>
            <button type="submit">送信</button>
            <button type="button" onclick="clearText()">クリア</button>
            <button type="button" onclick="clearData()">データクリア</button>
        </form>
    </body>
    </html>
    """)

@app.route("/upload_text", methods=["POST"])
def upload_text():
    try:
        text = request.form.get("text", "")
        parsed_data = parse_text(text)
        write_to_sheet(parsed_data)
        return "データを書き込みました。", 200
    except Exception as e:
        return f"エラー: {str(e)}", 500

@app.route("/clear", methods=["POST"])
def clear_sheet():
    try:
        client = get_gspread_client()
        sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
        sheet.batch_clear(["A3:O1000"])  # 実際の必要行数に応じて調整可能
        return "シートのデータを削除しました。", 200
    except Exception as e:
        return f"エラー: {str(e)}", 500

# 実行用
if __name__ == "__main__":
    app.run(debug=True, port=10000)
