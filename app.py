from flask import Flask, render_template, request, send_file
import os
import io
import json
import re
from openpyxl import load_workbook
import gspread
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)

# Secret File 経由の認証ファイルパス
CREDENTIAL_FILE_PATH = "/etc/secrets/credentials.json"

def get_credentials():
    with open(CREDENTIAL_FILE_PATH, "r", encoding="utf-8") as f:
        credentials_dict = json.load(f)
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    return ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)

def extract_data(text):
    patterns = {
        "製造番号": r"製造番号[:：]\s*([^\s)]+)",
        "印刷番号": r"印刷番号[:：]\s*([^\s\n]+)",
        "製造日": r"製造日[:：]\s*(.+?)(?:\n|$)",
        "会社名": r"会社名[:：]\s*(.+?)(?:\n|$)",
        "製品名": r"製品名[:：]\s*(.+?)(?:\n|$)",
        "製品種類": r"製品種類[:：]\s*(.+?)(?:\n|$)",
        "外装包材": r"外装包材[:：]\s*(.+?)(?:\n|$)",
        "表面印刷": r"表面印刷[:：][^\n]+.*?表面印刷[:：]\s*(.+?)(?:\n|$)",
        "製造個数": r"製造個数[:：]\s*(.+?)(?:\n|$)",
        "ファイル名": r"<印刷用データ\(\.FMT\)>.*?ファイル名[:：]\s*(.+?)(?:\n|$)",
        "印刷データ（元）": r"印刷データ[:：]\s*(.+?)(?:\n|$)"
    }

    results = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.DOTALL)
        if match:
            results[key] = match.group(1).strip()

    # 「印刷データ」区分判定（"従来の" が含まれていればリピート）
    if "印刷データ（元）" in results:
        raw = results.pop("印刷データ（元）")
        results["印刷データ"] = "リピート" if "従来の" in raw else "新規"
    else:
        results["印刷データ"] = ""

    return results

@app.route("/", methods=["GET", "POST"])
def index():
    extracted_data = {}
    if request.method == "POST":
        text = request.form["text"]
        extracted_data = extract_data(text)

        # Excel テンプレート読込
        template_path = "printlist_form.xlsx"
        wb = load_workbook(template_path)
        ws = wb.active

        # セルマッピング
        cell_map = {
            "製造日": "B1",
            "製造番号": "E1",
            "印刷番号": "E2",
            "会社名": "B3",
            "製品名": "B4"
        }

        for key, cell in cell_map.items():
            if key in extracted_data:
                ws[cell] = extracted_data[key]

        # 出力ストリーム
        excel_stream = io.BytesIO()
        wb.save(excel_stream)
        excel_stream.seek(0)

        # Google スプレッドシート書き込み
        creds = get_credentials()
        client = gspread.authorize(creds)
        SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
        SHEET_NAME = os.getenv("SHEET_NAME")  # データ専用シート名
        sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)

        values = sheet.get_all_values()
        start_row = len(values) + 2
        for i, (k, v) in enumerate(extracted_data.items()):
            sheet.update_cell(start_row + i, 1, k)
            sheet.update_cell(start_row + i, 2, v)

        return send_file(
            excel_stream,
            as_attachment=True,
            download_name="output.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    return render_template("index.html")
