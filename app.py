from flask import Flask, render_template, request, send_file
import os
import io
import json
import re
from openpyxl import Workbook
import gspread
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)

# Render Secret File 経由の認証ファイル
CREDENTIAL_FILE_PATH = "/etc/secrets/credentials.json"

def get_credentials():
    with open(CREDENTIAL_FILE_PATH, "r", encoding="utf-8") as f:
        credentials_dict = json.load(f)
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    return ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)

def extract_data(text):
    patterns = {
        "製造番号": r"製造番号[:：]\s*(.+)",
        "印刷番号": r"印刷番号[:：]\s*(.+)",
        "製造日": r"製造日[:：]\s*(.+)",
        "会社名": r"会社名[:：]\s*(.+)",
        "製品名": r"製品名[:：]\s*(.+)",
        "製品種類": r"製品種類[:：]\s*(.+)",
        "外装包材": r"外装包材[:：]\s*(.+)",
        "表面印刷": r"表面印刷[:：][^\n]+.*?表面印刷[:：]\s*(.+)",
        "製造個数": r"製造個数[:：]\s*(.+)",
        "ファイル名": r"ファイル名[:：]\s*(.+)",
        "印刷データ（元）": r"印刷データ[:：]\s*(.+)"
    }

    results = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.DOTALL)
        if match:
            value = match.group(1).strip()
            # 空白行までで切り取り
            split_value = re.split(r"\n\s*\n", value)[0].strip()
            results[key] = split_value

    if "印刷データ（元）" in results:
        raw = results.pop("印刷データ（元）")
        results["印刷データ"] = "リピート" if "同じデータ" in raw else "新規"
    else:
        results["印刷データ"] = ""

    return results

@app.route("/", methods=["GET", "POST"])
def index():
    extracted_data = {}
    if request.method == "POST":
        text = request.form["text"]
        extracted_data = extract_data(text)

        # Excelファイル作成
        wb = Workbook()
        ws = wb.active
        for i, (k, v) in enumerate(extracted_data.items(), start=1):
            ws.cell(row=i, column=1, value=k)
            ws.cell(row=i, column=2, value=v)

        excel_stream = io.BytesIO()
        wb.save(excel_stream)
        excel_stream.seek(0)

        # Google スプレッドシートに書き込み
        creds = get_credentials()
        client = gspread.authorize(creds)
        SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
        SHEET_NAME = os.getenv("SHEET_NAME")
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
