from flask import Flask, render_template, request, send_file
import os
import io
import json
import re
from dotenv import load_dotenv
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from openpyxl import Workbook

load_dotenv()
app = Flask(__name__)

def get_credentials():
    credentials_dict = json.loads(os.getenv("GOOGLE_CREDENTIALS"))
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
            results[key] = match.group(1).strip()
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

        wb = Workbook()
        ws = wb.active
        for i, (k, v) in enumerate(extracted_data.items(), start=1):
            ws.cell(row=i, column=1, value=k)
            ws.cell(row=i, column=2, value=v)

        excel_stream = io.BytesIO()
        wb.save(excel_stream)
        excel_stream.seek(0)

        creds = get_credentials()
        client = gspread.authorize(creds)
        sheet = client.open_by_key(os.getenv("SPREADSHEET_ID")).worksheet(os.getenv("SHEET_NAME"))
        values = sheet.get_all_values()
        start_row = len(values) + 2
        for i, (k, v) in enumerate(extracted_data.items()):
            sheet.update_cell(start_row + i, 1, k)
            sheet.update_cell(start_row + i, 2, v)

        return send_file(excel_stream, as_attachment=True, download_name="output.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return render_template("index.html")

# 🔽 Render対応のFlask起動設定
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 10000))
    app.run(host='0.0.0.0', port=port)
