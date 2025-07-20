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
TEMPLATE_PATH = "printlist_form.xlsx"

def get_credentials():
    with open(CREDENTIAL_FILE_PATH, "r", encoding="utf-8") as f:
        credentials_dict = json.load(f)
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    return ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)

def extract_data(text):
    patterns = {
        "製造番号": r"製造番号[:：]\s*([^\s)]+)",
        "印刷番号": r"印刷番号[:：]\s*(.+?)(?:\n|$)",
        "製造日": r"製造日[:：]\s*(.+?)(?:\n|$)",
        "会社名": r"会社名[:：]\s*(.+?)(?:\n|$)",
        "製品名": r"製品名[:：]\s*(.+?)(?:\n|$)",
        "印刷データ": r"印刷データ[:：]\s*(.+?)(?:\n|$)"
    }
    results = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.DOTALL)
        if match:
            results[key] = match.group(1).strip()

    # 印刷データ分類：従来の〜という記載があれば「リピート」
    raw = results.get("印刷データ", "")
    if "従来の" in raw:
        results["印刷データ"] = "リピート"
    elif raw:
        results["印刷データ"] = "新規"
    else:
        results["印刷データ"] = ""

    # ファイル名（印刷用データ）のFMT部分のみ抽出
    fmt_match = re.findall(r"<印刷用データ\(\.FMT\)>.*?ファイル名[：:\s]+(.+?\.FMT)", text, re.DOTALL)
    if fmt_match:
        results["ファイル名"] = fmt_match[-1].strip()
    else:
        results["ファイル名"] = ""

    return results

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        text = request.form["text"]
        extracted_data = extract_data(text)

        # Excelテンプレート読み込み
        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active

        # セルへの書き込みマッピング
        cell_map = {
            "製造日": "B2",
            "製造番号": "E1",
            "印刷番号": "E2",
            "会社名": "B3",
            "製品名": "B4"
        }

        for key, cell in cell_map.items():
            value = extracted_data.get(key)
            if value:
                ws[cell] = value

        # 保存
        stream = io.BytesIO()
        wb.save(stream)
        stream.seek(0)

        return send_file(
            stream,
            as_attachment=True,
            download_name="printlist_output.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    return render_template("index.html")
