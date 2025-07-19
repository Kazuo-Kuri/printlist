from flask import Flask, render_template, request, send_file
import io
import re
from openpyxl import load_workbook

app = Flask(__name__)

# テンプレートExcelファイルのパス
TEMPLATE_PATH = "printlist_form.xlsx"

# テンプレート内の出力セルマッピング
FIELD_CELL_MAP = {
    "製造日": "B1",
    "製造番号": "E1",
    "印刷番号": "E2",
    "会社名": "B3",
    "製品名": "B4"
}

# テキストからフィールドを抽出
def extract_fields(text):
    patterns = {
        "製造番号": r"製造番号[:：]\s*(.+)",
        "会社名": r"会社名[:：]\s*(.+)",
        "製品名": r"製品名[:：]\s*(.+)",
        "製品種類": r"製品種類[:：]\s*(.+)",
        "製造日": r"製造日[:：]\s*(.+)",
        "製造個数": r"製造個数[:：]\s*(.+)",
        "製品番号": r"製品番号[:：]\s*(.+)",
        "印刷番号": r"印刷番号[:：]\s*(.+)",
        "外装包材": r"外装包材[:：]\s*(.+)",
        "表面印刷": r"表面印刷[:：]\s*(.+)",
        "印刷データ": r"印刷データ[:：]\s*(.+)"
    }
    results = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            results[key] = match.group(1).strip()
    return results

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        text = request.form['text']
        extracted = extract_fields(text)

        # テンプレートを読み込み
        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active

        # 該当するセルに書き込む
        for field, value in extracted.items():
            if field in FIELD_CELL_MAP:
                ws[FIELD_CELL_MAP[field]] = value

        # Excelをメモリに保存して送信
        excel_stream = io.BytesIO()
        wb.save(excel_stream)
        excel_stream.seek(0)
        return send_file(
            excel_stream,
            as_attachment=True,
            download_name='extracted.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    return render_template('index.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
