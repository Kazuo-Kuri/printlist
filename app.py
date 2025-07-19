from flask import Flask, render_template, request, send_file
import io
import openpyxl
import re

app = Flask(__name__)

def extract_fields(text):
    patterns = {
        "製造番号": r"製造番号[:：]\s*(.+)",
        "会社名": r"会社名[:：]\s*(.+)",
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
        results[key] = match.group(1) if match else ""
    return results

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        text = request.form['text']
        fields = extract_fields(text)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(list(fields.keys()))
        ws.append(list(fields.values()))

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
