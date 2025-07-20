FROM python:3.11-slim

# 作業ディレクトリを設定
WORKDIR /app

# ファイル一式をコピー
COPY . .

# 依存パッケージのインストール
RUN pip install --no-cache-dir -r requirements.txt

# Flaskアプリを起動
ENV PORT=10000
CMD ["gunicorn", "-w", "1", "-b", "0.0.0.0:10000", "app:app"]
