FROM python:3.11-slim

# 作業ディレクトリを設定
WORKDIR /app

# ファイル一式をコピー
COPY . .

# 依存パッケージのインストール
RUN pip install --no-cache-dir -r requirements.txt

# Flaskアプリを起動
ENV PORT=10000
CMD ["python", "app.py"]
