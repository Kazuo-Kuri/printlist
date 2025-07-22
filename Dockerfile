FROM python:3.11-slim

# システムパッケージのアップデートと必要な依存追加（例: gcc, libffi など）
RUN apt-get update && apt-get install -y \
    build-essential \
    libffi-dev \
    libxml2-dev \
    libxslt1-dev \
    libjpeg-dev \
    zlib1g-dev \
    curl \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# 作業ディレクトリを設定
WORKDIR /app

# requirements.txt を先にコピー（キャッシュ活用のため）
COPY requirements.txt .

# 依存パッケージのインストール
RUN pip install --no-cache-dir -r requirements.txt

# アプリケーションファイルをコピー
COPY . .

# ポート番号の環境変数を設定（Render用）
ENV PORT=10000

# Flaskアプリをgunicornで起動
CMD ["gunicorn", "-w", "1", "-b", "0.0.0.0:10000", "app:app"]
