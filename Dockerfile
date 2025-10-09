FROM ghcr.io/astral-sh/uv:python3.13-bookworm

# タイムゾーンを日本に設定
ENV TZ=Asia/Tokyo
# デフォルトのポート設定
ENV PORT=8000

# 作業ディレクトリを設定
WORKDIR /app

# プロジェクトファイルをコピー
COPY pyproject.toml uv.lock ./

# Pythonの依存関係をインストール
RUN uv sync --frozen --no-cache

# ソースコードをコピー
COPY src/ ./src/

# 環境変数PORTで指定されたポートを公開
EXPOSE ${PORT:-8000}

# 起動コマンド - HTTPモードでSharePoint MCPサーバーを起動
ENTRYPOINT ["sh", "-c", "uv run sharepoint-docs-mcp --transport http --host 0.0.0.0 --port ${PORT:-8000}"]