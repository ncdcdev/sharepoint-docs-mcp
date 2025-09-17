# SharePoint Search MCP Server

SharePoint検索機能を提供するModel Context Protocol (MCP) サーバーです。stdioとHTTPの両方のトランスポートに対応しています。

## 機能

- **デュアルトランスポート対応**: stdio（デスクトップアプリ統合）とHTTP（ネットワークサービス）の両方をサポート
- **システム情報取得**: Pythonバージョンとプラットフォーム情報を取得するサンプルツール
- **適切なロギング**: stdioモードでstdout汚染を防ぐstderrベースのログ設定

## 必要要件

- Python 3.12以上
- uv (パッケージマネージャー)

## インストール

```bash
# リポジトリをクローン
git clone <repository-url>
cd sharepoint-search-mcp

# 依存関係をインストール
uv sync --dev
```

## 使用方法

### MCPサーバーの起動

**stdioモード（デスクトップアプリ統合用）:**
```bash
uv run sharepoint-search-mcp --transport stdio
```

**HTTPモード（ネットワークサービス用）:**
```bash
uv run sharepoint-search-mcp --transport http --host 127.0.0.1 --port 8000
```

**ヘルプの表示:**
```bash
uv run sharepoint-search-mcp --help
```

### 開発用コマンド

**コード品質チェック:**
```bash
# Lint（静的解析）
uv run lint

# 型チェック（ty）
uv run typecheck

# 全体チェック（型チェック + lint）
uv run check
```

**コードフォーマット:**
```bash
# フォーマットのみ
uv run fmt

# 自動修正 + フォーマット
uv run fix
```

## プロジェクト構造

```
sharepoint-search-mcp/
├── src/
│   ├── __init__.py
│   ├── server.py       # MCPサーバーのコアロジック
│   └── main.py         # CLIエントリポイント
├── scripts.py          # 開発用ユーティリティコマンド
├── pyproject.toml      # プロジェクト設定
└── README.md
```

## MCP Inspector での検証

### stdioモード
1. MCP Inspectorを開く
2. 「Command」を選択
3. Command: `uv`
4. Arguments: `run,sharepoint-search-mcp,--transport,stdio`
5. Working Directory: プロジェクトのルートディレクトリ
6. 「Connect」をクリック

### HTTPモード
1. サーバーを起動: `uv run sharepoint-search-mcp --transport http`
2. MCP Inspectorで「URL」を選択
3. URL: `http://127.0.0.1:8000/mcp/`
4. 「Connect」をクリック

## 開発

### コード品質ツール

- **ruff**: 高速なPythonリンター・フォーマッター
- **ty**: 高速型チェッカー（プレリリース版）

### 設定ファイル

- `pyproject.toml`: プロジェクト設定、依存関係、開発ツールの設定
- ruff設定: コードスタイル、ルール設定
- ty設定: 型チェックの詳細設定

## ライセンス

[ライセンスを記載]