# 開発ガイド

このガイドでは、プロジェクト構造、開発コマンド、コード品質ツールなど、開発に関する情報を説明します。

## 目次

- [プロジェクト構造](#プロジェクト構造)
- [開発用コマンド](#開発用コマンド)
- [テストフレームワーク](#テストフレームワーク)
- [コード品質ツール](#コード品質ツール)
- [デバッグ](#デバッグ)

## プロジェクト構造

```
sharepoint-docs-mcp/
├── src/
│   ├── __init__.py
│   ├── server.py            # MCPサーバーのコアロジック
│   ├── main.py              # CLIエントリポイント
│   ├── config.py            # 設定管理
│   ├── sharepoint_auth.py   # Azure AD認証
│   ├── sharepoint_search.py # SharePoint検索クライアント
│   └── error_messages.py    # エラーハンドリング
├── tests/
│   ├── __init__.py
│   ├── conftest.py          # テストフィクスチャとモック
│   ├── test_config.py       # 設定管理のテスト
│   ├── test_server.py       # サーバー機能のテスト
│   └── test_error_messages.py # エラーハンドリングのテスト
├── docs/                    # ドキュメントファイル
├── scripts.py               # 開発用ユーティリティコマンド
├── pyproject.toml           # プロジェクト設定と依存関係
├── README.md                # 英語ドキュメント
└── README_ja.md             # 日本語ドキュメント
```

## 開発用コマンド

### テスト

```bash
# テスト実行
uv run test

# カバレッジレポート付きテスト実行
uv run test --cov=src --cov-report=html
```

### コード品質チェック

```bash
# Lint（静的解析）
uv run lint

# 型チェック（ty）
uv run typecheck

# 全体チェック（型チェック + lint + テスト）
uv run check
```

### コードフォーマット

```bash
# フォーマットのみ
uv run fmt

# 自動修正 + フォーマット
uv run fix
```

## テストフレームワーク

- **pytest**: フィクスチャとモック機能を持つPythonテストフレームワーク
- **pytest-cov**: コードカバレッジレポート
- **pytest-mock**: 強化されたモック機能

## コード品質ツール

- **ruff**: 高速なPythonリンター・フォーマッター
- **ty**: 高速型チェッカー（プレリリース版）

### 設定ファイル

- `pyproject.toml`: プロジェクト設定、依存関係、開発ツールの設定
  - pytest設定: テスト発見とカバレッジ設定
  - ruff設定: コードスタイル、ルール設定
  - ty設定: 型チェックの詳細設定

## デバッグ

### MCP Inspectorを使用

```bash
npx @modelcontextprotocol/inspector uv run sharepoint-docs-mcp --transport stdio
```

### ログレベルの調整

サーバー起動時に詳細なログが出力されます。エラーの詳細は標準エラー出力に表示されます。
