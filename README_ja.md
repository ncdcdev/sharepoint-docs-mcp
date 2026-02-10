# SharePoint Docs MCP Server

> [🇬🇧 English version](README.md)

SharePointドキュメント検索機能を提供するModel Context Protocol (MCP) サーバーです。
stdioとHTTPの両方のトランスポートに対応しています。

## 認証方式

2つの認証方式をサポートしています

- **証明書認証**（アプリケーション権限）
  - Azure AD証明書ベース認証を使用
  - stdioとHTTPの両方のトランスポートに対応
  - サーバーアプリケーションや自動化に推奨
- **OAuth認証**（ユーザー権限）
  - OAuth 2.0 Authorization Code Flow with PKCEを使用
  - HTTPトランスポート専用（ブラウザベース認証が必要）
  - ユーザー委任アクセスシナリオに推奨

## 機能

### SharePoint機能

- **sharepoint_docs_search**
  - キーワードによるドキュメント検索
  - SharePointサイトとOneDriveの両方に対応
  - 複数検索対象（サイト、OneDriveフォルダー、混在）のサポート
  - ファイル拡張子フィルタリング（pdf、docx、xlsx等）
  - レスポンス形式オプション（詳細/簡潔）でトークン効率を改善
- **sharepoint_docs_download**
  - 検索結果からファイルをダウンロード
  - SharePoint/OneDriveファイルに応じた自動メソッド選択
- **sharepoint_excel**
  - SharePoint上のExcelファイルの読み取りと検索
  - **検索モード**: `query`パラメータで特定テキストを含むセルを検索
    - **複数キーワードOR検索**: カンマ区切りでキーワード指定（例: `"予算,見積"`）
    - **行コンテキスト**: `include_surrounding_cells=True`で行全体のデータを取得（API呼び出しをN+1から1に削減）
  - **読み取りモード**: `sheet`と`cell_range`パラメータで特定シート/範囲を取得
  - **ヘッダー自動追加**: `cell_range`指定時、デフォルトで固定行（ヘッダー）を自動的に含める
    - `include_frozen_rows=False`を指定すると、指定範囲のみを取得
    - `frozen_rows=0`のシートでは、`expand_axis_range=True`で1行目（列の場合）またはA列（行の場合）から自動取得
  - **セルスタイル情報**（オプション）: `include_cell_styles=True`を指定すると、背景色・列幅・行高さを取得
    - デフォルトは`False`でトークン消費を最小化
    - 強調表示されたセル、色付きヘッダー、視覚的に強調されたコンテンツの識別に便利
  - レスポンスには`rows`内のセルデータ（値と座標）と構造情報（利用可能な場合）を含む
  - 構造情報: シート名、dimensions、frozen_rows、frozen_cols、freeze_panes（存在する場合）、merged_ranges（結合セルが存在する場合）
  - Excel Services不要 - 直接ファイルダウンロード+openpyxl解析方式

### OneDrive対応

SharePointサイトとOneDriveコンテンツの両方を柔軟な設定で検索できます

- OneDrive統合: 特定ユーザーのOneDriveコンテンツを検索
- フォルダーレベルの対象指定: OneDrive内の特定フォルダーを検索
- 混在検索: SharePointサイトとOneDriveを1つの検索で組み合わせ
- 柔軟な設定: シンプルな環境変数による設定

## 必要要件

- Python 3.12
- uv (パッケージマネージャー)

## クイックスタート

### 1. インストール

```bash
# GitHubから直接実行（クローン不要）
uvx --from git+https://github.com/ncdcdev/sharepoint-docs-mcp sharepoint-docs-mcp --transport stdio
```

### 2. 設定

`.env`ファイルを作成して、SharePoint認証情報を設定します：

```bash
# 基本設定
SHAREPOINT_BASE_URL=https://yourcompany.sharepoint.com
SHAREPOINT_TENANT_ID=your-tenant-id-here
SHAREPOINT_CLIENT_ID=your-client-id-here
SHAREPOINT_SITE_NAME=yoursite

# 証明書認証の場合
SHAREPOINT_CERTIFICATE_PATH=path/to/certificate.pem
SHAREPOINT_PRIVATE_KEY_PATH=path/to/private_key.pem

# OAuth認証の場合（HTTPトランスポート専用）
# SHAREPOINT_AUTH_MODE=oauth
# SHAREPOINT_OAUTH_CLIENT_SECRET=your-oauth-client-secret-here
# SHAREPOINT_OAUTH_SERVER_BASE_URL=https://your-server.com
# 未設定: すべてのURI許可（開発環境）。設定時: 指定パターンのみ許可（本番環境推奨）
# SHAREPOINT_OAUTH_ALLOWED_REDIRECT_URIS=https://claude.ai/*,https://*.anthropic.com/*
```

詳細な設定手順は[セットアップガイド](docs/setup_ja.md)をご覧ください。

### 3. サーバーの起動

```bash
# stdioモード（Claude Desktop用）
uv run sharepoint-docs-mcp --transport stdio

# HTTPモード（ネットワークサービス用）
uv run sharepoint-docs-mcp --transport http --host 127.0.0.1 --port 8000
```

## ドキュメント

- 📘 [セットアップガイド](docs/setup_ja.md) - Azure ADと環境変数の詳細設定
- 📗 [使用ガイド](docs/usage_ja.md) - MCPクライアント統合と検索例
- 📙 [開発ガイド](docs/development_ja.md) - プロジェクト構造と開発コマンド
- 📕 [トラブルシューティングガイド](docs/troubleshooting_ja.md) - よくある問題とデバッグ

## ライセンス

MIT License - 詳細は[LICENSE](LICENSE)ファイルを参照してください。
