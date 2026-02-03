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
- **sharepoint_excel_operations**
  - SharePoint上のExcelファイルを操作
  - シート一覧の取得（XML形式）
  - シートのキャプチャ画像取得（base64形式）
  - セル範囲のデータ取得（XML形式）
  - SharePoint Excel Services REST APIを使用

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
