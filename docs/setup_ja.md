# SharePoint設定セットアップ

このガイドでは、MCPサーバーでSharePoint認証を設定するための詳細な手順を説明します。

## 目次

- [環境変数の設定](#環境変数の設定)
- [証明書の作成](#証明書の作成)
- [Azure ADアプリケーションの設定](#azure-adアプリケーションの設定)
- [ツール説明文のカスタマイズ](#ツール説明文のカスタマイズ)

## 環境変数の設定

`.env`ファイルを作成し、以下の設定を行います（`.env.example`を参考）：

### 共通設定（両方の認証方式共通）

```bash
# SharePoint設定
SHAREPOINT_BASE_URL=https://yourcompany.sharepoint.com
SHAREPOINT_TENANT_ID=your-tenant-id-here

# 認証モード（"certificate" または "oauth"）
# デフォルト: certificate
SHAREPOINT_AUTH_MODE=certificate

# 検索対象（複数指定可、カンマ区切り）
# オプション:
#   - @onedrive: OneDriveを検索に含める（SHAREPOINT_ONEDRIVE_PATHSが必要）
#   - @all: テナント全体を検索（セキュリティ上推奨されません）
#   - site-name: 特定のSharePointサイト名
# 例:
#   - 単一サイト: SHAREPOINT_SITE_NAME=team-site
#   - 複数サイト: SHAREPOINT_SITE_NAME=team-site,project-alpha,hr-docs
#   - OneDriveのみ: SHAREPOINT_SITE_NAME=@onedrive
#   - 混在: SHAREPOINT_SITE_NAME=@onedrive,team-site,project-alpha
SHAREPOINT_SITE_NAME=yoursite

# OneDrive設定（オプション）
# 形式: user@domain.com[:/folder/path][,user2@domain.com[:/folder/path]]...
# 例:
# SHAREPOINT_ONEDRIVE_PATHS=user@company.com,manager@company.com:/Documents/重要書類
# SHAREPOINT_ONEDRIVE_PATHS=user1@company.com:/Documents/プロジェクト,user2@company.com:/Documents/アーカイブ

# 検索設定（オプション）
SHAREPOINT_DEFAULT_MAX_RESULTS=20
SHAREPOINT_ALLOWED_FILE_EXTENSIONS=pdf,docx,xlsx,pptx,txt,md

# ツール説明文のカスタマイズ（オプション）
# SHAREPOINT_SEARCH_TOOL_DESCRIPTION=社内文書を検索します
# SHAREPOINT_DOWNLOAD_TOOL_DESCRIPTION=検索結果からファイルをダウンロードします
```

### 証明書認証設定（SHAREPOINT_AUTH_MODE=certificate）

```bash
# 証明書認証用クライアントID
SHAREPOINT_CLIENT_ID=your-client-id-here

# 証明書認証設定（ファイルパスまたはテキストのいずれかを指定）
# 優先順位: 1. テキスト、2. ファイルパス

# ファイルパスで指定する場合
SHAREPOINT_CERTIFICATE_PATH=path/to/your/certificate.pem
SHAREPOINT_PRIVATE_KEY_PATH=path/to/your/private_key.pem

# または、テキストで直接指定する場合（Cloud Run等での利用）
# テキストが設定されている場合、ファイルパスより優先されます
# SHAREPOINT_CERTIFICATE_TEXT="-----BEGIN CERTIFICATE-----\n...\n-----END CERTIFICATE-----"
# SHAREPOINT_PRIVATE_KEY_TEXT="-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----"
```

### OAuth認証設定（SHAREPOINT_AUTH_MODE=oauth）

**注**: OAuth認証はHTTPトランスポート（`--transport http`）が必要です

```bash
# OAuthクライアントID（Azure ADアプリの登録で取得）
# 未設定の場合は SHAREPOINT_CLIENT_ID にフォールバックします
# 通常は SHAREPOINT_CLIENT_ID のみ設定すれば両方の認証モードで使用できます
SHAREPOINT_OAUTH_CLIENT_ID=your-oauth-client-id-here

# OAuthクライアントシークレット（Azure ADアプリの「証明書とシークレット」で作成）
# OAuth認証モードでは必須
SHAREPOINT_OAUTH_CLIENT_SECRET=your-oauth-client-secret-here

# FastMCPサーバーのベースURL（OAuthコールバック用）
# Azure ADのリダイレクトURIは {SERVER_BASE_URL}/auth/callback になります
# デフォルト: http://localhost:8000
SHAREPOINT_OAUTH_SERVER_BASE_URL=http://localhost:8000

# 許可するMCPクライアントのリダイレクトURI（カンマ区切り、ワイルドカード対応）
# 未設定の場合: すべてのリダイレクトURIを許可（開発環境向け、本番環境では非推奨）
# 設定した場合: 指定されたパターンのみ許可（本番環境推奨）
# ローカル開発用:
# SHAREPOINT_OAUTH_ALLOWED_REDIRECT_URIS=http://localhost:*,http://127.0.0.1:*
# 本番環境用（例: Claude.ai統合）:
# SHAREPOINT_OAUTH_ALLOWED_REDIRECT_URIS=https://claude.ai/*,https://*.anthropic.com/*
```

## 証明書の作成

証明書ベース認証用の自己署名証明書を作成します：

```bash
mkdir -p cert
openssl genrsa -out cert/private_key.pem 2048
openssl req -new -key cert/private_key.pem -out cert/certificate.csr -subj "/CN=SharePointAuth"
openssl x509 -req -in cert/certificate.csr -signkey cert/private_key.pem -out cert/certificate.pem -days 365
rm cert/certificate.csr
```

作成されるファイル
- `cert/certificate.pem` - 公開証明書（Azure ADにアップロード）
- `cert/private_key.pem` - 秘密鍵（サーバーで使用）

## Azure ADアプリケーションの設定

認証方式に応じて、適切な設定を選択してください：

### オプションA: 証明書認証の設定（アプリケーション権限）

**1. Azure ADアプリケーションの登録**
1. [Azure Portal](https://portal.azure.com/) → EntraID → アプリの登録
2. 「新規登録」をクリック
3. アプリケーション名を入力（例: SharePoint MCP Server）
4. 登録ボタンをクリック

**2. 証明書のアップロード**
1. 作成したアプリを選択 → 「証明書とシークレット」
2. 「証明書」タブで「証明書のアップロード」をクリック
3. 作成した `cert/certificate.pem` をアップロード

**3. API権限の設定**
1. 「API権限」タブに移動
2. 「権限の追加」→「Microsoft Graph」→「アプリケーションの権限」
3. 以下の権限を追加
   - `Sites.FullControl.All` - SharePointサイトへのフルアクセス
4. 「管理者の同意を与える」をクリック

**4. 必要な情報の取得**
- テナントID: 「概要」ページのディレクトリ（テナント）ID
- クライアントID: 「概要」ページのアプリケーション（クライアント）ID

### オプションB: OAuth認証の設定（ユーザー権限）

**1. Azure ADアプリケーションの登録**
1. [Azure Portal](https://portal.azure.com/) → EntraID → アプリの登録
2. 「新規登録」をクリック
3. 以下を入力：
   - 名前: SharePoint MCP OAuth Client
   - サポートされているアカウントの種類: この組織ディレクトリのみのアカウント
   - リダイレクトURI: Web - `http://localhost:8000/auth/callback`
4. 「登録」をクリック

**2. クライアントシークレットの設定**
1. 作成したアプリを選択 → 「証明書とシークレット」
2. 「新しいクライアントシークレット」をクリック
3. 説明を追加（例: MCP Server Secret）
4. 有効期限を設定（例: 24ヶ月）
5. 「追加」をクリック
6. **重要**: シークレット値をすぐにコピー（再表示されません）
7. この値を `SHAREPOINT_OAUTH_CLIENT_SECRET` 環境変数に保存

**3. 認証の設定**
1. 作成したアプリを選択 → 「認証」
2. 「プラットフォームの構成」で、リダイレクトURIが `http://localhost:8000/auth/callback` に設定されていることを確認
3. 「詳細設定」で：
   - パブリッククライアントフローを許可: いいえ
4. 変更を保存

**4. API権限の設定（委任された権限）**
1. 「API権限」タブに移動
2. 「権限の追加」→「SharePoint」→「委任された権限」
3. 以下の権限を追加：
   - `AllSites.Read` - すべてのサイトコレクション内のアイテムを読み取る
   - `AllSites.Write` - すべてのサイトコレクション内のアイテムを読み書きする（ファイルダウンロードに必要な場合）
   - `User.Read` - ユーザープロファイルを読み取る（自動的に追加）
4. 「管理者の同意を与える」をクリック（管理者の同意が必要）

**5. 必要な情報の取得**
- テナントID: 「概要」ページのディレクトリ（テナント）ID
- OAuthクライアントID: 「概要」ページのアプリケーション（クライアント）ID
- OAuthクライアントシークレット: 手順2で取得したシークレット値

**6. 認証フロー**

このMCPサーバーのOAuth認証は**FastMCPのOIDCProxy**によって処理され、安全な2層認証を実装します

1. **レイヤー1 - MCPクライアント認証**
   - MCPクライアント（Claude Desktop、MCP Inspectorなど）がFastMCPサーバーに接続
   - FastMCPプロキシがユーザーをMicrosoft Entra IDで認証
   - クライアント側でシークレットが不要なPKCEを使用してセキュリティを確保

2. **レイヤー2 - SharePoint APIアクセス**
   - 認証されたユーザーのトークンを使用してSharePoint APIにアクセス
   - ユーザーの委任された権限を使用（AllSites.Read/Write）

**セキュリティ機能**
- PKCEによりトークンの傍受を防止
- クライアントシークレットはサーバー側でのみ保存
- トークン検証はAzure ADのOAuthフローを信頼

**重要な注意事項**
- 認証はMCPクライアントのOAuthフローを通じて実行されます
- 手動でブラウザログインする必要はありません - MCPクライアントがOAuthフローを自動処理
- トークンはFastMCPによって管理され、安全にキャッシュされます
- サーバーは `/auth/callback` エンドポイント（FastMCP標準）をOAuthコールバックに使用
- MCPクライアントは動的ポート（例: http://localhost:6274/oauth/callback ）を使用可能で、FastMCPはワイルドカードlocalhost URIを許可

## ツール説明文のカスタマイズ

MCPツールの説明文を日本語などにカスタマイズできます

- `SHAREPOINT_SEARCH_TOOL_DESCRIPTION`: 検索ツールの説明文（デフォルト: "Search for documents in SharePoint"）
- `SHAREPOINT_DOWNLOAD_TOOL_DESCRIPTION`: ダウンロードツールの説明文（デフォルト: "Download a file from SharePoint"）

例：
```bash
SHAREPOINT_SEARCH_TOOL_DESCRIPTION=社内文書を検索します
SHAREPOINT_DOWNLOAD_TOOL_DESCRIPTION=検索結果からファイルをダウンロードします
```
