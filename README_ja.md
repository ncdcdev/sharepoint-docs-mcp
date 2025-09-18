# SharePoint Docs MCP Server

SharePointドキュメント検索機能を提供するModel Context Protocol (MCP) サーバーです。
stdioとHTTPの両方のトランスポートに対応しています。

認証はAzure ADの証明書ベース認証のみをサポートしています。
その他の認証方式には対応していないのでご注意ください。

## 機能

- SharePoint検索
  - 証明書認証によるSharePointドキュメント検索
- 証明書認証
  - Azure AD証明書ベース認証をサポート
- デュアルトランスポート対応
  - stdio（デスクトップアプリ統合）とHTTP（ネットワークサービス）の両方をサポート
- 適切なロギング
  - stdioモードでstdout汚染を防ぐstderrベースのログ設定

### SharePoint機能

- sharepoint_docs_search
  - キーワードによるドキュメント検索
  - レスポンス形式オプション（詳細/簡潔）でトークン効率を改善
- sharepoint_docs_download
  - 検索結果からファイルをダウンロード

## 必要要件

- Python 3.12以上
- uv (パッケージマネージャー)

## インストール

### 方法1: uvxで直接実行（推奨）

```bash
# GitHubから直接実行（クローン不要）
uvx --from git+https://github.com/ncdcdev/sharepoint-docs-mcp sharepoint-docs-mcp --transport stdio

# HTTPモードの場合
uvx --from git+https://github.com/ncdcdev/sharepoint-docs-mcp sharepoint-docs-mcp --transport http --host 127.0.0.1 --port 8000
```

### 方法2: 開発環境セットアップ

```bash
# リポジトリをクローン
git clone https://github.com/ncdcdev/sharepoint-docs-mcp
cd sharepoint-docs-mcp

# 依存関係をインストール
uv sync --dev
```

## SharePoint設定

### 1. 環境変数の設定

`.env`ファイルを作成し、以下の設定を行います（`.env.example`を参考）：

```bash
# SharePoint設定
SHAREPOINT_BASE_URL=https://yourcompany.sharepoint.com
SHAREPOINT_SITE_NAME=yoursite
SHAREPOINT_TENANT_ID=your-tenant-id-here
SHAREPOINT_CLIENT_ID=your-client-id-here

# SHAREPOINT_SITE_NAME を空にするとテナント全体を検索対象にできます
# SHAREPOINT_SITE_NAME=

# 証明書認証設定（ファイルパスまたはテキストのいずれかを指定）
# 優先順位: 1. テキスト、2. ファイルパス

# ファイルパスで指定する場合
SHAREPOINT_CERTIFICATE_PATH=path/to/your/certificate.pem
SHAREPOINT_PRIVATE_KEY_PATH=path/to/your/private_key.pem

# または、テキストで直接指定する場合（Cloud Run等での利用）
# テキストが設定されている場合、ファイルパスより優先されます
# SHAREPOINT_CERTIFICATE_TEXT="-----BEGIN CERTIFICATE-----\n...\n-----END CERTIFICATE-----"
# SHAREPOINT_PRIVATE_KEY_TEXT="-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----"

# 検索設定（オプション）
SHAREPOINT_DEFAULT_MAX_RESULTS=20
SHAREPOINT_ALLOWED_FILE_EXTENSIONS=pdf,docx,xlsx,pptx,txt,md

# ツール説明文のカスタマイズ（オプション）
# SHAREPOINT_SEARCH_TOOL_DESCRIPTION=社内文書を検索します
# SHAREPOINT_DOWNLOAD_TOOL_DESCRIPTION=検索結果からファイルをダウンロードします
```

### 2. 証明書の作成

証明書ベース認証用の自己署名証明書を作成します：

```bash
mkdir -p cert
openssl genrsa -out cert/private_key.pem 2048
openssl req -new -key cert/private_key.pem -out cert/certificate.csr -subj "/CN=SharePointAuth"
openssl x509 -req -in cert/certificate.csr -signkey cert/private_key.pem -out cert/certificate.pem -days 365
rm cert/certificate.csr
```

作成されるファイル
- `cert/certificate.pem`
  - 公開証明書（Azure ADにアップロード）
- `cert/private_key.pem`
  - 秘密鍵（サーバーで使用）

### 3. Azure AD証明書認証の設定

#### 1. Azure ADアプリケーションの登録
1. [Azure Portal](https://portal.azure.com/) → EntraID → アプリの登録
2. 「新規登録」をクリック
3. アプリケーション名を入力（例: SharePoint MCP Server）
4. 登録ボタンをクリック

#### 2. 証明書のアップロード
1. 作成したアプリを選択 → 「証明書とシークレット」
2. 「証明書」タブで「証明書のアップロード」をクリック
3. 作成した `cert/certificate.pem` をアップロード

#### 3. API権限の設定
1. 「API権限」タブに移動
2. 「権限の追加」→「Microsoft Graph」→「アプリケーションの権限」
3. 以下の権限を追加
   - `Sites.FullControl.All`
     - SharePointサイトへのフルアクセス
4. 「管理者の同意を与える」をクリック

#### 4. 必要な情報の取得
- テナントID
  - 「概要」ページのディレクトリ（テナント）ID
- クライアントID
  - 「概要」ページのアプリケーション（クライアント）ID

### 4. ツール説明文のカスタマイズ（オプション）

MCPツールの説明文を日本語などにカスタマイズできます：

- `SHAREPOINT_SEARCH_TOOL_DESCRIPTION`: 検索ツールの説明文（デフォルト: "Search for documents in SharePoint"）
- `SHAREPOINT_DOWNLOAD_TOOL_DESCRIPTION`: ダウンロードツールの説明文（デフォルト: "Download a file from SharePoint"）

例：
```bash
SHAREPOINT_SEARCH_TOOL_DESCRIPTION=社内文書を検索します
SHAREPOINT_DOWNLOAD_TOOL_DESCRIPTION=検索結果からファイルをダウンロードします
```

## 使用方法

### MCPサーバーの起動

**stdioモード（デスクトップアプリ統合用）**
```bash
uv run sharepoint-docs-mcp --transport stdio
```

**HTTPモード（ネットワークサービス用）**
```bash
uv run sharepoint-docs-mcp --transport http --host 127.0.0.1 --port 8000
```

**ヘルプの表示**
```bash
uv run sharepoint-docs-mcp --help
```

### MCP Inspector での検証

**stdioモード**
1. MCP Inspectorを開く
2. 「Command」を選択
3. Command: `uv`
4. Arguments: `run,sharepoint-docs-mcp,--transport,stdio`
5. Working Directory: プロジェクトのルートディレクトリ
6. 「Connect」をクリック

**HTTPモード**
1. サーバーを起動: `uv run sharepoint-docs-mcp --transport http`
2. MCP Inspectorで「URL」を選択
3. URL: `http://127.0.0.1:8000/mcp/`
4. 「Connect」をクリック

### 開発用コマンド

**テスト**
```bash
# テスト実行
uv run test

# カバレッジレポート付きテスト実行
uv run test --cov=src --cov-report=html
```

**コード品質チェック**
```bash
# Lint（静的解析）
uv run lint

# 型チェック（ty）
uv run typecheck

# 全体チェック（型チェック + lint + テスト）
uv run check
```

**コードフォーマット**
```bash
# フォーマットのみ
uv run fmt

# 自動修正 + フォーマット
uv run fix
```

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
├── scripts.py               # 開発用ユーティリティコマンド
├── pyproject.toml           # プロジェクト設定と依存関係
├── README.md                # 英語ドキュメント
└── README_ja.md             # 日本語ドキュメント
```

## Claude Desktop との統合

Claude Desktopと統合するには、設定ファイルを更新してください

- Windows
  - `%APPDATA%/Claude/claude_desktop_config.json`
- macOS
  - `~/Library/Application\ Support/Claude/claude_desktop_config.json`

### 設定例1: 環境変数を直接指定

```json
{
  "mcpServers": {
    "sharepoint-docs": {
      "command": "uv",
      "args": ["run", "sharepoint-docs-mcp", "--transport", "stdio"],
      "cwd": "/path/to/sharepoint-docs-mcp",
      "env": {
        "SHAREPOINT_BASE_URL": "https://yourcompany.sharepoint.com",
        "SHAREPOINT_SITE_NAME": "yoursite",
        "SHAREPOINT_TENANT_ID": "your-tenant-id-here",
        "SHAREPOINT_CLIENT_ID": "your-client-id-here",
        "SHAREPOINT_CERTIFICATE_PATH": "./cert/certificate.pem",
        "SHAREPOINT_PRIVATE_KEY_PATH": "./cert/private_key.pem"
      }
    }
  }
}
```

### 設定例2: .envファイルを使用（推奨）

```json
{
  "mcpServers": {
    "sharepoint-docs": {
      "command": "uv",
      "args": ["run", "sharepoint-docs-mcp", "--transport", "stdio"],
      "cwd": "/path/to/sharepoint-docs-mcp"
    }
  }
}
```

この場合、プロジェクトルートの`.env`ファイルに設定を記載します。

### 設定例3: uvxを使用（クローン不要）

```json
{
  "mcpServers": {
    "sharepoint-docs": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/ncdcdev/sharepoint-docs-mcp", "sharepoint-docs-mcp", "--transport", "stdio"],
      "env": {
        "SHAREPOINT_BASE_URL": "https://yourcompany.sharepoint.com",
        "SHAREPOINT_SITE_NAME": "yoursite",
        "SHAREPOINT_TENANT_ID": "your-tenant-id-here",
        "SHAREPOINT_CLIENT_ID": "your-client-id-here",
        "SHAREPOINT_CERTIFICATE_PATH": "/path/to/certificate.pem",
        "SHAREPOINT_PRIVATE_KEY_PATH": "/path/to/private_key.pem"
      }
    }
  }
}
```

この設定では、リポジトリをローカルにクローンすることなく、GitHubから直接MCPサーバーを実行できます。`SHAREPOINT_CERTIFICATE_PATH`と`SHAREPOINT_PRIVATE_KEY_PATH`には、ファイルへの絶対パスを指定する必要がある点にご注意ください。


## 開発

### テストフレームワーク

- **pytest**: フィクスチャとモック機能を持つPythonテストフレームワーク
- **pytest-cov**: コードカバレッジレポート
- **pytest-mock**: 強化されたモック機能
- 主要機能をカバーする24のユニットテスト（カバレッジ48%）

### コード品質ツール

- **ruff**: 高速なPythonリンター・フォーマッター
- **ty**: 高速型チェッカー（プレリリース版）

### 設定ファイル

- `pyproject.toml`: プロジェクト設定、依存関係、開発ツールの設定
- pytest設定: テスト発見とカバレッジ設定
- ruff設定: コードスタイル、ルール設定
- ty設定: 型チェックの詳細設定

## トラブルシューティング

### よくある問題

#### 1. 認証エラー
```
SharePoint configuration is invalid: SHAREPOINT_TENANT_ID is required
```
- `.env`ファイルが正しく設定されているか確認
- 環境変数が正しく読み込まれているか確認

#### 2. 証明書エラー
```
Certificate file not found: path/to/certificate.pem
```
- 証明書ファイルのパスが正しいか確認
- 証明書が正しく作成されているか確認
- ファイルの読み取り権限があるか確認

#### 3. API権限エラー
```
Access token request failed
```
- Azure ADアプリの権限設定を確認
- 管理者の同意が行われているか確認
- クライアントIDとテナントIDが正しいか確認

#### 4. 設定確認コマンド
```bash
# 設定ステータスを確認（MCP Inspector使用）
# get_sharepoint_config_status ツールを実行
```

### デバッグ方法

#### MCP Inspectorを使用
```bash
npx @modelcontextprotocol/inspector uv run sharepoint-docs-mcp --transport stdio
```

#### ログレベルの調整
サーバー起動時に詳細なログが出力されます。エラーの詳細は標準エラー出力に表示されます。

## ライセンス

MIT License - 詳細は[LICENSE](LICENSE)ファイルを参照してください。
