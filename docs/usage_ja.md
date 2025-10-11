# 使用ガイド

このガイドでは、SharePoint MCPサーバーの各種クライアントでの使用方法と検索シナリオについて説明します。

## 目次

- [MCPサーバーの起動](#mcpサーバーの起動)
- [MCP Inspectorでの検証](#mcp-inspectorでの検証)
- [Claude Desktopとの統合](#claude-desktopとの統合)
- [検索の使用例](#検索の使用例)

## MCPサーバーの起動

### stdioモード（デスクトップアプリ統合用）
```bash
uv run sharepoint-docs-mcp --transport stdio
```

### HTTPモード（ネットワークサービス用）
```bash
uv run sharepoint-docs-mcp --transport http --host 127.0.0.1 --port 8000
```

### ヘルプの表示
```bash
uv run sharepoint-docs-mcp --help
```

## MCP Inspectorでの検証

### stdioモード
1. MCP Inspectorを開く
2. 「Command」を選択
3. Command: `uv`
4. Arguments: `run,sharepoint-docs-mcp,--transport,stdio`
5. Working Directory: プロジェクトのルートディレクトリ
6. 「Connect」をクリック

### HTTPモード
1. サーバーを起動: `uv run sharepoint-docs-mcp --transport http`
2. MCP Inspectorで「URL」を選択
3. URL: `http://127.0.0.1:8000/mcp/`
4. 「Connect」をクリック

## Claude Desktopとの統合

Claude Desktopと統合するには、設定ファイルを更新してください

- Windows: `%APPDATA%/Claude/claude_desktop_config.json`
- macOS: `~/Library/Application\ Support/Claude/claude_desktop_config.json`

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

## 検索の使用例

### SharePointサイトのみ検索

```bash
# 特定のSharePointサイトを検索
SHAREPOINT_SITE_NAME=team-site

# 複数のSharePointサイトを検索
SHAREPOINT_SITE_NAME=team-site,project-alpha,hr-docs
```

### OneDriveのみ検索

```bash
# 特定ユーザーのOneDrive全体を検索
SHAREPOINT_ONEDRIVE_PATHS=user1@company.com,user2@company.com
SHAREPOINT_SITE_NAME=@onedrive

# OneDrive内の特定フォルダーを検索
SHAREPOINT_ONEDRIVE_PATHS=manager@company.com:/Documents/重要書類,user@company.com:/Documents/プロジェクト
SHAREPOINT_SITE_NAME=@onedrive
```

### 混在検索（OneDrive + SharePoint）

```bash
# OneDriveとSharePointサイトを一緒に検索
SHAREPOINT_ONEDRIVE_PATHS=user1@company.com:/Documents/プロジェクト,manager@company.com:/Documents/重要書類
SHAREPOINT_SITE_NAME=@onedrive,team-site,project-alpha
```

### 一般的な使用例

**経営層向け設定**
```bash
# 経営陣のOneDriveフォルダーと取締役会文書を検索
SHAREPOINT_ONEDRIVE_PATHS=ceo@company.com:/Documents/経営資料,cfo@company.com:/Documents/財務
SHAREPOINT_SITE_NAME=@onedrive,executive-team,board-documents
```

**プロジェクトチーム向け設定**
```bash
# プロジェクトメンバーの作業フォルダーとチームサイトを検索
SHAREPOINT_ONEDRIVE_PATHS=pm@company.com:/Documents/ProjectA,dev@company.com:/Documents/ProjectA
SHAREPOINT_SITE_NAME=@onedrive,project-a-team,project-a-docs
```

**営業チーム向け設定**
```bash
# 営業担当のOneDriveフォルダーと顧客サイトを検索
SHAREPOINT_ONEDRIVE_PATHS=sales1@company.com:/Documents/顧客情報,sales2@company.com:/Documents/提案書
SHAREPOINT_SITE_NAME=@onedrive,sales-team,customer-portal
```
