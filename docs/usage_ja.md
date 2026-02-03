# 使用ガイド

このガイドでは、SharePoint MCPサーバーの各種クライアントでの使用方法と検索シナリオについて説明します。

## 目次

- [MCPサーバーの起動](#mcpサーバーの起動)
- [MCP Inspectorでの検証](#mcp-inspectorでの検証)
- [Claude Desktopとの統合](#claude-desktopとの統合)
- [検索の使用例](#検索の使用例)
- [Excel操作の使用例](#excel操作の使用例)

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

## Excel操作の使用例

SharePoint上のExcelファイルに対して、シート一覧取得、シート画像取得、セル範囲データ取得の操作を実行できます。

### 前提条件

- SharePoint Excel Services が有効化されていること
- 対象のExcelファイルがSharePointライブラリに保存されていること
- 適切なアクセス権限があること

### 基本的なワークフロー

1. **Excelファイルを検索**
```python
# sharepoint_docs_search ツールを使用
results = sharepoint_docs_search(
    query="予算",
    file_extensions=["xlsx"]
)
# 結果から file_path を取得
file_path = results[0]["path"]
# 例: "/sites/finance/Shared Documents/budget_2024.xlsx"
```

2. **シート一覧を取得**
```python
# sharepoint_excel_operations ツールを使用
sheets_xml = sharepoint_excel_operations(
    operation="list_sheets",
    file_path=file_path
)
# XML形式でシート一覧が返される
# 結果から必要なシート名を特定
```

3. **シートの画像を取得**
```python
# 特定のシートのビジュアルプレビューを取得
image_base64 = sharepoint_excel_operations(
    operation="get_image",
    file_path=file_path,
    sheet_name="Sheet1"
)
# base64エンコードされた画像データが返される
# 画像として保存または表示が可能
```

4. **セル範囲のデータを取得**
```python
# 特定のセル範囲のデータを取得
range_xml = sharepoint_excel_operations(
    operation="get_range",
    file_path=file_path,
    range_spec="Sheet1!A1:D10"
)
# XML形式でセルデータが返される
# データ分析やレポート生成に使用可能
```

### 操作タイプ

#### list_sheets
シート一覧をXML形式で取得します。

**パラメータ:**
- `operation`: "list_sheets"
- `file_path`: Excelファイルのパス（検索結果から取得）

**戻り値:** XML形式のシート一覧

**使用例:**
```python
sheets = sharepoint_excel_operations(
    operation="list_sheets",
    file_path="/sites/team/Shared Documents/report.xlsx"
)
```

#### get_image
指定したシートのキャプチャ画像をbase64形式で取得します。

**パラメータ:**
- `operation`: "get_image"
- `file_path`: Excelファイルのパス
- `sheet_name`: シート名（必須）

**戻り値:** base64エンコードされた画像データ（PNG形式）

**使用例:**
```python
image = sharepoint_excel_operations(
    operation="get_image",
    file_path="/sites/team/Shared Documents/report.xlsx",
    sheet_name="Summary"
)
# 画像として保存
import base64
with open("sheet_preview.png", "wb") as f:
    f.write(base64.b64decode(image))
```

#### get_range
指定したセル範囲のデータをXML形式で取得します。

**パラメータ:**
- `operation`: "get_range"
- `file_path`: Excelファイルのパス
- `range_spec`: セル範囲（必須、例: "Sheet1!A1:C10"）

**戻り値:** XML形式のセルデータ

**使用例:**
```python
data = sharepoint_excel_operations(
    operation="get_range",
    file_path="/sites/team/Shared Documents/report.xlsx",
    range_spec="Sheet1!A1:E20"
)
```

### 特殊文字の扱い

シート名やセル範囲にシングルクォート（'）が含まれる場合、自動的にエスケープされます。

**例:**
```python
# シート名に「John's Report」を指定
image = sharepoint_excel_operations(
    operation="get_image",
    file_path=file_path,
    sheet_name="John's Report"  # 自動的に "John''s Report" にエスケープ
)
```

### 一般的な使用例

**予算データの分析**
```python
# 1. 予算ファイルを検索
results = sharepoint_docs_search(query="予算 2024", file_extensions=["xlsx"])
file_path = results[0]["path"]

# 2. シート一覧を確認
sheets = sharepoint_excel_operations(operation="list_sheets", file_path=file_path)

# 3. 予算データを取得
budget_data = sharepoint_excel_operations(
    operation="get_range",
    file_path=file_path,
    range_spec="予算!A1:F100"
)
```

**レポートのビジュアルプレビュー**
```python
# 1. レポートファイルを検索
results = sharepoint_docs_search(query="月次レポート", file_extensions=["xlsx"])
file_path = results[0]["path"]

# 2. サマリーシートの画像を取得
summary_image = sharepoint_excel_operations(
    operation="get_image",
    file_path=file_path,
    sheet_name="Summary"
)

# 3. 画像を保存または表示
import base64
with open("monthly_summary.png", "wb") as f:
    f.write(base64.b64decode(summary_image))
```
