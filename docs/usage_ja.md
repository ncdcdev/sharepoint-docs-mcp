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

`sharepoint_excel`ツールを使用して、SharePoint上のExcelファイルの読み取りと検索ができます。2つのモードをサポートしています：
- **検索モード**: 特定のコンテンツを検索してセル位置を特定（`query`パラメータを使用）
- **読み取りモード**: シート/範囲フィルタリングオプション付きでデータを取得

### 前提条件

- ExcelファイルがSharePointライブラリまたはOneDriveに保存されていること
- 適切なアクセス権限があること
- Excel Services不要

### ツールパラメータ

| パラメータ | 型 | デフォルト | 説明 |
|-----------|------|---------|-------------|
| `file_path` | str | 必須 | Excelファイルのパス |
| `query` | str \| None | None | 検索キーワード（検索モードを有効化） |
| `sheet` | str \| None | None | シート名（特定シートのみ取得） |
| `cell_range` | str \| None | None | セル範囲（例: "A1:D10"） |
| `include_formatting` | bool | False | 書式情報を含めるか |

### 基本的なワークフロー

**推奨: まず検索し、その後特定範囲を読み取る**

```python
# ステップ1: 関連コンテンツを検索
result = sharepoint_excel(file_path="/path/to/file.xlsx", query="合計")
# → "合計"がSheet1のセルC10にあることが分かる

# ステップ2: 周辺データを読み取る
data = sharepoint_excel(file_path="/path/to/file.xlsx", sheet="Sheet1", cell_range="A1:D15")
```

### 使用パターン

#### 1. 検索モード（queryパラメータ使用）
```python
# "予算"を含むセルを検索
result = sharepoint_excel(
    file_path="/sites/finance/Shared Documents/report.xlsx",
    query="予算"
)
```

**検索レスポンス:**
```json
{
  "file_path": "/sites/finance/Shared Documents/report.xlsx",
  "mode": "search",
  "query": "予算",
  "match_count": 3,
  "matches": [
    {"sheet": "Sheet1", "coordinate": "A1", "value": "予算報告"},
    {"sheet": "Sheet1", "coordinate": "B5", "value": "月間予算"},
    {"sheet": "Summary", "coordinate": "C3", "value": "予算合計"}
  ]
}
```

#### 2. 全データ取得（デフォルト）
```python
# 全シート・全データを取得
result = sharepoint_excel(
    file_path="/sites/finance/Shared Documents/report.xlsx"
)
```

#### 3. 特定シートの取得
```python
# 特定シートのデータのみ取得
result = sharepoint_excel(
    file_path="/sites/finance/Shared Documents/report.xlsx",
    sheet="Summary"
)
```

#### 4. 特定範囲の取得
```python
# シート内の特定範囲のデータを取得
result = sharepoint_excel(
    file_path="/sites/finance/Shared Documents/report.xlsx",
    sheet="Sheet1",
    cell_range="A1:D10"
)
```

#### 5. 書式情報を含む読み取り
```python
# 書式（色、結合セルなど）を含むデータを取得
result = sharepoint_excel(
    file_path="/sites/finance/Shared Documents/report.xlsx",
    sheet="Sheet1",
    include_formatting=True
)
```

### JSON出力形式

#### 読み取りモード（デフォルト）

```json
{
  "file_path": "/sites/test/Shared Documents/budget.xlsx",
  "sheets": [
    {
      "name": "Summary",
      "dimensions": "A1:E10",
      "rows": [
        [
          {"value": "部門", "coordinate": "A1"},
          {"value": 12500, "coordinate": "B1"}
        ]
      ]
    }
  ]
}
```

#### 範囲指定時の読み取りモード

```json
{
  "file_path": "/sites/test/Shared Documents/budget.xlsx",
  "sheets": [
    {
      "name": "Summary",
      "dimensions": "A1:E10",
      "requested_range": "A1:B2",
      "rows": [
        [
          {"value": "部門", "coordinate": "A1"},
          {"value": "予算", "coordinate": "B1"}
        ],
        [
          {"value": "営業", "coordinate": "A2"},
          {"value": 50000, "coordinate": "B2"}
        ]
      ]
    }
  ]
}
```

#### 書式情報を含む形式（include_formatting=true）

```json
{
  "file_path": "/sites/test/Shared Documents/budget.xlsx",
  "sheets": [
    {
      "name": "Summary",
      "dimensions": "A1:E10",
      "rows": [
        [
          {
            "value": "部門",
            "coordinate": "A1",
            "data_type": "s",
            "fill": {
              "pattern_type": "solid",
              "fg_color": "#CCCCCC",
              "bg_color": null
            },
            "merged": {
              "range": "A1:B1",
              "is_top_left": true
            },
            "width": 15.0,
            "height": 20.0
          }
        ]
      ]
    }
  ]
}
```

### 利用可能なセル情報

**デフォルト（常に含まれる）**
- **value**: セル値（文字列、数値、日付、数式など）
- **coordinate**: セル位置（例: "A1"、"B2"）

**include_formatting=true の場合**
- **data_type**: データ型コード（`s`=文字列、`n`=数値、`f`=数式など）
- **fill**: 塗りつぶし色情報（パターンタイプ、前景色/背景色）
- **merged**: 結合セル情報（範囲、位置）
- **width**: 列幅
- **height**: 行高

### 一般的な使用例

**予算データを検索して抽出**
```python
# 1. 予算ファイルを検索
results = sharepoint_docs_search(query="予算 2024", file_extensions=["xlsx"])
file_path = results[0]["path"]

# 2. 必要なデータを検索
search_result = sharepoint_excel(file_path=file_path, query="売上合計")
# → Sheet1:C15 で見つかった

# 3. 関連セクションを取得
data = sharepoint_excel(file_path=file_path, sheet="Sheet1", cell_range="A1:D20")
```

**セル書式の分析**
```python
# 書式情報を含むExcelデータを取得
json_data = sharepoint_excel(file_path=file_path, include_formatting=True)
data = json.loads(json_data)

# 特定の書式を持つセルを検索
for sheet in data["sheets"]:
    for row in sheet["rows"]:
        for cell in row:
            if cell.get("fill", {}).get("fg_color"):
                print(f"色付きセル {cell['coordinate']}: {cell['value']}")
```

**特定シートをCSVにエクスポート**
```python
# 特定シートのデータを取得
json_data = sharepoint_excel(file_path=file_path, sheet="Summary")
data = json.loads(json_data)

# CSVに変換
import csv
sheet = data["sheets"][0]
with open(f"{sheet['name']}.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    for row in sheet["rows"]:
        values = [cell["value"] if cell["value"] is not None else "" for cell in row]
        writer.writerow(values)
```

**複数シートの処理**
```python
# 全てのExcelデータを取得
json_data = sharepoint_excel(file_path=file_path)
data = json.loads(json_data)

# 各シートを処理
summary = {}
for sheet in data["sheets"]:
    sheet_name = sheet["name"]
    row_count = len(sheet["rows"])
    col_count = len(sheet["rows"][0]) if sheet["rows"] else 0

    summary[sheet_name] = {
        "dimensions": sheet["dimensions"],
        "rows": row_count,
        "columns": col_count
    }

print(json.dumps(summary, indent=2, ensure_ascii=False))
```
