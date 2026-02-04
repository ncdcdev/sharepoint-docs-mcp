# トラブルシューティングガイド

このガイドでは、SharePoint MCPサーバーのよくある問題とデバッグ方法について説明します。

## 目次

- [よくある問題](#よくある問題)
- [デバッグ方法](#デバッグ方法)

## よくある問題

### 1. 認証エラー

```
SharePoint configuration is invalid: SHAREPOINT_TENANT_ID is required
```

**解決方法**
- `.env`ファイルが正しく設定されているか確認
- 環境変数が正しく読み込まれているか確認

### 2. 証明書エラー

```
Certificate file not found: path/to/certificate.pem
```

**解決方法**
- 証明書ファイルのパスが正しいか確認
- 証明書が正しく作成されているか確認
- ファイルの読み取り権限があるか確認

### 3. API権限エラー

```
Access token request failed
```

**解決方法**
- Azure ADアプリの権限設定を確認
- 管理者の同意が行われているか確認
- クライアントIDとテナントIDが正しいか確認

### 4. 設定確認コマンド

```bash
# 設定ステータスを確認（MCP Inspector使用）
# get_sharepoint_config_status ツールを実行
```

### 5. Excel操作エラー

#### 無効なExcelファイル形式

```
The file is not a valid Excel file or is corrupted. Please verify the file is a valid .xlsx file. Try opening it in Excel locally to check for corruption, or re-upload the file to SharePoint.
```

**解決方法**
- ファイルが有効な.xlsxファイルか確認（.xlsや他の形式ではない）
- ファイルが破損していないか、ローカルのExcelで開いて確認
- ファイルをSharePointに再アップロードしてみる

#### Excelファイルが見つからない

```
The specified Excel file was not found: /sites/team/documents/report.xlsx Please verify the file path is correct and the file exists. You can search for the file using sharepoint_docs_search with file_extensions=['xlsx'] to get the correct path.
```

**解決方法**
- ファイルパスが正しいか確認（`sharepoint_docs_search`で最新のパスを取得）
- ファイルが削除または移動されていないか確認
- ファイルへのアクセス権限があるか確認

#### シートが見つからない

```
The specified sheet was not found: Sheet2 Run sharepoint_excel without specifying 'sheet' to list available sheets (check sheets[].name in the response), then use a valid sheet name.
```

**解決方法**
- まず`sheet`パラメータなしでファイルを読み取り、利用可能なシート一覧を確認
- シート名のスペルが正しいか確認（大文字小文字を区別）
- シート名の前後にスペースがないか確認

#### 無効なセル範囲

```
The specified cell range is invalid: ZZ999999 Please use a valid range format like 'A1:C10' or 'A1'. Ensure the range is within the actual bounds of the Excel file.
```

**解決方法**
- セル範囲の形式が正しいか確認（例: "A1:C10" または "A1"）
- セル範囲がExcelファイルの実際の範囲内か確認
- 列文字と行番号が有効か確認

## デバッグ方法

### MCP Inspectorを使用

```bash
npx @modelcontextprotocol/inspector uv run sharepoint-docs-mcp --transport stdio
```

### ログレベルの調整

サーバー起動時に詳細なログが出力されます。エラーの詳細は標準エラー出力に表示されます。
