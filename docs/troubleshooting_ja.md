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

#### Excel Servicesが無効

```
Excel Services is not enabled or not available for this SharePoint site.
```

**解決方法**
- SharePoint管理者にExcel Servicesの有効化を依頼
- 対象のSharePointサイトでExcel Servicesが利用可能か確認
- ファイルがExcel Servicesが有効な場所に保存されているか確認

#### Excelファイルが見つからない

```
The specified Excel file was not found: /sites/team/documents/report.xlsx
```

**解決方法**
- ファイルパスが正しいか確認（`sharepoint_docs_search`で最新のパスを取得）
- ファイルが削除または移動されていないか確認
- ファイルへのアクセス権限があるか確認

#### シートが見つからない

```
The specified sheet was not found: Sheet2
```

**解決方法**
- `list_sheets`操作でシート一覧を確認
- シート名のスペルが正しいか確認
- シート名に特殊文字（シングルクォートなど）が含まれる場合は正確に指定

#### 無効なセル範囲

```
The specified cell range is invalid: InvalidRange
```

**解決方法**
- セル範囲の形式が正しいか確認（例: "Sheet1!A1:C10"）
- シート名がセル範囲に含まれているか確認
- セル範囲がExcelファイルの実際の範囲内か確認

## デバッグ方法

### MCP Inspectorを使用

```bash
npx @modelcontextprotocol/inspector uv run sharepoint-docs-mcp --transport stdio
```

### ログレベルの調整

サーバー起動時に詳細なログが出力されます。エラーの詳細は標準エラー出力に表示されます。
