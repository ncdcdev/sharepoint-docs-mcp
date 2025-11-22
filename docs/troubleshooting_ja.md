# トラブルシューティングガイド

このガイドでは、SharePoint MCPサーバーのよくある問題について説明します。

## 目次

- [よくある問題](#よくある問題)

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
