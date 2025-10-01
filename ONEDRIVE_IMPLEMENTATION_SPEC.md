# OneDrive対応実装仕様書

## 1. 概要

SharePoint Search MCPサーバーにOneDrive検索機能を追加し、OneDriveとSharePointサイトの混合検索を実現する。

## 2. 新機能

### 2.1 環境変数追加
```bash
# OneDriveユーザーとフォルダーの指定
SHAREPOINT_ONEDRIVE_PATHS=user1@company.com,user2@company.com:/Documents/Projects
```

### 2.2 SHAREPOINT_SITE_NAME拡張
```bash
# @onedriveキーワードでOneDrive検索を有効化
SHAREPOINT_SITE_NAME=@onedrive,team-site,project-alpha
```

## 3. 設定構文仕様

### 3.1 SHAREPOINT_ONEDRIVE_PATHS形式
```
形式: user@domain.com[:/folder/path][,user2@domain.com[:/folder/path]]...

例:
- user@company.com                    # ユーザー全体
- user@company.com:/Documents/Projects # 特定フォルダー
- user1@company.com,user2@company.com:/Documents/重要書類 # 混合
```

### 3.2 パス変換ルール
```
メールアドレス変換:
user@company.com → user_company_com

OneDriveパス構築:
user@company.com:/Documents/Projects → personal/user_company_com/Documents/Projects
```

## 4. 実装要件

### 4.1 config.py拡張
- `onedrive_paths`プロパティ追加
- `parse_onedrive_paths()`メソッド実装
- `get_onedrive_targets()`メソッド実装
- `sites`プロパティ追加（複数サイト対応）
- `include_onedrive`プロパティ追加

### 4.2 sharepoint_search.py拡張  
- 複数サイト対応のKQLクエリ構築
- OneDrive pathフィルター生成
- @onedriveキーワード処理
- `build_multi_target_query()`メソッド実装

### 4.3 後方互換性
- 既存のSHAREPOINT_SITE_NAME単一サイト指定は維持
- 空文字列指定時のテナント全体検索は維持
- 既存の`is_site_specific`プロパティは維持

## 5. エラーハンドリング
- 無効なメールアドレス形式 → 警告ログ + スキップ
- @onedriveが指定されているがSHAREPOINT_ONEDRIVE_PATHSが空 → 警告ログ + @onedrive無視
- 重複ユーザー → より具体的な指定（フォルダー指定）を優先

## 6. KQLクエリ例

### 6.1 OneDriveのみ
```kql
検索語 AND (path:"https://company.sharepoint.com/personal/user1_company_com" OR path:"https://company.sharepoint.com/personal/user2_company_com/Documents/Projects")
```

### 6.2 OneDrive + SharePoint混合
```kql
検索語 AND (path:"https://company.sharepoint.com/personal/user1_company_com" OR site:"https://company.sharepoint.com/sites/team-site")
```

## 7. 設定例

### 7.1 基本設定
```bash
SHAREPOINT_BASE_URL=https://company.sharepoint.com
SHAREPOINT_TENANT_ID=your-tenant-id
SHAREPOINT_CLIENT_ID=your-client-id
SHAREPOINT_CERTIFICATE_PATH=./cert/certificate.pem
SHAREPOINT_PRIVATE_KEY_PATH=./cert/private_key.pem

# OneDrive設定
SHAREPOINT_ONEDRIVE_PATHS=user1@company.com,user2@company.com:/Documents/Projects

# 検索対象
SHAREPOINT_SITE_NAME=@onedrive,team-site
```

### 7.2 用途別パターン
```bash
# パターン1: OneDriveのみ
SHAREPOINT_SITE_NAME=@onedrive

# パターン2: 特定サイトのみ（従来通り）
SHAREPOINT_SITE_NAME=team-site

# パターン3: OneDrive + 複数サイト
SHAREPOINT_SITE_NAME=@onedrive,team-site,project-alpha

# パターン4: テナント全体（従来通り）
SHAREPOINT_SITE_NAME=@all
```

## 8. 実装順序

1. config.py拡張（設定解析機能）
2. sharepoint_search.py拡張（検索ロジック）
3. 動作テスト
4. 品質チェック（lint、typecheck）