# SharePoint MCP OAuth認証対応 設計ドキュメント

## 概要

SharePoint Document Search MCPサーバーにOAuth 2.0認証（ユーザー権限）のサポートを追加する設計です。
現在の証明書ベース認証（アプリケーション権限）と並行して、環境変数で認証方式を切り替え可能にします。

## 現状分析

### 現在の認証方式

証明書ベース認証（OAuth 2.0 Client Credentials Flow）
- アプリケーション権限でSharePointにアクセス
- Azure AD証明書による認証
- `SharePointCertificateAuth`クラスでトークン管理
- トークンのメモリキャッシュ（5分のマージン付き）

### 現在のアーキテクチャ

```
main.py (CLI)
  ├─ server.py (MCP tools)
  │    ├─ config.py (環境変数)
  │    ├─ sharepoint_auth.py (証明書認証)
  │    └─ sharepoint_search.py (SharePoint API)
  └─ FastMCP (stdio/http transport)
```

### トランスポート対応状況
- stdio: 対応済み
- streamable-http: 対応済み

## OAuth認証の要件

### 認証フローの選択

Authorization Code Flow with PKCE（RFC 7636）を採用

選択理由
- ユーザー権限でのアクセスが必要
- リフレッシュトークンによる長期利用
- クライアントシークレット不要（PKCEによる保護）
- EntraIDが公式サポート
- DCR（Dynamic Client Registration）は不要

### 対象トランスポート

streamable-http専用
- stdio: OAuth認証は対応しない
  - 理由: ブラウザ認証フローとの相性が悪い
  - stdioはローカル実行が前提のため証明書認証で十分
- streamable-http: OAuth認証を完全サポート
  - HTTPベースのストリーミング通信に対応

## アーキテクチャ設計

### 認証方式の切り替え

環境変数による認証モード選択

```bash
# 証明書ベース認証（既存）
SHAREPOINT_AUTH_MODE=certificate

# OAuth認証（新規）
SHAREPOINT_AUTH_MODE=oauth
```

### OAuth認証クラス設計

新規クラス: `SharePointOAuthAuth`

```python
class SharePointOAuthAuth:
    """SharePoint OAuth認証クラス（Authorization Code Flow with PKCE）"""

    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        site_url: str,
        redirect_uri: str = "http://localhost:8000/oauth/callback",
        token_cache_path: str = ".sharepoint_tokens.json",
    ):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.site_url = site_url
        self.redirect_uri = redirect_uri
        self.token_cache_path = Path(token_cache_path)

        self._access_token = None
        self._refresh_token = None
        self._token_expires_at = 0

    def start_auth_flow(self) -> tuple[str, str]:
        """認証フローを開始（Authorization URL生成 + PKCE）"""

    def handle_auth_callback(self, code: str, state: str, code_verifier: str) -> dict:
        """認証コールバックを処理してトークンを取得"""

    def get_access_token(self) -> str:
        """有効なアクセストークンを取得（自動リフレッシュ）"""

    def refresh_access_token(self) -> dict:
        """リフレッシュトークンを使用して新しいトークンを取得"""

    def save_tokens(self, token_data: dict) -> None:
        """トークンをファイルに保存"""

    def load_tokens(self) -> bool:
        """トークンをファイルから読み込み"""
```

### トークン管理

トークンの永続化
- ファイルベース（JSON形式）
- デフォルトパス: `.sharepoint_tokens.json`
- 環境変数で変更可能: `SHAREPOINT_TOKEN_CACHE_PATH`

トークンデータ構造
```json
{
  "access_token": "eyJ0eXAiOiJKV1QiLCJub...",
  "refresh_token": "0.AXoA3H...",
  "expires_at": 1704123456.789,
  "scope": "https://tenant.sharepoint.com/.default",
  "token_type": "Bearer"
}
```

### HTTPエンドポイント設計

FastMCPのHTTPサーバーに認証用エンドポイントを追加

```python
# 認証開始エンドポイント
GET /oauth/login
  → Authorization URLにリダイレクト

# コールバックエンドポイント
GET /oauth/callback?code=xxx&state=xxx
  → トークン取得 → 成功ページ表示

# 認証状態確認エンドポイント
GET /oauth/status
  → {"authenticated": true/false, "expires_at": timestamp}
```

### 環境変数設計

新規追加する環境変数

```bash
# 認証モード（必須）
SHAREPOINT_AUTH_MODE=oauth  # "certificate" または "oauth"

# OAuth設定（OAuthモード時に必須）
SHAREPOINT_OAUTH_CLIENT_ID=your-client-id
SHAREPOINT_OAUTH_REDIRECT_URI=http://localhost:8000/oauth/callback

# トークンキャッシュ（オプション）
SHAREPOINT_TOKEN_CACHE_PATH=.sharepoint_tokens.json

# 既存の環境変数（両モードで共通）
SHAREPOINT_BASE_URL=https://company.sharepoint.com
SHAREPOINT_TENANT_ID=your-tenant-id
SHAREPOINT_SITE_NAME=sitename
```

証明書モード（既存）の環境変数
```bash
SHAREPOINT_AUTH_MODE=certificate
SHAREPOINT_CLIENT_ID=your-app-id
SHAREPOINT_CERTIFICATE_PATH=/path/to/cert.pem
SHAREPOINT_PRIVATE_KEY_PATH=/path/to/key.pem
```

### 設定クラスの拡張

`SharePointConfig`クラスに認証モード関連を追加

```python
class SharePointConfig:
    def __init__(self):
        # 認証モード
        self.auth_mode = os.getenv("SHAREPOINT_AUTH_MODE", "certificate")

        # OAuth設定
        self.oauth_client_id = os.getenv("SHAREPOINT_OAUTH_CLIENT_ID", "")
        self.oauth_redirect_uri = os.getenv(
            "SHAREPOINT_OAUTH_REDIRECT_URI",
            "http://localhost:8000/oauth/callback"
        )
        self.token_cache_path = os.getenv(
            "SHAREPOINT_TOKEN_CACHE_PATH",
            ".sharepoint_tokens.json"
        )

        # 既存の設定...

    @property
    def is_oauth_mode(self) -> bool:
        """OAuth認証モードかどうか"""
        return self.auth_mode.lower() == "oauth"

    @property
    def is_certificate_mode(self) -> bool:
        """証明書認証モードかどうか"""
        return self.auth_mode.lower() == "certificate"

    def validate(self) -> list[str]:
        """設定の検証（認証モード別）"""
        errors = []

        # 共通検証
        if not self.base_url:
            errors.append("SHAREPOINT_BASE_URL is required")
        if not self.tenant_id:
            errors.append("SHAREPOINT_TENANT_ID is required")

        # 認証モード別検証
        if self.is_oauth_mode:
            if not self.oauth_client_id:
                errors.append("SHAREPOINT_OAUTH_CLIENT_ID is required for OAuth mode")
        elif self.is_certificate_mode:
            # 既存の証明書検証...
            if not self.client_id:
                errors.append("SHAREPOINT_CLIENT_ID is required for certificate mode")
        else:
            errors.append(f"Invalid SHAREPOINT_AUTH_MODE: {self.auth_mode}")

        return errors
```

### server.pyの改修

認証クライアント取得の抽象化

```python
def _get_auth_client() -> SharePointCertificateAuth | SharePointOAuthAuth:
    """認証クライアントを取得（モードに応じて切り替え）"""
    if config.is_oauth_mode:
        return SharePointOAuthAuth(
            tenant_id=config.tenant_id,
            client_id=config.oauth_client_id,
            site_url=config.site_url,
            redirect_uri=config.oauth_redirect_uri,
            token_cache_path=config.token_cache_path,
        )
    else:
        return SharePointCertificateAuth(
            tenant_id=config.tenant_id,
            client_id=config.client_id,
            site_url=config.site_url,
            certificate_path=config.certificate_path,
            certificate_text=config.certificate_text,
            private_key_path=config.private_key_path,
            private_key_text=config.private_key_text,
        )

def _get_sharepoint_client() -> SharePointSearchClient:
    """SharePointクライアントを取得または初期化"""
    global _sharepoint_client

    if _sharepoint_client is None:
        validation_errors = config.validate()
        if validation_errors:
            raise ValueError("Configuration errors: " + "; ".join(validation_errors))

        auth = _get_auth_client()
        _sharepoint_client = SharePointSearchClient(
            site_url=config.site_url,
            auth=auth,
        )

    return _sharepoint_client
```

### FastMCPへのOAuthエンドポイント追加

FastMCPは基本的にMCPプロトコルのハンドラーのため、OAuth用のHTTPエンドポイントは別途実装

オプション1: FastMCPのHTTPサーバーを拡張
- FastMCPの内部実装（Starlette/ASGI）に直接エンドポイントを追加

オプション2: 別ポートでOAuth専用サーバーを起動
- FastMCPとは独立したHTTPサーバー（Flask/FastAPI）
- 認証完了後にトークンファイルに保存

**推奨**: オプション1（FastMCP拡張）
- 理由: 単一プロセス、ポート管理が簡単、設定が統一

実装方法（FastMCP拡張）
```python
from fastmcp import FastMCP
from starlette.responses import RedirectResponse, HTMLResponse
from starlette.routing import Route

# MCPインスタンス
mcp = FastMCP(name="SharePointDocsMCP")

def setup_oauth_endpoints():
    """OAuth認証用のHTTPエンドポイントを追加"""
    if not config.is_oauth_mode:
        return

    async def oauth_login(request):
        """OAuth認証開始"""
        auth = _get_auth_client()
        auth_url, state = auth.start_auth_flow()
        # state, code_verifierをセッション/メモリに保存
        return RedirectResponse(url=auth_url)

    async def oauth_callback(request):
        """OAuth認証コールバック"""
        code = request.query_params.get("code")
        state = request.query_params.get("state")
        # セッションからcode_verifierを取得
        auth = _get_auth_client()
        token_data = auth.handle_auth_callback(code, state, code_verifier)
        auth.save_tokens(token_data)
        return HTMLResponse("<h1>認証成功</h1><p>ウィンドウを閉じてください</p>")

    async def oauth_status(request):
        """認証状態確認"""
        auth = _get_auth_client()
        authenticated = auth.load_tokens()
        return {
            "authenticated": authenticated,
            "expires_at": auth._token_expires_at if authenticated else None
        }

    # FastMCPのASGIアプリにルートを追加
    # ※ FastMCPの内部実装に依存するため、実装時に調査が必要
```

## Azure AD (Entra ID) アプリ登録設定

OAuth認証を使用するために必要なAzure ADアプリの設定

### アプリケーション登録
1. Azure Portal → Azure Active Directory → アプリの登録
2. 新しいアプリケーションを登録
   - 名前: SharePoint MCP OAuth Client
   - サポートされているアカウントの種類: 単一テナント
   - リダイレクトURI: Web - `http://localhost:8000/oauth/callback`

### API のアクセス許可
必要なMicrosoft Graph/SharePoint権限（委任されたアクセス許可）:
- `Sites.Read.All` - SharePointサイトの読み取り
- `Files.Read.All` - ファイルの読み取り
- `User.Read` - ユーザープロファイルの読み取り（基本）

### 認証設定
- パブリッククライアントフローを許可: いいえ
- リダイレクトURI: `http://localhost:8000/oauth/callback`
- ログアウトURL: （オプション）
- 暗黙的な許可とハイブリッドフロー: すべてオフ

## PKCEフロー詳細

### 1. 認証開始

```python
import secrets
import hashlib
import base64

# Code Verifier生成（43-128文字のランダム文字列）
code_verifier = base64.urlsafe_b64encode(secrets.token_bytes(32)).decode('utf-8').rstrip('=')

# Code Challenge生成（SHA256ハッシュ）
code_challenge = base64.urlsafe_b64encode(
    hashlib.sha256(code_verifier.encode('utf-8')).digest()
).decode('utf-8').rstrip('=')

# State生成（CSRF対策）
state = secrets.token_urlsafe(32)

# Authorization URL構築
auth_url = (
    f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/authorize"
    f"?client_id={client_id}"
    f"&response_type=code"
    f"&redirect_uri={redirect_uri}"
    f"&scope=https://{tenant}.sharepoint.com/.default offline_access"
    f"&state={state}"
    f"&code_challenge={code_challenge}"
    f"&code_challenge_method=S256"
)
```

### 2. トークン取得

```python
# Authorization Codeを受け取り後
token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

data = {
    "client_id": client_id,
    "grant_type": "authorization_code",
    "code": authorization_code,
    "redirect_uri": redirect_uri,
    "code_verifier": code_verifier,  # PKCEの検証
}

response = requests.post(token_url, data=data)
token_data = response.json()
```

### 3. トークンリフレッシュ

```python
data = {
    "client_id": client_id,
    "grant_type": "refresh_token",
    "refresh_token": refresh_token,
    "scope": f"https://{tenant}.sharepoint.com/.default offline_access",
}

response = requests.post(token_url, data=data)
```

## セキュリティ考慮事項

### トークンの保護
- ファイルパーミッション: 0600（所有者のみ読み書き可能）
- gitignore: トークンキャッシュファイルを除外
- 環境変数での上書き可能（パス指定）

### PKCE
- Code Verifierの安全な生成（高エントロピー）
- Code Challengeの検証
- StateパラメータによるCSRF対策

### セッション管理
- メモリ内でstate/code_verifierを一時保存
- タイムアウト: 5分（認証フロー完了までの猶予）
- 使用後は即座に削除

### トークンのリフレッシュ
- アクセストークン有効期限: 通常1時間
- リフレッシュトークン有効期限: 90日（Azure AD デフォルト）
- 5分のマージン付き自動リフレッシュ

## 実装ステップ

### Phase 1: OAuth認証基盤
1. ブランチ作成: `feature/oauth-authentication`
2. `sharepoint_oauth_auth.py`を作成
   - `SharePointOAuthAuth`クラス実装
   - PKCE対応
   - トークン永続化
3. `config.py`の拡張
   - 認証モード設定
   - OAuth関連環境変数
4. 単体テスト作成

### Phase 2: HTTPエンドポイント追加
1. FastMCPのHTTP拡張を調査
2. OAuthエンドポイント実装
   - `/oauth/login`
   - `/oauth/callback`
   - `/oauth/status`
3. セッション管理実装（state/code_verifier）

### Phase 3: 既存コードの統合
1. `server.py`の改修
   - 認証クライアント抽象化
   - モード別クライアント生成
2. `main.py`の改修
   - OAuthモード時のstdio無効化
   - HTTPサーバー起動時の認証状態チェック

### Phase 4: ドキュメント整備
1. README更新
   - OAuth認証の設定方法
   - Azure ADアプリ登録手順
2. 環境変数サンプル（.env.example）
3. トラブルシューティングガイド

### Phase 5: テストと検証
1. 統合テスト
2. 実環境でのOAuthフローテスト
3. トークンリフレッシュの動作確認

## 既存コードとの互換性

### 下位互換性の維持
- デフォルトは証明書モード（`SHAREPOINT_AUTH_MODE=certificate`）
- 既存の環境変数はすべて動作継続
- 既存のstdio利用者には影響なし

### 移行パス
証明書モード → OAuthモード
1. Azure ADアプリの設定変更（リダイレクトURI追加）
2. 環境変数の更新
3. 初回起動時に`/oauth/login`でブラウザ認証
4. 以降は自動的にトークンリフレッシュ

## 補足: FastMCPの制約事項

FastMCPはMCPプロトコルに特化しているため、カスタムHTTPエンドポイントの追加には制約がある可能性があります。

代替案として、OAuth認証専用の軽量HTTPサーバーを別プロセスで起動することも検討できます。

```python
# 別ポート（例: 8001）でOAuth専用サーバー
# トークン取得後、ファイルに保存
# FastMCPサーバー（8000）は保存されたトークンを読み込み
```

ただし、この場合もトークンの永続化により、初回認証後は自動リフレッシュで運用可能です。

## まとめ

認証方式の対応
- 証明書認証（既存）: アプリケーション権限、stdio/streamable-http対応
- OAuth認証（新規）: ユーザー権限、streamable-http専用

主な変更点
- 新規クラス: `SharePointOAuthAuth`
- 設定拡張: 認証モード切り替え
- HTTPエンドポイント: OAuth認証フロー
- トークン管理: ファイルベース永続化

利点
- ユーザー権限でのアクセス
- 長期利用（リフレッシュトークン）
- セキュアな認証（PKCE）
- 既存機能との共存
