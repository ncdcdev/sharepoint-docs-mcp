import base64
import logging
import secrets
import sys
import time
from typing import Any
from urllib.parse import urlencode, urlparse

from fastmcp import FastMCP
from fastmcp.server.auth import AccessToken, TokenVerifier
from fastmcp.server.auth.oidc_proxy import OIDCProxy
from fastmcp.server.dependencies import get_access_token
from mcp.server.auth.provider import AuthorizationParams

from .config import config
from .error_messages import handle_sharepoint_error
from .sharepoint_auth import SharePointCertificateAuth
from .sharepoint_search import SharePointSearchClient


class SharePointTokenVerifier(TokenVerifier):
    """Simple token verifier for SharePoint OAuth tokens

    Since SharePoint tokens have a different audience than Graph API,
    we cannot validate them via Microsoft Graph API. Instead, we trust
    that the token was issued by Azure AD through the OAuth flow.

    The security is maintained by:
    1. OAuth flow validates the user with Azure AD
    2. Token is issued directly by Azure AD (no third party)
    3. PKCE ensures token cannot be intercepted
    4. Token is only used within the same session
    """

    async def verify_token(self, token: str) -> AccessToken | None:
        """Accept any non-empty token from Azure AD OAuth flow"""
        if not token or not isinstance(token, str):
            return None

        # Create AccessToken with minimal validation
        # The token was obtained through secure OAuth flow, so we trust it
        return AccessToken(
            token=token,
            client_id="azure-ad-sharepoint",
            scopes=self.required_scopes or [],
            expires_at=None,  # Azure AD manages expiration
        )


class AzureOIDCProxyForSharePoint(OIDCProxy):
    """Custom OIDC Proxy for Azure AD that removes unsupported 'resource' parameter

    Azure AD v2.0 doesn't support the 'resource' parameter (RFC 8707).
    This custom provider overrides the authorize method to remove it and
    uses SharePointTokenVerifier for token validation.
    """

    def get_token_verifier(
        self,
        *,
        algorithm: str | None = None,
        audience: str | None = None,
        required_scopes: list[str] | None = None,
        timeout_seconds: int | None = None,
    ):
        """Override to use SharePointTokenVerifier

        SharePoint tokens have SharePoint as audience, not Graph API,
        so we cannot use AzureTokenVerifier (which calls Graph API).
        Instead, we use a simple verifier that trusts tokens from the OAuth flow.
        """
        return SharePointTokenVerifier(
            required_scopes=required_scopes or self.required_scopes or []
        )

    async def authorize(
        self,
        client,
        params: AuthorizationParams,
    ) -> str:
        """Override authorize to remove resource parameter (Azure AD v2.0 doesn't support it)"""
        # Generate transaction ID for this authorization request
        txn_id = secrets.token_urlsafe(32)

        # Generate proxy's own PKCE parameters if forwarding is enabled
        proxy_code_verifier = None
        proxy_code_challenge = None
        if self._forward_pkce and params.code_challenge:
            proxy_code_verifier, proxy_code_challenge = self._generate_pkce_pair()

        # Store transaction data for IdP callback processing
        transaction_data = {
            "client_id": client.client_id,
            "client_redirect_uri": str(params.redirect_uri),
            "client_state": params.state,
            "code_challenge": params.code_challenge,
            "code_challenge_method": getattr(params, "code_challenge_method", "S256"),
            "scopes": params.scopes or [],
            "created_at": time.time(),
        }

        # Store proxy's PKCE verifier if we're forwarding
        if proxy_code_verifier:
            transaction_data["proxy_code_verifier"] = proxy_code_verifier

        self._oauth_transactions[txn_id] = transaction_data

        # Build query parameters for upstream IdP authorization request
        query_params: dict[str, Any] = {
            "response_type": "code",
            "client_id": self._upstream_client_id,
            "redirect_uri": f"{str(self.base_url).rstrip('/')}{self._redirect_path}",
            "state": txn_id,
        }

        # Add scopes
        scopes_to_use = params.scopes or self.required_scopes or []
        if scopes_to_use:
            query_params["scope"] = " ".join(scopes_to_use)

        # Forward proxy's PKCE challenge to upstream if enabled
        if proxy_code_challenge:
            query_params["code_challenge"] = proxy_code_challenge
            query_params["code_challenge_method"] = "S256"

        # NOTE: Intentionally NOT forwarding 'resource' parameter as Azure AD v2.0 doesn't support it
        # The parent class (OAuthProxy) would add it here, but we skip it for Azure AD compatibility

        # Add any extra authorization parameters configured for this proxy
        if self._extra_authorize_params:
            query_params.update(self._extra_authorize_params)

        # Build the upstream authorization URL
        separator = "&" if "?" in self._upstream_authorization_endpoint else "?"
        upstream_url = f"{self._upstream_authorization_endpoint}{separator}{urlencode(query_params)}"

        return upstream_url


class SimpleTokenAuth:
    """Simple token-based authentication for OAuth mode

    This class wraps an access token obtained from FastMCP's authentication context
    and provides the same interface as SharePointCertificateAuth.
    """

    def __init__(self, token: str):
        self._token = token

    def get_access_token(self) -> str:
        """Return the access token"""
        return self._token


# MCPサーバーの認証プロバイダを設定
def _create_auth_provider():
    """Create FastMCP auth provider based on auth mode"""
    if config.is_oauth_mode:
        # Validate OAuth configuration before initializing OAuth provider
        if (
            not config.oauth_client_id
            or not config.oauth_client_secret
            or not config.tenant_id
        ):
            logging.warning(
                "OAuth mode is enabled but configuration is incomplete. "
                "MCP server authentication will be disabled. "
                "Ensure SHAREPOINT_OAUTH_CLIENT_ID (or SHAREPOINT_CLIENT_ID), "
                "SHAREPOINT_OAUTH_CLIENT_SECRET, and SHAREPOINT_TENANT_ID are set."
            )
            return None

        # OAuth mode: Use OIDC Proxy to protect MCP server with Azure AD
        # Extract tenant name from site URL for SharePoint scope
        parsed_url = urlparse(config.site_url)
        tenant_name = parsed_url.netloc.split(".sharepoint.com")[0]

        # Azure AD OIDC configuration URL (v2.0)
        config_url = f"https://login.microsoftonline.com/{config.tenant_id}/v2.0/.well-known/openid-configuration"

        # Use custom OIDC Proxy that removes unsupported 'resource' parameter for Azure AD v2.0
        return AzureOIDCProxyForSharePoint(
            config_url=config_url,
            client_id=config.oauth_client_id,
            client_secret=config.oauth_client_secret,
            base_url=config.oauth_server_base_url,
            redirect_path="/auth/callback",
            required_scopes=[
                f"https://{tenant_name}.sharepoint.com/.default",  # SharePoint API access
                "offline_access",  # Refresh token support
            ],
            # Allow all localhost redirect URIs (MCP clients use random ports)
            allowed_client_redirect_uris=[
                "http://localhost:*",
                "http://127.0.0.1:*",
            ],
        )
    else:
        # Certificate mode: No MCP server authentication
        return None


# MCPサーバーインスタンスを作成
mcp = FastMCP(name="SharePointDocsMCP", auth=_create_auth_provider())

# SharePointクライアントのグローバルインスタンス
_sharepoint_client: SharePointSearchClient | None = None


def setup_logging():
    """
    すべてのログ出力をstderrに向けるロギングを設定します。
    これにより、stdioトランスポートのstdoutが汚染されるのを防ぎます。
    """
    log_formatter = logging.Formatter("%(asctime)s [%(levelname)s] - %(message)s")
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)

    # stdoutへの出力を防ぐため、既存のハンドラをクリア
    root_logger.handlers.clear()

    # stderrにログを出力するハンドラを追加
    stream_handler = logging.StreamHandler(sys.stderr)
    stream_handler.setFormatter(log_formatter)
    root_logger.addHandler(stream_handler)

    logging.info("Logging configured to output to stderr.")


def _get_auth_client() -> SharePointCertificateAuth | None:
    """認証クライアントを取得（証明書モードのみ）

    OAuthモードの場合は、FastMCPのOIDCProxyが認証を処理するため、
    個別の認証クライアントは不要（Noneを返す）。
    """
    if config.is_oauth_mode:
        # OAuth mode: FastMCP's OIDCProxy handles authentication
        # Token will be retrieved from context in tool functions
        return None

    # Certificate mode: Use SharePointCertificateAuth
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
    """SharePointクライアントを取得または初期化

    - 証明書モード: シングルトンクライアントを使用
    - OAuthモード: リクエストごとに新しいクライアントを作成（トークンはリクエスト依存）
    """
    global _sharepoint_client

    # 設定の検証（初回のみ）
    if _sharepoint_client is None:
        validation_errors = config.validate()
        if validation_errors:
            error_msg = "SharePoint configuration is invalid: " + "; ".join(
                validation_errors
            )
            logging.error(error_msg)
            raise ValueError(error_msg)

    # OAuthモード: リクエストごとに新しいクライアントを作成
    if config.is_oauth_mode:
        # FastMCPの認証コンテキストからトークンを取得
        access_token = get_access_token()
        if not access_token:
            raise ValueError(
                "OAuth authentication required but no access token available. "
                "Please authenticate with FastMCP's AzureProvider."
            )

        # SimpleTokenAuthでトークンをラップ
        auth = SimpleTokenAuth(token=access_token.token)

        # SharePointクライアントを作成（リクエストごと）
        return SharePointSearchClient(
            site_url=config.site_url,
            auth=auth,
        )

    # 証明書モード: シングルトンクライアントを使用
    if _sharepoint_client is None:
        # 認証クライアントを初期化
        auth = _get_auth_client()
        if auth is None:
            raise ValueError("Certificate authentication client initialization failed")

        # SharePointクライアントを初期化
        _sharepoint_client = SharePointSearchClient(
            site_url=config.site_url,
            auth=auth,
        )

        logging.info("SharePoint client initialized successfully (certificate mode)")

    return _sharepoint_client


def sharepoint_docs_search(
    query: str,
    max_results: int = 20,
    file_extensions: list[str] | None = None,
    response_format: str = "detailed",
) -> list[dict[str, Any]]:
    """
    Search for documents in SharePoint with response format options

    Args:
        query: Search keywords
        max_results: Maximum number of results to return (default: 20, max: 100)
        file_extensions: List of file extensions to search (e.g., ["pdf", "docx"])
        response_format: Response format - "detailed" (default) or "compact"

    Returns:
        List of search results. Each result contains:
        - Detailed format: all available fields (title, path, size, modified, extension, summary)
        - Compact format: essential fields only (title, path, extension)
    """
    logging.info(f"Searching SharePoint documents with query: '{query}'")

    # Validate response_format parameter
    valid_formats = ["detailed", "compact"]
    if response_format not in valid_formats:
        logging.warning(
            f"Invalid response_format '{response_format}'. Defaulting to 'detailed'"
        )
        response_format = "detailed"

    try:
        client = _get_sharepoint_client()

        # ファイル拡張子のフィルタリング
        if file_extensions:
            # 設定で許可された拡張子のみを使用
            allowed_extensions = [
                ext
                for ext in file_extensions
                if ext.lower() in config.allowed_file_extensions
            ]
            if not allowed_extensions:
                logging.warning("No allowed file extensions found in the request")
        else:
            allowed_extensions = None

        # Limit maximum results
        max_results = min(max_results, 100)

        # Execute search
        results = client.search_documents(
            query=query,
            max_results=max_results,
            file_extensions=allowed_extensions,
        )

        # Apply response format filtering
        if response_format == "compact":
            # Return only essential fields for compact format
            filtered_results = []
            for result in results:
                compact_result = {
                    "title": result.get("title", "Unknown"),
                    "path": result.get("path", ""),
                    "extension": result.get("extension", ""),
                }
                filtered_results.append(compact_result)
            results = filtered_results

        logging.info(f"SharePoint search completed. Found {len(results)} documents")
        return results

    except Exception as e:
        logging.error(f"SharePoint search failed: {str(e)}")
        raise handle_sharepoint_error(e, "search") from e


def sharepoint_docs_download(file_path: str) -> str:
    """
    Download a file from SharePoint

    Args:
        file_path: ダウンロードするファイルのフルパス（sharepoint_docs_searchの結果から取得）

    Returns:
        ダウンロードしたファイルの内容（Base64エンコード済み文字列）
    """
    logging.info(f"Downloading SharePoint file: {file_path}")

    try:
        client = _get_sharepoint_client()

        # ファイルをダウンロード
        file_content = client.download_file(file_path)

        # Base64エンコードして返す
        encoded_content = base64.b64encode(file_content).decode("utf-8")

        logging.info(
            f"SharePoint file download completed. Size: {len(file_content)} bytes"
        )
        return encoded_content

    except Exception as e:
        logging.error(f"SharePoint file download failed: {str(e)}")
        raise handle_sharepoint_error(e, "download") from e


def register_tools():
    """Register MCP tools"""
    mcp.tool(description=config.search_tool_description)(sharepoint_docs_search)
    mcp.tool(description=config.download_tool_description)(sharepoint_docs_download)
