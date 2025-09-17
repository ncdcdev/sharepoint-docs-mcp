"""
SharePoint証明書認証モジュール
"""

import base64
import logging
import time
import uuid
from pathlib import Path

import jwt
import requests
from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa

logger = logging.getLogger(__name__)


class SharePointCertificateAuth:
    """SharePoint証明書認証クラス"""

    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        certificate_path: str,
        private_key_path: str,
        site_url: str,
    ):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.certificate_path = Path(certificate_path)
        self.private_key_path = Path(private_key_path)
        self.site_url = site_url
        self._access_token = None
        self._token_expires_at = 0

    def _load_certificate(self) -> x509.Certificate:
        """PEM形式の証明書を読み込む"""
        with open(self.certificate_path, "rb") as cert_file:
            cert_data = cert_file.read()
        return x509.load_pem_x509_certificate(cert_data)

    def _load_private_key(self) -> rsa.RSAPrivateKey:
        """PEM形式の秘密鍵を読み込む"""
        with open(self.private_key_path, "rb") as key_file:
            key_data = key_file.read()
        return serialization.load_pem_private_key(key_data, password=None)

    def _get_certificate_thumbprint(self) -> str:
        """証明書の拇印（thumbprint）を取得"""
        cert = self._load_certificate()
        # SHA1ハッシュを計算
        fingerprint = cert.fingerprint(hashes.SHA1())
        # Base64URLエンコーディング
        return base64.urlsafe_b64encode(fingerprint).decode("utf-8").rstrip("=")

    def _create_client_assertion(self) -> str:
        """クライアントアサーション（JWT）を作成"""
        private_key = self._load_private_key()
        thumbprint = self._get_certificate_thumbprint()

        # JWTヘッダー
        headers = {"alg": "RS256", "typ": "JWT", "x5t": thumbprint}

        # JWTペイロード
        now = int(time.time())
        payload = {
            "aud": f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token",
            "exp": now + 300,  # 5分後に期限切れ
            "iss": self.client_id,
            "jti": str(uuid.uuid4()),
            "nbf": now,
            "sub": self.client_id,
        }

        # JWTを作成
        return jwt.encode(payload, private_key, algorithm="RS256", headers=headers)

    def _request_access_token(self) -> dict[str, str]:
        """アクセストークンを要求"""
        client_assertion = self._create_client_assertion()

        # OAuth2 v2.0トークンエンドポイント
        token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"

        # SharePointサイトのテナント名を取得
        from urllib.parse import urlparse

        parsed_url = urlparse(self.site_url)
        tenant_name = parsed_url.netloc.split('.sharepoint.com')[0]
        scope = f"https://{tenant_name}.sharepoint.com/.default"

        # リクエストパラメータ
        data = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
            "client_assertion": client_assertion,
            "scope": scope,
        }

        headers = {"Content-Type": "application/x-www-form-urlencoded"}

        logger.info("Requesting access token from Microsoft OAuth2 endpoint")
        response = requests.post(token_url, data=data, headers=headers, timeout=30)
        response.raise_for_status()

        token_data = response.json()
        logger.info("Successfully obtained access token")
        return token_data

    def get_access_token(self) -> str:
        """有効なアクセストークンを取得（キャッシュ機能付き）"""
        current_time = time.time()

        # トークンが期限切れまたは存在しない場合は新しく取得
        if not self._access_token or current_time >= self._token_expires_at:
            logger.info("Access token expired or not found, requesting new token")
            token_data = self._request_access_token()
            self._access_token = token_data["access_token"]
            # 期限切れ時刻を設定（実際の期限より5分早く設定してマージンを持たせる）
            self._token_expires_at = current_time + int(token_data["expires_in"]) - 300

        return self._access_token
