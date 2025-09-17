"""
設定管理モジュール
"""

import os
from pathlib import Path

from dotenv import load_dotenv

# .envファイルを読み込み
load_dotenv()


class SharePointConfig:
    """SharePoint設定クラス"""

    def __init__(self):
        # SharePoint設定
        self.site_url = os.getenv("SHAREPOINT_SITE_URL", "")
        self.tenant_id = os.getenv("SHAREPOINT_TENANT_ID", "")
        self.client_id = os.getenv("SHAREPOINT_CLIENT_ID", "")

        # 証明書認証設定（ファイルパスまたはテキスト）
        self.certificate_path = os.getenv("SHAREPOINT_CERTIFICATE_PATH", "")
        self.certificate_text = os.getenv("SHAREPOINT_CERTIFICATE_TEXT", "")
        self.private_key_path = os.getenv("SHAREPOINT_PRIVATE_KEY_PATH", "")
        self.private_key_text = os.getenv("SHAREPOINT_PRIVATE_KEY_TEXT", "")

        # 検索設定
        self.default_max_results = int(
            os.getenv("SHAREPOINT_DEFAULT_MAX_RESULTS", "20")
        )
        self.allowed_file_extensions = self._parse_file_extensions(
            os.getenv("SHAREPOINT_ALLOWED_FILE_EXTENSIONS", "pdf,docx,xlsx,pptx,txt")
        )

    def _parse_file_extensions(self, extensions_str: str) -> list[str]:
        """ファイル拡張子文字列をリストに変換"""
        if not extensions_str:
            return []
        return [ext.strip().lower() for ext in extensions_str.split(",") if ext.strip()]

    def validate(self) -> list[str]:
        """設定の検証を行い、エラーメッセージのリストを返す"""
        errors = []

        if not self.site_url:
            errors.append("SHAREPOINT_SITE_URL is required")

        if not self.tenant_id:
            errors.append("SHAREPOINT_TENANT_ID is required")

        if not self.client_id:
            errors.append("SHAREPOINT_CLIENT_ID is required")

        # 証明書：ファイルパスまたはテキストのいずれかが必要
        if not self.certificate_path and not self.certificate_text:
            errors.append(
                "Either SHAREPOINT_CERTIFICATE_PATH or SHAREPOINT_CERTIFICATE_TEXT is required"
            )
        elif self.certificate_path and not Path(self.certificate_path).exists():
            errors.append(f"Certificate file not found: {self.certificate_path}")

        # 秘密鍵：ファイルパスまたはテキストのいずれかが必要
        if not self.private_key_path and not self.private_key_text:
            errors.append(
                "Either SHAREPOINT_PRIVATE_KEY_PATH or SHAREPOINT_PRIVATE_KEY_TEXT is required"
            )
        elif self.private_key_path and not Path(self.private_key_path).exists():
            errors.append(f"Private key file not found: {self.private_key_path}")

        return errors

    @property
    def is_valid(self) -> bool:
        """設定が有効かどうかを返す"""
        return len(self.validate()) == 0


# グローバル設定インスタンス
config = SharePointConfig()
