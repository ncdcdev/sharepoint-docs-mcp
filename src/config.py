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
        self.base_url = os.getenv(
            "SHAREPOINT_BASE_URL", ""
        )  # https://company.sharepoint.com
        self.site_name = os.getenv("SHAREPOINT_SITE_NAME", "")  # sitename（オプション）

        # OneDrive設定
        self.onedrive_paths = os.getenv("SHAREPOINT_ONEDRIVE_PATHS", "")
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

        # ツール説明文のカスタマイズ
        self.search_tool_description = os.getenv(
            "SHAREPOINT_SEARCH_TOOL_DESCRIPTION",
            "Search for documents in SharePoint. Use response_format='compact' for token-efficient results with only title, path, and extension.",
        )
        self.download_tool_description = os.getenv(
            "SHAREPOINT_DOWNLOAD_TOOL_DESCRIPTION", "Download a file from SharePoint"
        )

    @property
    def site_url(self) -> str:
        """サイトURLを取得（サイト名が指定されている場合のみ）"""
        if self.site_name:
            return f"{self.base_url}/sites/{self.site_name}"
        return self.base_url

    @property
    def is_site_specific(self) -> bool:
        """特定のサイトに限定されているかどうか"""
        return bool(self.site_name) and not self.has_multiple_targets

    @property
    def has_multiple_targets(self) -> bool:
        """複数サイトまたはOneDriveを含む検索かどうか"""
        if not self.site_name:
            return False

        # 特別キーワードを含めた全サイトリストをチェック
        all_sites = [site.strip() for site in self.site_name.split(",") if site.strip()]
        return self.include_onedrive or len(self.sites) > 1 or "@all" in all_sites

    @property
    def include_onedrive(self) -> bool:
        """OneDrive検索が含まれるかどうか"""
        if not self.site_name:
            return False
        sites_with_keywords = [
            site.strip() for site in self.site_name.split(",") if site.strip()
        ]
        return "@onedrive" in sites_with_keywords and bool(self.onedrive_paths)

    @property
    def sites(self) -> list[str]:
        """検索対象サイトのリスト（@onedriveなど特別キーワードを除く）"""
        if not self.site_name:
            return []

        sites = [site.strip() for site in self.site_name.split(",") if site.strip()]
        # 特別キーワードを除外
        return [site for site in sites if not site.startswith("@")]

    def _parse_file_extensions(self, extensions_str: str) -> list[str]:
        """ファイル拡張子文字列をリストに変換"""
        if not extensions_str:
            return []
        return [ext.strip().lower() for ext in extensions_str.split(",") if ext.strip()]

    def parse_onedrive_paths(self) -> list[dict[str, str]]:
        """OneDriveパス設定を解析してユーザーとフォルダー情報を返す"""
        if not self.onedrive_paths:
            return []

        result = []
        entries = [
            entry.strip() for entry in self.onedrive_paths.split(",") if entry.strip()
        ]

        for entry in entries:
            if ":" in entry:
                # user@domain.com:/folder/path形式
                email, folder_path = entry.split(":", 1)
                email = email.strip()
                folder_path = folder_path.strip()
            else:
                # user@domain.com形式（ユーザー全体）
                email = entry.strip()
                folder_path = ""

            # メールアドレス形式の簡易チェック
            if "@" not in email:
                continue

            # OneDriveパスを構築
            onedrive_path = self._email_to_onedrive_path(email, folder_path)

            result.append(
                {
                    "email": email,
                    "folder_path": folder_path,
                    "onedrive_path": onedrive_path,
                }
            )

        return result

    def _email_to_onedrive_path(self, email: str, folder_path: str = "") -> str:
        """メールアドレスをOneDriveパスに変換"""
        # user@company.com → user_company_com
        onedrive_user = email.replace("@", "_").replace(".", "_")

        # personal/user_company_com[/folder/path]
        onedrive_path = f"personal/{onedrive_user}"
        if folder_path:
            # 先頭の/を除去
            folder_path = folder_path.lstrip("/")
            if folder_path:
                onedrive_path = f"{onedrive_path}/{folder_path}"

        return onedrive_path

    def get_onedrive_targets(self) -> list[dict[str, str]]:
        """OneDrive検索対象を取得"""
        if not self.include_onedrive:
            return []

        return self.parse_onedrive_paths()

    def validate(self) -> list[str]:
        """設定の検証を行い、エラーメッセージのリストを返す"""
        errors = []

        if not self.base_url:
            errors.append("SHAREPOINT_BASE_URL is required")

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
