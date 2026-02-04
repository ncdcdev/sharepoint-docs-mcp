import pytest
from unittest.mock import patch, mock_open
import os

from src.config import SharePointConfig


class TestSharePointConfig:
    """SharePointConfig のテスト"""

    def test_valid_config_from_env(self, mock_env_vars):
        """環境変数から有効な設定を読み込むテスト"""
        config = SharePointConfig()
        validation_errors = config.validate()

        assert validation_errors == []
        assert config.site_url == "https://test.sharepoint.com/sites/test"
        assert config.tenant_id == "test-tenant-id"
        assert config.client_id == "test-client-id"

    def test_missing_required_env_vars(self):
        """必須の環境変数が不足している場合のテスト"""
        with patch.dict(os.environ, {}, clear=True):
            config = SharePointConfig()
            validation_errors = config.validate()

            assert len(validation_errors) > 0
            assert any("SHAREPOINT_TENANT_ID is required" in error for error in validation_errors)
            assert any("SHAREPOINT_CLIENT_ID is required" in error for error in validation_errors)

    def test_certificate_text_priority_over_file(self):
        """証明書テキストがファイルパスより優先されることのテスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_SITE_NAME": "test",
            "SHAREPOINT_TENANT_ID": "test-tenant-id",
            "SHAREPOINT_CLIENT_ID": "test-client-id",
            "SHAREPOINT_CERTIFICATE_PATH": "/path/to/cert.pem",
            "SHAREPOINT_CERTIFICATE_TEXT": "certificate-text",
            "SHAREPOINT_PRIVATE_KEY_PATH": "/path/to/key.pem",
            "SHAREPOINT_PRIVATE_KEY_TEXT": "private-key-text",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()

            assert config.certificate_text == "certificate-text"
            assert config.private_key_text == "private-key-text"
            assert config.certificate_path == "/path/to/cert.pem"
            assert config.private_key_path == "/path/to/key.pem"

    def test_file_extensions_parsing(self):
        """ファイル拡張子のパースのテスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_SITE_NAME": "test",
            "SHAREPOINT_TENANT_ID": "test-tenant-id",
            "SHAREPOINT_CLIENT_ID": "test-client-id",
            "SHAREPOINT_CERTIFICATE_TEXT": "cert",
            "SHAREPOINT_PRIVATE_KEY_TEXT": "key",
            "SHAREPOINT_ALLOWED_FILE_EXTENSIONS": "pdf,docx,xlsx,pptx",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()

            assert config.allowed_file_extensions == ["pdf", "docx", "xlsx", "pptx"]

    def test_default_values(self):
        """デフォルト値のテスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_SITE_NAME": "test",
            "SHAREPOINT_TENANT_ID": "test-tenant-id",
            "SHAREPOINT_CLIENT_ID": "test-client-id",
            "SHAREPOINT_CERTIFICATE_TEXT": "cert",
            "SHAREPOINT_PRIVATE_KEY_TEXT": "key",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()

            assert config.default_max_results == 20
            assert "pdf" in config.allowed_file_extensions
            assert "Search for documents in SharePoint" in config.search_tool_description
            assert "Download a file from SharePoint" in config.download_tool_description

    @pytest.mark.unit
    def test_empty_site_name_creates_tenant_url(self):
        """サイト名が空の場合、テナント全体のURLが作成されることのテスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_SITE_NAME": "",  # Empty site name
            "SHAREPOINT_TENANT_ID": "test-tenant-id",
            "SHAREPOINT_CLIENT_ID": "test-client-id",
            "SHAREPOINT_CERTIFICATE_TEXT": "cert",
            "SHAREPOINT_PRIVATE_KEY_TEXT": "key",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()

            assert config.site_url == "https://test.sharepoint.com"


class TestOneDriveConfig:
    """OneDrive設定のテスト"""

    def test_parse_onedrive_paths_basic(self):
        """基本的なOneDriveパス解析のテスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_ONEDRIVE_PATHS": "user1@company.com,user2@company.com:/Documents/Projects",
            "SHAREPOINT_SITE_NAME": "@onedrive",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            targets = config.parse_onedrive_paths()

            assert len(targets) == 2
            
            # user1@company.com（フォルダー指定なし）
            assert targets[0]["email"] == "user1@company.com"
            assert targets[0]["folder_path"] == ""
            assert targets[0]["onedrive_path"] == "personal/user1_company_com"
            
            # user2@company.com:/Documents/Projects
            assert targets[1]["email"] == "user2@company.com"
            assert targets[1]["folder_path"] == "/Documents/Projects"
            assert targets[1]["onedrive_path"] == "personal/user2_company_com/Documents/Projects"

    def test_parse_onedrive_paths_empty(self):
        """OneDriveパス設定が空の場合のテスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_ONEDRIVE_PATHS": "",
            "SHAREPOINT_SITE_NAME": "@onedrive",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            targets = config.parse_onedrive_paths()

            assert targets == []

    def test_email_to_onedrive_path_conversion(self):
        """メールアドレスからOneDriveパスへの変換テスト"""
        config = SharePointConfig()
        
        # 基本的な変換
        path = config._email_to_onedrive_path("user@company.com")
        assert path == "personal/user_company_com"
        
        # フォルダーパス付き
        path = config._email_to_onedrive_path("user@company.com", "/Documents/Projects")
        assert path == "personal/user_company_com/Documents/Projects"
        
        # 先頭スラッシュの除去
        path = config._email_to_onedrive_path("user@company.com", "Documents/Projects")
        assert path == "personal/user_company_com/Documents/Projects"
        
        # onmicrosoft.com ドメイン
        path = config._email_to_onedrive_path("admin@company.onmicrosoft.com", "/Documents")
        assert path == "personal/admin_company_onmicrosoft_com/Documents"

    def test_include_onedrive_property(self):
        """include_onedriveプロパティのテスト"""
        # OneDriveが含まれる場合
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_ONEDRIVE_PATHS": "user1@company.com",
            "SHAREPOINT_SITE_NAME": "@onedrive,team-site",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            assert config.include_onedrive is True

        # @onedriveがあるがパス設定がない場合
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_ONEDRIVE_PATHS": "",
            "SHAREPOINT_SITE_NAME": "@onedrive,team-site",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            assert config.include_onedrive is False

        # @onedriveがない場合
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_ONEDRIVE_PATHS": "user1@company.com",
            "SHAREPOINT_SITE_NAME": "team-site",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            assert config.include_onedrive is False

    def test_sites_property_excludes_keywords(self):
        """sitesプロパティが特別キーワードを除外することのテスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_SITE_NAME": "@onedrive,team-site,@all,project-alpha",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            sites = config.sites

            assert "team-site" in sites
            assert "project-alpha" in sites
            assert "@onedrive" not in sites
            assert "@all" not in sites

    def test_has_multiple_targets_property(self):
        """has_multiple_targetsプロパティのテスト"""
        # OneDriveが含まれる場合
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_ONEDRIVE_PATHS": "user1@company.com",
            "SHAREPOINT_SITE_NAME": "@onedrive",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            assert config.has_multiple_targets is True

        # 複数サイトの場合
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_SITE_NAME": "site1,site2",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            assert config.has_multiple_targets is True

        # @allの場合
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_SITE_NAME": "@all",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            assert config.has_multiple_targets is True

        # 単一サイトの場合
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_SITE_NAME": "single-site",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            assert config.has_multiple_targets is False

    def test_get_onedrive_targets(self):
        """get_onedrive_targetsメソッドのテスト"""
        # OneDriveが有効な場合
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_ONEDRIVE_PATHS": "user1@company.com,user2@company.com:/Documents/Projects",
            "SHAREPOINT_SITE_NAME": "@onedrive,team-site",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            targets = config.get_onedrive_targets()

            assert len(targets) == 2
            assert targets[0]["email"] == "user1@company.com"
            assert targets[1]["email"] == "user2@company.com"

        # OneDriveが無効な場合
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_ONEDRIVE_PATHS": "user1@company.com",
            "SHAREPOINT_SITE_NAME": "team-site",  # @onedriveなし
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            targets = config.get_onedrive_targets()

            assert targets == []

    def test_invalid_email_format_skipped(self):
        """無効なメールアドレス形式がスキップされることのテスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_ONEDRIVE_PATHS": "invalid-email,user@company.com,another-invalid",
            "SHAREPOINT_SITE_NAME": "@onedrive",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            targets = config.parse_onedrive_paths()

            # 有効なメールアドレスのみが含まれる
            assert len(targets) == 1
            assert targets[0]["email"] == "user@company.com"


class TestDisabledTools:
    """ツール無効化機能のテスト"""

    def test_no_disabled_tools_by_default(self):
        """デフォルトでは全ツールが有効であることのテスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_TENANT_ID": "test-tenant-id",
            "SHAREPOINT_CLIENT_ID": "test-client-id",
            "SHAREPOINT_CERTIFICATE_TEXT": "cert",
            "SHAREPOINT_PRIVATE_KEY_TEXT": "key",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()

            assert config.disabled_tools == set()
            assert config.is_tool_enabled("sharepoint_docs_search") is True
            assert config.is_tool_enabled("sharepoint_docs_download") is True
            assert config.is_tool_enabled("sharepoint_excel") is True

    def test_disable_single_tool(self):
        """単一ツールを無効化するテスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_TENANT_ID": "test-tenant-id",
            "SHAREPOINT_CLIENT_ID": "test-client-id",
            "SHAREPOINT_CERTIFICATE_TEXT": "cert",
            "SHAREPOINT_PRIVATE_KEY_TEXT": "key",
            "SHAREPOINT_DISABLED_TOOLS": "sharepoint_excel",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()

            assert config.disabled_tools == {"sharepoint_excel"}
            assert config.is_tool_enabled("sharepoint_docs_search") is True
            assert config.is_tool_enabled("sharepoint_docs_download") is True
            assert config.is_tool_enabled("sharepoint_excel") is False

    def test_disable_multiple_tools(self):
        """複数ツールを無効化するテスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_TENANT_ID": "test-tenant-id",
            "SHAREPOINT_CLIENT_ID": "test-client-id",
            "SHAREPOINT_CERTIFICATE_TEXT": "cert",
            "SHAREPOINT_PRIVATE_KEY_TEXT": "key",
            "SHAREPOINT_DISABLED_TOOLS": "sharepoint_excel,sharepoint_docs_download",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()

            assert config.disabled_tools == {"sharepoint_excel", "sharepoint_docs_download"}
            assert config.is_tool_enabled("sharepoint_docs_search") is True
            assert config.is_tool_enabled("sharepoint_docs_download") is False
            assert config.is_tool_enabled("sharepoint_excel") is False

    def test_disable_all_tools(self):
        """全ツールを無効化するテスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_TENANT_ID": "test-tenant-id",
            "SHAREPOINT_CLIENT_ID": "test-client-id",
            "SHAREPOINT_CERTIFICATE_TEXT": "cert",
            "SHAREPOINT_PRIVATE_KEY_TEXT": "key",
            "SHAREPOINT_DISABLED_TOOLS": "sharepoint_docs_search,sharepoint_docs_download,sharepoint_excel",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()

            assert config.is_tool_enabled("sharepoint_docs_search") is False
            assert config.is_tool_enabled("sharepoint_docs_download") is False
            assert config.is_tool_enabled("sharepoint_excel") is False

    def test_case_insensitive_tool_names(self):
        """ツール名の大文字小文字を区別しないテスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_TENANT_ID": "test-tenant-id",
            "SHAREPOINT_CLIENT_ID": "test-client-id",
            "SHAREPOINT_CERTIFICATE_TEXT": "cert",
            "SHAREPOINT_PRIVATE_KEY_TEXT": "key",
            "SHAREPOINT_DISABLED_TOOLS": "SHAREPOINT_EXCEL,SharePoint_Docs_Download",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()

            assert config.is_tool_enabled("sharepoint_excel") is False
            assert config.is_tool_enabled("SHAREPOINT_EXCEL") is False
            assert config.is_tool_enabled("sharepoint_docs_download") is False
            assert config.is_tool_enabled("SharePoint_Docs_Download") is False
            assert config.is_tool_enabled("sharepoint_docs_search") is True

    def test_whitespace_handling(self):
        """空白文字の処理テスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_TENANT_ID": "test-tenant-id",
            "SHAREPOINT_CLIENT_ID": "test-client-id",
            "SHAREPOINT_CERTIFICATE_TEXT": "cert",
            "SHAREPOINT_PRIVATE_KEY_TEXT": "key",
            "SHAREPOINT_DISABLED_TOOLS": " sharepoint_excel , sharepoint_docs_download ",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()

            assert config.disabled_tools == {"sharepoint_excel", "sharepoint_docs_download"}
            assert config.is_tool_enabled("sharepoint_excel") is False
            assert config.is_tool_enabled("sharepoint_docs_download") is False
