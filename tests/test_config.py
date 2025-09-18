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