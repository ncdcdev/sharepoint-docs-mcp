import pytest
from unittest.mock import patch, MagicMock, Mock
import os

from src.sharepoint_search import SharePointSearchClient
from src.config import SharePointConfig


class TestSharePointSearchOneDrive:
    """SharePoint検索のOneDrive機能テスト"""

    def setup_method(self):
        """テストメソッド実行前のセットアップ"""
        self.mock_auth = MagicMock()
        self.mock_auth.get_access_token.return_value = "test-token"
        self.client = SharePointSearchClient(
            site_url="https://test.sharepoint.com",
            auth=self.mock_auth
        )

    def test_build_onedrive_filters(self):
        """OneDriveフィルター構築のテスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_ONEDRIVE_PATHS": "user1@company.com,user2@company.com:/Documents/Projects",
            "SHAREPOINT_SITE_NAME": "@onedrive",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            filters = self.client._build_onedrive_filters(config)

            expected_filters = [
                'path:"https://test-my.sharepoint.com/personal/user1_company_com"',
                'path:"https://test-my.sharepoint.com/personal/user2_company_com/Documents/Projects"'
            ]
            assert filters == expected_filters

    def test_build_onedrive_filters_empty(self):
        """OneDriveが無効な場合のフィルター構築テスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_SITE_NAME": "team-site",  # @onedriveなし
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            filters = self.client._build_onedrive_filters(config)

            assert filters == []

    def test_build_sharepoint_filters_multiple_sites(self):
        """複数SharePointサイトのフィルター構築テスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_SITE_NAME": "site1,site2,site3",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            filters = self.client._build_sharepoint_filters(config)

            expected_filters = [
                'site:"https://test.sharepoint.com/sites/site1"',
                'site:"https://test.sharepoint.com/sites/site2"',
                'site:"https://test.sharepoint.com/sites/site3"'
            ]
            assert filters == expected_filters

    def test_build_site_filters_mixed(self):
        """OneDriveとSharePointサイトの混合フィルター構築テスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_ONEDRIVE_PATHS": "user1@company.com:/Documents/Projects",
            "SHAREPOINT_SITE_NAME": "@onedrive,team-site,project-alpha",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            site_filters = self.client._build_site_filters(config)

            # OneDriveフィルターとSharePointフィルターが含まれることを確認
            assert 'path:"https://test-my.sharepoint.com/personal/user1_company_com/Documents/Projects"' in site_filters
            assert 'site:"https://test.sharepoint.com/sites/team-site"' in site_filters
            assert 'site:"https://test.sharepoint.com/sites/project-alpha"' in site_filters

    def test_build_search_query_with_onedrive(self):
        """OneDriveを含む検索クエリ構築のテスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_ONEDRIVE_PATHS": "user1@company.com",
            "SHAREPOINT_SITE_NAME": "@onedrive,team-site",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            search_query = self.client._build_search_query("test query", config)

            # 基本クエリとフィルターが結合されていることを確認
            assert search_query.startswith("test query AND (")
            assert 'path:"https://test-my.sharepoint.com/personal/user1_company_com"' in search_query
            assert 'site:"https://test.sharepoint.com/sites/team-site"' in search_query

    def test_build_search_query_no_filters(self):
        """フィルターなしの場合の検索クエリ構築テスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_SITE_NAME": "",  # サイト指定なし
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            search_query = self.client._build_search_query("test query", config)

            # フィルターが追加されないことを確認
            assert search_query == "test query"

    def test_site_specific_behavior_unchanged(self):
        """既存の単一サイト指定動作が変更されていないことのテスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_SITE_NAME": "single-site",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            
            # 従来の動作確認
            assert config.is_site_specific is True
            assert config.has_multiple_targets is False
            
            filters = self.client._build_sharepoint_filters(config)
            assert len(filters) == 1
            assert 'site:"https://test.sharepoint.com/sites/single-site"' in filters[0]

    def test_special_keywords_handling(self):
        """特別キーワード（@onedrive, @all）の処理テスト"""
        # @all の場合
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_SITE_NAME": "@all",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            config = SharePointConfig()
            
            assert config.sites == []  # @allは通常サイトリストに含まれない
            assert config.has_multiple_targets is True  # @allは複数対象扱い
            
            # @allの場合、SharePointフィルターは空になる（テナント全体検索）
            filters = self.client._build_sharepoint_filters(config)
            assert filters == []


class TestSharePointUpload:
    """SharePointアップロード機能テスト"""

    def setup_method(self):
        """テストメソッド実行前のセットアップ"""
        self.mock_auth = MagicMock()
        self.mock_auth.get_access_token.return_value = "test-token"
        self.client = SharePointSearchClient(
            site_url="https://test.sharepoint.com",
            auth=self.mock_auth
        )
        # モック設定
        self.mock_config = Mock()
        self.mock_config.base_url = "https://test.sharepoint.com"

    def test_parse_site_path(self):
        """サイト指定パスの解析テスト"""
        with patch("src.sharepoint_search.global_config", self.mock_config):
            api_base_url, server_relative_folder = self.client._parse_site_path(
                "TeamSite:/Shared Documents/Reports"
            )

            assert api_base_url == "https://test.sharepoint.com/sites/TeamSite"
            assert server_relative_folder == "/sites/TeamSite/Shared Documents/Reports"

    def test_parse_onedrive_path(self):
        """OneDrive指定パスの解析テスト"""
        with patch("src.sharepoint_search.global_config", self.mock_config):
            api_base_url, server_relative_folder = self.client._parse_onedrive_path(
                "@onedrive:user@company.com:/Documents/Projects"
            )

            assert api_base_url == "https://test-my.sharepoint.com/personal/user_company_com"
            assert server_relative_folder == "/personal/user_company_com/Documents/Projects"

    def test_parse_full_url_path_sharepoint(self):
        """完全URL（SharePoint）の解析テスト"""
        with patch("src.sharepoint_search.global_config", self.mock_config):
            api_base_url, server_relative_folder = self.client._parse_full_url_path(
                "https://test.sharepoint.com/sites/TeamSite/Shared Documents/Reports"
            )

            assert api_base_url == "https://test.sharepoint.com/sites/TeamSite"
            assert server_relative_folder == "/sites/TeamSite/Shared Documents/Reports"

    def test_parse_full_url_path_onedrive(self):
        """完全URL（OneDrive）の解析テスト"""
        with patch("src.sharepoint_search.global_config", self.mock_config):
            api_base_url, server_relative_folder = self.client._parse_full_url_path(
                "https://test-my.sharepoint.com/personal/user_company_com/Documents"
            )

            assert api_base_url == "https://test-my.sharepoint.com/personal/user_company_com"
            assert server_relative_folder == "/personal/user_company_com/Documents"

    def test_parse_folder_path_site_format(self):
        """_parse_folder_path - サイト形式のテスト"""
        with patch("src.sharepoint_search.global_config", self.mock_config):
            api_base_url, server_relative_folder = self.client._parse_folder_path(
                "TestSite:/Documents"
            )

            assert api_base_url == "https://test.sharepoint.com/sites/TestSite"
            assert server_relative_folder == "/sites/TestSite/Documents"

    def test_parse_folder_path_onedrive_format(self):
        """_parse_folder_path - OneDrive形式のテスト"""
        with patch("src.sharepoint_search.global_config", self.mock_config):
            api_base_url, server_relative_folder = self.client._parse_folder_path(
                "@onedrive:test@example.com:/Folder"
            )

            assert api_base_url == "https://test-my.sharepoint.com/personal/test_example_com"
            assert server_relative_folder == "/personal/test_example_com/Folder"

    def test_parse_folder_path_full_url(self):
        """_parse_folder_path - 完全URL形式のテスト"""
        with patch("src.sharepoint_search.global_config", self.mock_config):
            api_base_url, server_relative_folder = self.client._parse_folder_path(
                "https://test.sharepoint.com/sites/MySite/Library"
            )

            assert api_base_url == "https://test.sharepoint.com/sites/MySite"
            assert server_relative_folder == "/sites/MySite/Library"

    def test_parse_folder_path_invalid_format(self):
        """無効なフォルダパス形式のテスト"""
        with patch("src.sharepoint_search.global_config", self.mock_config):
            with pytest.raises(ValueError) as excinfo:
                self.client._parse_folder_path("invalid-path-without-colon")

            assert "Invalid folder_path format" in str(excinfo.value)

    def test_parse_onedrive_path_invalid_format(self):
        """無効なOneDriveパス形式のテスト"""
        with patch("src.sharepoint_search.global_config", self.mock_config):
            with pytest.raises(ValueError) as excinfo:
                self.client._parse_onedrive_path("@onedrive:invalid-format")

            assert "Invalid OneDrive path format" in str(excinfo.value)

    def test_build_full_url(self):
        """完全URL構築のテスト"""
        result = self.client._build_full_url(
            "https://test.sharepoint.com/sites/TeamSite",
            "/sites/TeamSite/Documents/file.pdf"
        )
        assert result == "https://test.sharepoint.com/sites/TeamSite/Documents/file.pdf"

    def test_upload_file_invalid_filename(self):
        """無効なファイル名のテスト（パストラバーサル防止）"""
        # 相対パスを含むファイル名
        with pytest.raises(ValueError) as excinfo:
            self.client.upload_file(
                file_content=b"test content",
                file_name="../malicious.txt",
                folder_path="TestSite:/Documents"
            )
        assert "Invalid file name" in str(excinfo.value)

        # スラッシュを含むファイル名
        with pytest.raises(ValueError) as excinfo:
            self.client.upload_file(
                file_content=b"test content",
                file_name="path/to/file.txt",
                folder_path="TestSite:/Documents"
            )
        assert "Invalid file name" in str(excinfo.value)

    @patch('src.sharepoint_search.requests.post')
    def test_upload_file_success(self, mock_post):
        """ファイルアップロード成功のテスト"""
        # モックレスポンスの設定
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "d": {
                "Name": "test.txt",
                "ServerRelativeUrl": "/sites/TeamSite/Documents/test.txt",
                "Length": 12,
                "TimeLastModified": "2025-01-15T10:00:00Z"
            }
        }
        mock_response.raise_for_status = MagicMock()
        mock_post.return_value = mock_response

        with patch("src.sharepoint_search.global_config", self.mock_config):
            result = self.client.upload_file(
                file_content=b"test content",
                file_name="test.txt",
                folder_path="TeamSite:/Documents"
            )

            assert result["title"] == "test.txt"
            assert result["path"] == "https://test.sharepoint.com/sites/TeamSite/Documents/test.txt"
            assert result["size"] == "12"
            assert result["extension"] == "txt"

            # APIが正しく呼び出されたことを確認
            mock_post.assert_called_once()
            call_args = mock_post.call_args
            assert "/_api/web/GetFolderByServerRelativeUrl" in call_args[0][0]
            assert "Files/add" in call_args[0][0]

    @patch('src.sharepoint_search.requests.post')
    def test_upload_file_with_overwrite(self, mock_post):
        """上書きオプション付きアップロードのテスト"""
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "d": {
                "Name": "test.txt",
                "ServerRelativeUrl": "/sites/TeamSite/Documents/test.txt",
                "Length": 12,
                "TimeLastModified": "2025-01-15T10:00:00Z"
            }
        }
        mock_response.raise_for_status = MagicMock()
        mock_post.return_value = mock_response

        with patch("src.sharepoint_search.global_config", self.mock_config):
            self.client.upload_file(
                file_content=b"test content",
                file_name="test.txt",
                folder_path="TeamSite:/Documents",
                overwrite=True
            )

            # overwrite=trueがURLに含まれていることを確認
            call_args = mock_post.call_args
            assert "overwrite=true" in call_args[0][0]
