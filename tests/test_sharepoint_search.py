import pytest
from unittest.mock import patch, MagicMock
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
