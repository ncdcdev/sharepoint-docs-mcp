import base64
from unittest.mock import MagicMock, Mock, patch

import pytest
import requests

from src.error_messages import SharePointError
from src.sharepoint_excel import SharePointExcelClient


class TestSharePointExcelClient:
    """SharePoint Excel操作クライアントのテスト"""

    def setup_method(self):
        """テストメソッド実行前のセットアップ"""
        self.mock_auth = MagicMock()
        self.mock_auth.get_access_token.return_value = "test-token"
        self.client = SharePointExcelClient(
            site_url="https://test.sharepoint.com/sites/test-site",
            auth=self.mock_auth,
        )

    def test_initialization(self):
        """クライアント初期化のテスト"""
        assert self.client.site_url == "https://test.sharepoint.com/sites/test-site"
        assert self.client.auth == self.mock_auth

    def test_initialization_strips_trailing_slash(self):
        """サイトURLの末尾スラッシュが削除されることのテスト"""
        client = SharePointExcelClient(
            site_url="https://test.sharepoint.com/sites/test-site/",
            auth=self.mock_auth,
        )
        assert client.site_url == "https://test.sharepoint.com/sites/test-site"

    def test_build_excel_rest_url_basic(self):
        """Excel REST API URLの基本的な構築テスト"""
        file_path = "https://test.sharepoint.com/sites/test-site/Shared Documents/test.xlsx"
        url = self.client._build_excel_rest_url(file_path, "Sheets", "atom")

        expected_url = (
            "https://test.sharepoint.com/sites/test-site/_vti_bin/ExcelRest.aspx/"
            "Shared%20Documents/test.xlsx/Model/Sheets?$format=atom"
        )
        assert url == expected_url

    def test_build_excel_rest_url_with_special_chars(self):
        """特殊文字を含むファイルパスのURL構築テスト"""
        file_path = "https://test.sharepoint.com/sites/test-site/Shared Documents/Test File (1).xlsx"
        url = self.client._build_excel_rest_url(file_path, "Sheets", "atom")

        # スペースと括弧がエンコードされることを確認
        assert "Test%20File%20%281%29.xlsx" in url
        assert "Shared%20Documents" in url

    def test_build_excel_rest_url_nested_folder(self):
        """ネストされたフォルダのURL構築テスト"""
        file_path = "https://test.sharepoint.com/sites/test-site/Documents/Reports/2024/Q1/data.xlsx"
        url = self.client._build_excel_rest_url(file_path, "Ranges('Sheet1!A1:C10')", "atom")

        expected_url = (
            "https://test.sharepoint.com/sites/test-site/_vti_bin/ExcelRest.aspx/"
            "Documents/Reports/2024/Q1/data.xlsx/Model/Ranges('Sheet1!A1:C10')?$format=atom"
        )
        assert url == expected_url

    def test_build_excel_rest_url_invalid_path(self):
        """無効なファイルパスでエラーが発生することのテスト"""
        file_path = "https://test.sharepoint.com/invalid/path/test.xlsx"

        with pytest.raises(ValueError) as exc_info:
            self.client._build_excel_rest_url(file_path, "Sheets", "atom")

        assert "Invalid file path format" in str(exc_info.value)

    @patch("src.sharepoint_excel.requests.get")
    def test_list_sheets_success(self, mock_get):
        """シート一覧取得の成功テスト"""
        # モックレスポンス
        mock_response = Mock()
        mock_response.text = '<?xml version="1.0"?><sheets><sheet>Sheet1</sheet></sheets>'
        mock_response.raise_for_status = Mock()
        mock_get.return_value = mock_response

        file_path = "https://test.sharepoint.com/sites/test-site/Shared Documents/test.xlsx"
        result = self.client.list_sheets(file_path)

        # 結果検証
        assert result == mock_response.text
        assert mock_get.called
        assert mock_get.call_args[1]["headers"]["Authorization"] == "Bearer test-token"
        assert mock_get.call_args[1]["headers"]["Accept"] == "application/atom+xml"

    @patch("src.sharepoint_excel.requests.get")
    def test_list_sheets_http_error(self, mock_get):
        """シート一覧取得でHTTPエラーが発生するテスト"""
        # HTTPエラーをシミュレート
        mock_response = Mock()
        mock_response.status_code = 404
        mock_response.raise_for_status.side_effect = requests.HTTPError("404 Not Found")
        mock_get.return_value = mock_response

        file_path = "https://test.sharepoint.com/sites/test-site/Shared Documents/test.xlsx"

        with pytest.raises(SharePointError):
            self.client.list_sheets(file_path)

    @patch("src.sharepoint_excel.requests.get")
    def test_get_sheet_image_success(self, mock_get):
        """シート画像取得の成功テスト"""
        # モックレスポンス（画像データ）
        mock_image_data = b"fake-image-data"
        mock_response = Mock()
        mock_response.content = mock_image_data
        mock_response.raise_for_status = Mock()
        mock_get.return_value = mock_response

        file_path = "https://test.sharepoint.com/sites/test-site/Shared Documents/test.xlsx"
        result = self.client.get_sheet_image(file_path, "Sheet1")

        # 結果検証（base64エンコードされていることを確認）
        expected_base64 = base64.b64encode(mock_image_data).decode("utf-8")
        assert result == expected_base64
        assert mock_get.called
        assert mock_get.call_args[1]["headers"]["Accept"] == "image/png"

    @patch("src.sharepoint_excel.requests.get")
    def test_get_sheet_image_with_special_chars(self, mock_get):
        """特殊文字を含むシート名の画像取得テスト"""
        mock_image_data = b"fake-image-data"
        mock_response = Mock()
        mock_response.content = mock_image_data
        mock_response.raise_for_status = Mock()
        mock_get.return_value = mock_response

        file_path = "https://test.sharepoint.com/sites/test-site/Shared Documents/test.xlsx"
        # シングルクォートを含むシート名
        result = self.client.get_sheet_image(file_path, "John's Sheet")

        # URLにシングルクォートがエスケープされて含まれることを確認
        assert mock_get.called
        called_url = mock_get.call_args[0][0]
        # シングルクォートが2つのシングルクォートにエスケープされる
        assert "John''s Sheet" in called_url

    @patch("src.sharepoint_excel.requests.get")
    def test_get_range_data_success(self, mock_get):
        """セル範囲データ取得の成功テスト"""
        mock_response = Mock()
        mock_response.text = '<?xml version="1.0"?><range><cell>A1</cell></range>'
        mock_response.raise_for_status = Mock()
        mock_get.return_value = mock_response

        file_path = "https://test.sharepoint.com/sites/test-site/Shared Documents/test.xlsx"
        result = self.client.get_range_data(file_path, "Sheet1!A1:C10")

        # 結果検証
        assert result == mock_response.text
        assert mock_get.called
        assert mock_get.call_args[1]["headers"]["Accept"] == "application/atom+xml"

    @patch("src.sharepoint_excel.requests.get")
    def test_get_range_data_with_quotes(self, mock_get):
        """シングルクォートを含む範囲指定のテスト"""
        mock_response = Mock()
        mock_response.text = '<?xml version="1.0"?><range><cell>A1</cell></range>'
        mock_response.raise_for_status = Mock()
        mock_get.return_value = mock_response

        file_path = "https://test.sharepoint.com/sites/test-site/Shared Documents/test.xlsx"
        # シート名にシングルクォートを含む範囲指定
        result = self.client.get_range_data(file_path, "John's Sheet!A1:C10")

        # URLにシングルクォートがエスケープされて含まれることを確認
        assert mock_get.called
        called_url = mock_get.call_args[0][0]
        assert "John''s Sheet!A1:C10" in called_url

    @patch("src.sharepoint_excel.requests.get")
    def test_get_range_data_http_error(self, mock_get):
        """セル範囲データ取得でHTTPエラーが発生するテスト"""
        mock_response = Mock()
        mock_response.status_code = 400
        mock_response.raise_for_status.side_effect = requests.HTTPError("400 Bad Request")
        mock_get.return_value = mock_response

        file_path = "https://test.sharepoint.com/sites/test-site/Shared Documents/test.xlsx"

        with pytest.raises(SharePointError):
            self.client.get_range_data(file_path, "InvalidRange")

    def test_auth_token_used_in_requests(self):
        """認証トークンがリクエストに使用されることのテスト"""
        with patch("src.sharepoint_excel.requests.get") as mock_get:
            mock_response = Mock()
            mock_response.text = "test"
            mock_response.raise_for_status = Mock()
            mock_get.return_value = mock_response

            file_path = "https://test.sharepoint.com/sites/test-site/Shared Documents/test.xlsx"
            self.client.list_sheets(file_path)

            # get_access_tokenが呼ばれたことを確認
            self.mock_auth.get_access_token.assert_called_once()

            # 正しいAuthorizationヘッダーが設定されていることを確認
            headers = mock_get.call_args[1]["headers"]
            assert headers["Authorization"] == "Bearer test-token"
