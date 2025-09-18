import pytest
from unittest.mock import Mock
import requests

from src.error_messages import handle_sharepoint_error


class TestHandleSharePointError:
    """handle_sharepoint_error 関数のテスト"""

    @pytest.mark.unit
    def test_http_error_401(self):
        """401 Unauthorized エラーのテスト"""
        mock_response = Mock()
        mock_response.status_code = 401
        mock_response.text = "Unauthorized"

        http_error = requests.HTTPError()
        http_error.response = mock_response

        result = handle_sharepoint_error(http_error, "search")

        assert "authentication" in str(result).lower()
        assert "configuration" in str(result).lower()

    @pytest.mark.unit
    def test_http_error_403(self):
        """403 Forbidden エラーのテスト"""
        mock_response = Mock()
        mock_response.status_code = 403
        mock_response.text = "Forbidden"

        http_error = requests.HTTPError()
        http_error.response = mock_response

        result = handle_sharepoint_error(http_error, "search")

        assert "access" in str(result).lower() or "denied" in str(result).lower()
        assert "permission" in str(result).lower()

    @pytest.mark.unit
    def test_http_error_404(self):
        """404 Not Found エラーのテスト"""
        mock_response = Mock()
        mock_response.status_code = 404
        mock_response.text = "Not Found"

        http_error = requests.HTTPError()
        http_error.response = mock_response

        result = handle_sharepoint_error(http_error, "search")

        assert "error" in str(result).lower()
        assert "configuration" in str(result).lower() or "administrator" in str(result).lower()

    @pytest.mark.unit
    def test_connection_error(self):
        """接続エラーのテスト"""
        connection_error = requests.ConnectionError("Connection failed")

        result = handle_sharepoint_error(connection_error, "search")

        assert "connection" in str(result).lower() or "network" in str(result).lower()
        assert "sharepoint" in str(result).lower()

    @pytest.mark.unit
    def test_timeout_error(self):
        """タイムアウトエラーのテスト"""
        timeout_error = requests.Timeout("Request timed out")

        result = handle_sharepoint_error(timeout_error, "search")

        assert "error" in str(result).lower()
        assert "configuration" in str(result).lower() or "administrator" in str(result).lower()

    @pytest.mark.unit
    def test_certificate_error(self):
        """証明書関連エラーのテスト"""
        cert_error = Exception("Certificate file not found")

        result = handle_sharepoint_error(cert_error, "search")

        assert "certificate" in str(result).lower()
        assert "configured" in str(result).lower() or "configuration" in str(result).lower()

    @pytest.mark.unit
    def test_generic_error(self):
        """一般的なエラーのテスト"""
        generic_error = Exception("Something went wrong")

        result = handle_sharepoint_error(generic_error, "search")

        assert "unexpected error" in str(result).lower()
        assert "configuration" in str(result).lower() or "administrator" in str(result).lower()

    @pytest.mark.unit
    def test_download_context_specific_message(self):
        """ダウンロード操作固有のエラーメッセージテスト"""
        generic_error = Exception("Download failed")

        result = handle_sharepoint_error(generic_error, "download")

        assert "error" in str(result).lower()
        assert "configuration" in str(result).lower() or "administrator" in str(result).lower()

    @pytest.mark.unit
    def test_http_error_without_response(self):
        """レスポンスオブジェクトがないHTTPエラーのテスト"""
        http_error = requests.HTTPError("HTTP Error without response")

        result = handle_sharepoint_error(http_error, "search")

        assert "error" in str(result).lower()
        assert "configuration" in str(result).lower() or "administrator" in str(result).lower()