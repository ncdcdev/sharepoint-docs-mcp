import zipfile
from unittest.mock import Mock

import pytest
import requests

from src.error_messages import (
    ErrorCategory,
    SharePointError,
    handle_sharepoint_error,
)


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
        assert (
            "configuration" in str(result).lower()
            or "administrator" in str(result).lower()
        )

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
        assert (
            "configuration" in str(result).lower()
            or "administrator" in str(result).lower()
        )

    @pytest.mark.unit
    def test_certificate_error(self):
        """証明書関連エラーのテスト"""
        cert_error = Exception("Certificate file not found")

        result = handle_sharepoint_error(cert_error, "search")

        assert "certificate" in str(result).lower()
        assert (
            "configured" in str(result).lower()
            or "configuration" in str(result).lower()
        )

    @pytest.mark.unit
    def test_generic_error(self):
        """一般的なエラーのテスト"""
        generic_error = Exception("Something went wrong")

        result = handle_sharepoint_error(generic_error, "search")

        assert "unexpected error" in str(result).lower()
        assert (
            "configuration" in str(result).lower()
            or "administrator" in str(result).lower()
        )

    @pytest.mark.unit
    def test_download_context_specific_message(self):
        """ダウンロード操作固有のエラーメッセージテスト"""
        generic_error = Exception("Download failed")

        result = handle_sharepoint_error(generic_error, "download")

        assert "error" in str(result).lower()
        assert (
            "configuration" in str(result).lower()
            or "administrator" in str(result).lower()
        )

    @pytest.mark.unit
    def test_http_error_without_response(self):
        """レスポンスオブジェクトがないHTTPエラーのテスト"""
        http_error = requests.HTTPError("HTTP Error without response")

        result = handle_sharepoint_error(http_error, "search")

        assert "error" in str(result).lower()
        assert (
            "configuration" in str(result).lower()
            or "administrator" in str(result).lower()
        )


class TestExcelErrorClassification:
    """Excel操作のエラー分類テスト"""

    @pytest.mark.unit
    def test_excel_file_not_found_http_404(self):
        """HTTP 404エラーのExcelコンテキストでの分類テスト"""
        mock_response = Mock()
        mock_response.status_code = 404
        mock_response.text = "Not Found"

        http_error = requests.HTTPError()
        http_error.response = mock_response

        result = handle_sharepoint_error(
            http_error,
            "excel_parse",
            excel_context={"file_path": "/sites/test/test.xlsx"},
        )

        assert result.category == ErrorCategory.EXCEL_FILE_NOT_FOUND
        assert "test.xlsx" in str(result)

    @pytest.mark.unit
    def test_excel_sheet_not_found(self):
        """シートが見つからないエラーの分類テスト"""
        error = ValueError(
            "Sheet 'NonExistent' not found. Available sheets: ['Sheet1']"
        )

        result = handle_sharepoint_error(
            error,
            "excel_parse",
            excel_context={"sheet_name": "NonExistent"},
        )

        assert result.category == ErrorCategory.EXCEL_SHEET_NOT_FOUND
        assert "NonExistent" in str(result)

    @pytest.mark.unit
    def test_excel_invalid_file_badzip(self):
        """無効なExcelファイル（BadZipFile）のエラー分類テスト"""
        error = zipfile.BadZipFile("File is not a zip file")

        result = handle_sharepoint_error(error, "excel_parse")

        assert result.category == ErrorCategory.EXCEL_INVALID_FILE
        assert "valid" in str(result).lower()

    @pytest.mark.unit
    def test_excel_invalid_file_corrupt(self):
        """破損したExcelファイルのエラー分類テスト"""
        error = Exception("This file is not a valid xlsx file or is corrupted")

        result = handle_sharepoint_error(error, "excel_parse")

        assert result.category == ErrorCategory.EXCEL_INVALID_FILE

    @pytest.mark.unit
    def test_sharepoint_error_not_rewrapped(self):
        """SharePointErrorが再ラップされないことのテスト"""
        original_error = SharePointError(
            category=ErrorCategory.EXCEL_SHEET_NOT_FOUND,
            message="Original message",
            solution="Original solution",
        )

        result = handle_sharepoint_error(original_error, "excel_parse")

        # 同じオブジェクトが返される
        assert result is original_error
        assert result.category == ErrorCategory.EXCEL_SHEET_NOT_FOUND
        assert result.message == "Original message"

    @pytest.mark.unit
    def test_excel_context_preserves_file_path(self):
        """excel_contextのfile_pathが保持されることのテスト"""
        mock_response = Mock()
        mock_response.status_code = 404
        mock_response.text = "Not Found"

        http_error = requests.HTTPError()
        http_error.response = mock_response

        file_path = "/sites/finance/Shared Documents/budget.xlsx"
        result = handle_sharepoint_error(
            http_error,
            "excel_parse",
            excel_context={"file_path": file_path},
        )

        assert file_path in str(result)
