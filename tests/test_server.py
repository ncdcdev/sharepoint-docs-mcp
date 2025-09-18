import pytest
from unittest.mock import patch, Mock
import base64

from src.server import sharepoint_docs_search, sharepoint_docs_download


class TestSharePointDocsSearch:
    """sharepoint_docs_search 関数のテスト"""

    @pytest.mark.unit
    def test_search_with_default_parameters(self, mock_config, mock_sharepoint_client):
        """デフォルトパラメータでの検索テスト"""
        with patch("src.server._get_sharepoint_client", return_value=mock_sharepoint_client):
            with patch("src.server.config", mock_config):
                results = sharepoint_docs_search("test query")

                assert len(results) == 1
                assert results[0]["title"] == "Test Document 1"
                assert results[0]["path"] == "/sites/test/documents/test1.pdf"
                mock_sharepoint_client.search_documents.assert_called_once_with(
                    query="test query",
                    max_results=20,
                    file_extensions=None,
                )

    @pytest.mark.unit
    def test_search_with_compact_format(self, mock_config, mock_sharepoint_client):
        """コンパクトフォーマットでの検索テスト"""
        with patch("src.server._get_sharepoint_client", return_value=mock_sharepoint_client):
            with patch("src.server.config", mock_config):
                results = sharepoint_docs_search("test query", response_format="compact")

                assert len(results) == 1
                assert "title" in results[0]
                assert "path" in results[0]
                assert "extension" in results[0]
                # コンパクトフォーマットでは以下のフィールドは含まれない
                assert "size" not in results[0]
                assert "modified" not in results[0]
                assert "summary" not in results[0]

    @pytest.mark.unit
    def test_search_with_invalid_response_format(self, mock_config, mock_sharepoint_client):
        """無効なresponse_formatでの検索テスト（デフォルトにフォールバック）"""
        with patch("src.server._get_sharepoint_client", return_value=mock_sharepoint_client):
            with patch("src.server.config", mock_config):
                results = sharepoint_docs_search("test query", response_format="invalid")

                # 無効なフォーマットはdetailedにフォールバックするため、全フィールドが含まれる
                assert len(results) == 1
                assert "title" in results[0]
                assert "size" in results[0]
                assert "modified" in results[0]

    @pytest.mark.unit
    def test_search_with_file_extensions(self, mock_config, mock_sharepoint_client):
        """ファイル拡張子フィルタでの検索テスト"""
        with patch("src.server._get_sharepoint_client", return_value=mock_sharepoint_client):
            with patch("src.server.config", mock_config):
                results = sharepoint_docs_search(
                    "test query",
                    file_extensions=["pdf", "docx"]
                )

                mock_sharepoint_client.search_documents.assert_called_once_with(
                    query="test query",
                    max_results=20,
                    file_extensions=["pdf", "docx"],
                )

    @pytest.mark.unit
    def test_search_max_results_limit(self, mock_config, mock_sharepoint_client):
        """最大結果数の制限テスト"""
        with patch("src.server._get_sharepoint_client", return_value=mock_sharepoint_client):
            with patch("src.server.config", mock_config):
                sharepoint_docs_search("test query", max_results=150)

                # 100を超える値は100に制限される
                mock_sharepoint_client.search_documents.assert_called_once_with(
                    query="test query",
                    max_results=100,
                    file_extensions=None,
                )


class TestSharePointDocsDownload:
    """sharepoint_docs_download 関数のテスト"""

    @pytest.mark.unit
    def test_download_file(self, mock_config, mock_sharepoint_client):
        """ファイルダウンロードのテスト"""
        with patch("src.server._get_sharepoint_client", return_value=mock_sharepoint_client):
            with patch("src.server.config", mock_config):
                result = sharepoint_docs_download("/sites/test/documents/test.pdf")

                expected_content = base64.b64encode(b"mock file content").decode("utf-8")
                assert result == expected_content
                mock_sharepoint_client.download_file.assert_called_once_with(
                    "/sites/test/documents/test.pdf"
                )

    @pytest.mark.unit
    def test_download_file_error_handling(self, mock_config, mock_sharepoint_client):
        """ファイルダウンロードエラーハンドリングのテスト"""
        mock_sharepoint_client.download_file.side_effect = Exception("Download failed")

        with patch("src.server._get_sharepoint_client", return_value=mock_sharepoint_client):
            with patch("src.server.config", mock_config):
                with pytest.raises(Exception) as exc_info:
                    sharepoint_docs_download("/sites/test/documents/test.pdf")

                # エラーハンドリング関数が呼ばれることを確認
                assert "Download failed" in str(exc_info.value.__cause__)


class TestGetSharePointClient:
    """_get_sharepoint_client 関数のテスト"""

    @pytest.mark.unit
    def test_client_initialization(self, mock_config):
        """SharePointクライアントの初期化テスト"""
        with patch("src.server.config", mock_config):
            with patch("src.server.SharePointCertificateAuth") as mock_auth_class:
                with patch("src.server.SharePointSearchClient") as mock_client_class:
                    from src.server import _get_sharepoint_client

                    # グローバル変数をリセット
                    import src.server
                    src.server._sharepoint_client = None

                    client = _get_sharepoint_client()

                    mock_auth_class.assert_called_once_with(
                        tenant_id=mock_config.tenant_id,
                        client_id=mock_config.client_id,
                        site_url=mock_config.site_url,
                        certificate_path=mock_config.certificate_path,
                        certificate_text=mock_config.certificate_text,
                        private_key_path=mock_config.private_key_path,
                        private_key_text=mock_config.private_key_text,
                    )
                    mock_client_class.assert_called_once()

    @pytest.mark.unit
    def test_client_singleton_behavior(self, mock_config):
        """SharePointクライアントのシングルトン動作テスト"""
        with patch("src.server.config", mock_config):
            with patch("src.server.SharePointCertificateAuth"):
                with patch("src.server.SharePointSearchClient") as mock_client_class:
                    from src.server import _get_sharepoint_client

                    # グローバル変数をリセット
                    import src.server
                    src.server._sharepoint_client = None

                    client1 = _get_sharepoint_client()
                    client2 = _get_sharepoint_client()

                    # 2回目の呼び出しでは新しいインスタンスを作成しない
                    assert mock_client_class.call_count == 1
                    assert client1 == client2