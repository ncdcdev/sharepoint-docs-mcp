import pytest
from unittest.mock import patch, Mock
import base64

from src.server import sharepoint_docs_search, sharepoint_docs_download, sharepoint_excel_to_json


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


class TestSharePointExcelToJson:
    """sharepoint_excel_to_json 関数のテスト"""

    @pytest.fixture
    def mock_excel_parser(self):
        """Mock Excel parser"""
        with patch("src.server.SharePointExcelParser") as mock_parser_class:
            parser_instance = Mock()
            parser_instance.parse_to_json.return_value = '{"file_path": "/test.xlsx", "sheets": []}'
            mock_parser_class.return_value = parser_instance
            yield parser_instance

    @pytest.mark.unit
    def test_excel_to_json_success(self, mock_config, mock_sharepoint_client, mock_excel_parser):
        """Excel to JSON変換の成功テスト（デフォルト）"""
        with patch("src.server._get_sharepoint_client", return_value=mock_sharepoint_client):
            with patch("src.server.config", mock_config):
                result = sharepoint_excel_to_json(
                    file_path="/sites/test/Shared Documents/test.xlsx"
                )

                # JSON文字列が返されることを確認
                assert '"file_path"' in result
                assert '"sheets"' in result
                # デフォルトでは include_formatting=False が渡される
                mock_excel_parser.parse_to_json.assert_called_once_with(
                    "/sites/test/Shared Documents/test.xlsx", False
                )

    @pytest.mark.unit
    def test_excel_to_json_with_real_json(self, mock_config, mock_sharepoint_client, mock_excel_parser):
        """実際のJSON構造での変換テスト"""
        import json

        mock_json_data = {
            "file_path": "/sites/test/Shared Documents/test.xlsx",
            "sheets": [
                {
                    "name": "Sheet1",
                    "dimensions": "A1:B2",
                    "rows": [
                        [
                            {"value": "Name", "data_type": "s", "coordinate": "A1"},
                            {"value": "Age", "data_type": "s", "coordinate": "B1"}
                        ],
                        [
                            {"value": "John", "data_type": "s", "coordinate": "A2"},
                            {"value": 25, "data_type": "n", "coordinate": "B2"}
                        ]
                    ]
                }
            ]
        }
        mock_excel_parser.parse_to_json.return_value = json.dumps(mock_json_data)

        with patch("src.server._get_sharepoint_client", return_value=mock_sharepoint_client):
            with patch("src.server.config", mock_config):
                result = sharepoint_excel_to_json(
                    file_path="/sites/test/Shared Documents/test.xlsx"
                )

                # JSON文字列をパース
                parsed_result = json.loads(result)
                assert parsed_result["file_path"] == "/sites/test/Shared Documents/test.xlsx"
                assert len(parsed_result["sheets"]) == 1
                assert parsed_result["sheets"][0]["name"] == "Sheet1"

    @pytest.mark.unit
    def test_excel_to_json_with_formatting(self, mock_config, mock_sharepoint_client, mock_excel_parser):
        """Excel to JSON変換の成功テスト（書式情報あり）"""
        with patch("src.server._get_sharepoint_client", return_value=mock_sharepoint_client):
            with patch("src.server.config", mock_config):
                result = sharepoint_excel_to_json(
                    file_path="/sites/test/Shared Documents/test.xlsx",
                    include_formatting=True
                )

                # JSON文字列が返されることを確認
                assert '"file_path"' in result
                assert '"sheets"' in result
                # include_formatting=True が渡される
                mock_excel_parser.parse_to_json.assert_called_once_with(
                    "/sites/test/Shared Documents/test.xlsx", True
                )

    @pytest.mark.unit
    def test_excel_to_json_error_handling(self, mock_config, mock_sharepoint_client):
        """Excel to JSON変換のエラーハンドリングテスト"""
        from src.error_messages import SharePointError

        with patch("src.server._get_sharepoint_client", return_value=mock_sharepoint_client):
            with patch("src.server.SharePointExcelParser") as mock_parser_class:
                parser_instance = Mock()
                parser_instance.parse_to_json.side_effect = Exception("Parse error")
                mock_parser_class.return_value = parser_instance

                with patch("src.server.config", mock_config):
                    with pytest.raises(SharePointError):
                        sharepoint_excel_to_json(
                            file_path="/sites/test/Shared Documents/test.xlsx"
                        )