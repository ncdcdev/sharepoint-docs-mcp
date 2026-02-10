import base64
import os
from unittest.mock import Mock, patch

import pytest

from src.server import (
    register_tools,
    sharepoint_docs_download,
    sharepoint_docs_search,
    sharepoint_excel,
)


class TestSharePointDocsSearch:
    """sharepoint_docs_search 関数のテスト"""

    @pytest.mark.unit
    def test_search_with_default_parameters(self, mock_config, mock_sharepoint_client):
        """デフォルトパラメータでの検索テスト"""
        with patch(
            "src.server._get_sharepoint_client", return_value=mock_sharepoint_client
        ):
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
        with patch(
            "src.server._get_sharepoint_client", return_value=mock_sharepoint_client
        ):
            with patch("src.server.config", mock_config):
                results = sharepoint_docs_search(
                    "test query", response_format="compact"
                )

                assert len(results) == 1
                assert "title" in results[0]
                assert "path" in results[0]
                assert "extension" in results[0]
                # コンパクトフォーマットでは以下のフィールドは含まれない
                assert "size" not in results[0]
                assert "modified" not in results[0]
                assert "summary" not in results[0]

    @pytest.mark.unit
    def test_search_with_invalid_response_format(
        self, mock_config, mock_sharepoint_client
    ):
        """無効なresponse_formatでの検索テスト（デフォルトにフォールバック）"""
        with patch(
            "src.server._get_sharepoint_client", return_value=mock_sharepoint_client
        ):
            with patch("src.server.config", mock_config):
                results = sharepoint_docs_search(
                    "test query", response_format="invalid"
                )

                # 無効なフォーマットはdetailedにフォールバックするため、全フィールドが含まれる
                assert len(results) == 1
                assert "title" in results[0]
                assert "size" in results[0]
                assert "modified" in results[0]

    @pytest.mark.unit
    def test_search_with_file_extensions(self, mock_config, mock_sharepoint_client):
        """ファイル拡張子フィルタでの検索テスト"""
        with patch(
            "src.server._get_sharepoint_client", return_value=mock_sharepoint_client
        ):
            with patch("src.server.config", mock_config):
                sharepoint_docs_search("test query", file_extensions=["pdf", "docx"])

                mock_sharepoint_client.search_documents.assert_called_once_with(
                    query="test query",
                    max_results=20,
                    file_extensions=["pdf", "docx"],
                )

    @pytest.mark.unit
    def test_search_max_results_limit(self, mock_config, mock_sharepoint_client):
        """最大結果数の制限テスト"""
        with patch(
            "src.server._get_sharepoint_client", return_value=mock_sharepoint_client
        ):
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
        with patch(
            "src.server._get_sharepoint_client", return_value=mock_sharepoint_client
        ):
            with patch("src.server.config", mock_config):
                result = sharepoint_docs_download("/sites/test/documents/test.pdf")

                expected_content = base64.b64encode(b"mock file content").decode(
                    "utf-8"
                )
                assert result == expected_content
                mock_sharepoint_client.download_file.assert_called_once_with(
                    "/sites/test/documents/test.pdf"
                )

    @pytest.mark.unit
    def test_download_file_error_handling(self, mock_config, mock_sharepoint_client):
        """ファイルダウンロードエラーハンドリングのテスト"""
        mock_sharepoint_client.download_file.side_effect = Exception("Download failed")

        with patch(
            "src.server._get_sharepoint_client", return_value=mock_sharepoint_client
        ):
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
                    # グローバル変数をリセット
                    import src.server
                    from src.server import _get_sharepoint_client

                    src.server._sharepoint_client = None

                    _get_sharepoint_client()

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
                    # グローバル変数をリセット
                    import src.server
                    from src.server import _get_sharepoint_client

                    src.server._sharepoint_client = None

                    client1 = _get_sharepoint_client()
                    client2 = _get_sharepoint_client()

                    # 2回目の呼び出しでは新しいインスタンスを作成しない
                    assert mock_client_class.call_count == 1
                    assert client1 == client2


class TestSharePointExcel:
    """sharepoint_excel 関数のテスト"""

    @pytest.fixture
    def mock_excel_parser(self):
        """Mock Excel parser"""
        with patch("src.server.SharePointExcelParser") as mock_parser_class:
            parser_instance = Mock()
            parser_instance.parse_to_json.return_value = (
                '{"file_path": "/test.xlsx", "sheets": []}'
            )
            parser_instance.search_cells.return_value = '{"file_path": "/test.xlsx", "mode": "search", "query": "test", "match_count": 0, "matches": []}'
            mock_parser_class.return_value = parser_instance
            yield parser_instance

    @pytest.mark.unit
    def test_excel_read_default(
        self, mock_config, mock_sharepoint_client, mock_excel_parser
    ):
        """Excelデータ取得の成功テスト（デフォルト）"""
        with patch(
            "src.server._get_sharepoint_client", return_value=mock_sharepoint_client
        ):
            with patch("src.server.config", mock_config):
                result = sharepoint_excel(
                    file_path="/sites/test/Shared Documents/test.xlsx"
                )

                # JSON文字列が返されることを確認
                assert '"file_path"' in result
                assert '"sheets"' in result
                # デフォルトでは全パラメータがデフォルト値で呼ばれる
                mock_excel_parser.parse_to_json.assert_called_once_with(
                    "/sites/test/Shared Documents/test.xlsx",
                    sheet_name=None,
                    cell_range=None,
                )

    @pytest.mark.unit
    def test_excel_search_mode(
        self, mock_config, mock_sharepoint_client, mock_excel_parser
    ):
        """Excel検索モードのテスト"""
        with patch(
            "src.server._get_sharepoint_client", return_value=mock_sharepoint_client
        ):
            with patch("src.server.config", mock_config):
                sharepoint_excel(
                    file_path="/sites/test/Shared Documents/test.xlsx", query="売上"
                )

                # 検索メソッドが呼ばれることを確認
                mock_excel_parser.search_cells.assert_called_once_with(
                    "/sites/test/Shared Documents/test.xlsx", "売上", sheet_name=None
                )
                # parse_to_jsonは呼ばれない
                mock_excel_parser.parse_to_json.assert_not_called()

    @pytest.mark.unit
    def test_excel_with_sheet_parameter(
        self, mock_config, mock_sharepoint_client, mock_excel_parser
    ):
        """シート指定パラメータのテスト"""
        with patch(
            "src.server._get_sharepoint_client", return_value=mock_sharepoint_client
        ):
            with patch("src.server.config", mock_config):
                sharepoint_excel(
                    file_path="/sites/test/Shared Documents/test.xlsx", sheet="Sheet2"
                )

                mock_excel_parser.parse_to_json.assert_called_once_with(
                    "/sites/test/Shared Documents/test.xlsx",
                    sheet_name="Sheet2",
                    cell_range=None,
                )

    @pytest.mark.unit
    def test_excel_with_cell_range_parameter(
        self, mock_config, mock_sharepoint_client, mock_excel_parser
    ):
        """セル範囲指定パラメータのテスト"""
        with patch(
            "src.server._get_sharepoint_client", return_value=mock_sharepoint_client
        ):
            with patch("src.server.config", mock_config):
                sharepoint_excel(
                    file_path="/sites/test/Shared Documents/test.xlsx",
                    sheet="Sheet1",
                    cell_range="A1:D10",
                )

                mock_excel_parser.parse_to_json.assert_called_once_with(
                    "/sites/test/Shared Documents/test.xlsx",
                    sheet_name="Sheet1",
                    cell_range="A1:D10",
                )

    @pytest.mark.unit
    def test_excel_with_real_json(
        self, mock_config, mock_sharepoint_client, mock_excel_parser
    ):
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
                            {"value": "Name", "coordinate": "A1"},
                            {"value": "Age", "coordinate": "B1"},
                        ],
                        [
                            {"value": "John", "coordinate": "A2"},
                            {"value": 25, "coordinate": "B2"},
                        ],
                    ],
                }
            ],
        }
        mock_excel_parser.parse_to_json.return_value = json.dumps(mock_json_data)

        with patch(
            "src.server._get_sharepoint_client", return_value=mock_sharepoint_client
        ):
            with patch("src.server.config", mock_config):
                result = sharepoint_excel(
                    file_path="/sites/test/Shared Documents/test.xlsx"
                )

                # JSON文字列をパース
                parsed_result = json.loads(result)
                assert (
                    parsed_result["file_path"]
                    == "/sites/test/Shared Documents/test.xlsx"
                )
                assert len(parsed_result["sheets"]) == 1
                assert parsed_result["sheets"][0]["name"] == "Sheet1"

    @pytest.mark.unit
    def test_excel_search_with_real_json(
        self, mock_config, mock_sharepoint_client, mock_excel_parser
    ):
        """検索モードの実際のJSON構造テスト"""
        import json

        mock_search_result = {
            "file_path": "/sites/test/Shared Documents/test.xlsx",
            "mode": "search",
            "query": "売上",
            "match_count": 2,
            "matches": [
                {"sheet": "Sheet1", "coordinate": "A1", "value": "売上実績"},
                {"sheet": "Sheet1", "coordinate": "B5", "value": "月間売上"},
            ],
        }
        mock_excel_parser.search_cells.return_value = json.dumps(mock_search_result)

        with patch(
            "src.server._get_sharepoint_client", return_value=mock_sharepoint_client
        ):
            with patch("src.server.config", mock_config):
                result = sharepoint_excel(
                    file_path="/sites/test/Shared Documents/test.xlsx", query="売上"
                )

                # JSON文字列をパース
                parsed_result = json.loads(result)
                assert parsed_result["mode"] == "search"
                assert parsed_result["query"] == "売上"
                assert parsed_result["match_count"] == 2
                assert len(parsed_result["matches"]) == 2

    @pytest.mark.unit
    def test_excel_error_handling(self, mock_config, mock_sharepoint_client):
        """エラーハンドリングテスト"""
        from src.error_messages import SharePointError

        with patch(
            "src.server._get_sharepoint_client", return_value=mock_sharepoint_client
        ):
            with patch("src.server.SharePointExcelParser") as mock_parser_class:
                parser_instance = Mock()
                parser_instance.parse_to_json.side_effect = Exception("Parse error")
                mock_parser_class.return_value = parser_instance

                with patch("src.server.config", mock_config):
                    with pytest.raises(SharePointError):
                        sharepoint_excel(
                            file_path="/sites/test/Shared Documents/test.xlsx"
                        )


class TestRegisterTools:
    """register_tools 関数のテスト"""

    @pytest.mark.unit
    def test_all_tools_registered_by_default(self):
        """デフォルトでは全ツールが登録されることのテスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_TENANT_ID": "test-tenant-id",
            "SHAREPOINT_CLIENT_ID": "test-client-id",
            "SHAREPOINT_CERTIFICATE_TEXT": "cert",
            "SHAREPOINT_PRIVATE_KEY_TEXT": "key",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            with patch("src.server.mcp") as mock_mcp:
                mock_tool_decorator = Mock(return_value=lambda f: f)
                mock_mcp.tool.return_value = mock_tool_decorator

                # Configを再読み込みして環境変数を反映
                from importlib import reload

                import src.config

                reload(src.config)

                with patch("src.server.config", src.config.config):
                    register_tools()

                # mcp.tool が3回呼ばれることを確認
                assert mock_mcp.tool.call_count == 3

    @pytest.mark.unit
    def test_single_tool_disabled(self):
        """単一ツールが無効化された場合のテスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_TENANT_ID": "test-tenant-id",
            "SHAREPOINT_CLIENT_ID": "test-client-id",
            "SHAREPOINT_CERTIFICATE_TEXT": "cert",
            "SHAREPOINT_PRIVATE_KEY_TEXT": "key",
            "SHAREPOINT_DISABLED_TOOLS": "sharepoint_excel",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            with patch("src.server.mcp") as mock_mcp:
                mock_tool_decorator = Mock(return_value=lambda f: f)
                mock_mcp.tool.return_value = mock_tool_decorator

                # Configを再読み込みして環境変数を反映
                from importlib import reload

                import src.config

                reload(src.config)

                with patch("src.server.config", src.config.config):
                    register_tools()

                # mcp.tool が2回呼ばれることを確認（sharepoint_excelは除外）
                assert mock_mcp.tool.call_count == 2

    @pytest.mark.unit
    def test_multiple_tools_disabled(self):
        """複数ツールが無効化された場合のテスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_TENANT_ID": "test-tenant-id",
            "SHAREPOINT_CLIENT_ID": "test-client-id",
            "SHAREPOINT_CERTIFICATE_TEXT": "cert",
            "SHAREPOINT_PRIVATE_KEY_TEXT": "key",
            "SHAREPOINT_DISABLED_TOOLS": "sharepoint_excel,sharepoint_docs_download",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            with patch("src.server.mcp") as mock_mcp:
                mock_tool_decorator = Mock(return_value=lambda f: f)
                mock_mcp.tool.return_value = mock_tool_decorator

                # Configを再読み込みして環境変数を反映
                from importlib import reload

                import src.config

                reload(src.config)

                with patch("src.server.config", src.config.config):
                    register_tools()

                # mcp.tool が1回呼ばれることを確認（sharepoint_docs_searchのみ）
                assert mock_mcp.tool.call_count == 1

    @pytest.mark.unit
    def test_all_tools_disabled(self):
        """全ツールが無効化された場合のテスト"""
        env_vars = {
            "SHAREPOINT_BASE_URL": "https://test.sharepoint.com",
            "SHAREPOINT_TENANT_ID": "test-tenant-id",
            "SHAREPOINT_CLIENT_ID": "test-client-id",
            "SHAREPOINT_CERTIFICATE_TEXT": "cert",
            "SHAREPOINT_PRIVATE_KEY_TEXT": "key",
            "SHAREPOINT_DISABLED_TOOLS": "sharepoint_docs_search,sharepoint_docs_download,sharepoint_excel",
        }

        with patch.dict(os.environ, env_vars, clear=True):
            with patch("src.server.mcp") as mock_mcp:
                mock_tool_decorator = Mock(return_value=lambda f: f)
                mock_mcp.tool.return_value = mock_tool_decorator

                # Configを再読み込みして環境変数を反映
                from importlib import reload

                import src.config

                reload(src.config)

                with patch("src.server.config", src.config.config):
                    register_tools()

                # mcp.tool が0回呼ばれることを確認
                assert mock_mcp.tool.call_count == 0
