"""
SharePoint検索機能モジュール
"""

import logging
from typing import Any

from office365.runtime.auth.token_response import TokenResponse
from office365.sharepoint.client_context import ClientContext

from .sharepoint_auth import SharePointCertificateAuth

logger = logging.getLogger(__name__)


class SharePointSearchClient:
    """SharePoint検索クライアント"""

    def __init__(self, site_url: str, auth: SharePointCertificateAuth):
        self.site_url = site_url.rstrip("/")
        self.auth = auth

    def _get_sp_context(self) -> ClientContext:
        """SharePointコンテキストを取得"""
        def token_func():
            """トークンコールバック関数"""
            access_token = self.auth.get_access_token()
            if not access_token:
                raise ValueError("Failed to get access token with certificate")
            return TokenResponse(access_token=access_token, token_type="Bearer")

        context = ClientContext(self.site_url)
        context.with_access_token(token_func)
        return context

    def search_documents(
        self,
        query: str,
        max_results: int = 20,
        file_extensions: list[str] | None = None,
    ) -> list[dict[str, Any]]:
        """
        SharePointでドキュメントを検索

        Args:
            query: 検索クエリ
            max_results: 最大結果数
            file_extensions: 検索対象のファイル拡張子（例: ['pdf', 'docx']）

        Returns:
            検索結果のリスト
        """
        logger.info(f"Searching for documents containing: {query}")

        sp_context = self._get_sp_context()

        # 検索クエリの構築 - まずは単純なクエリから開始
        search_query = query

        # ファイル拡張子フィルターを追加
        if file_extensions:
            ext_filter = " OR ".join(
                [f"fileextension:{ext}" for ext in file_extensions]
            )
            search_query += f" AND ({ext_filter})"

        try:
            # まず基本的な接続をテスト
            logger.info("Testing basic SharePoint connection...")
            try:
                test_response = sp_context.web.context.execute_request_direct("web/title")
                logger.info(f"Basic connection test successful: {test_response}")
            except Exception as e:
                logger.warning(f"Basic connection test failed: {e}")

            # SharePoint Search APIを正しい形式で呼び出し
            access_token = self.auth.get_access_token()

            import requests

            # SharePoint REST APIの正しい構文（パラメータを単一引用符で囲む）
            params = {
                "querytext": f"'{search_query}'",  # 単一引用符で囲む
                "selectproperties": "'Title,Path,Size,LastModifiedTime,FileExtension,HitHighlightedSummary'"  # 単一引用符で囲む
            }

            search_url = f"{self.site_url}/_api/search/query"

            headers = {
                'Accept': 'application/json;odata=verbose',
                'Authorization': f'Bearer {access_token}'
            }

            logger.info(f"Sending GET request to: {search_url}")
            logger.info(f"Params: {params}")

            response = requests.get(search_url, params=params, headers=headers, timeout=30)

            logger.info(f"Response status: {response.status_code}")
            if response.status_code != 200:
                logger.error(f"Response content: {response.text}")

            response.raise_for_status()
            search_results_json = response.json()

            logger.info(f"Response JSON structure: {list(search_results_json.keys()) if isinstance(search_results_json, dict) else type(search_results_json)}")
            logger.info(f"Content of 'd': {search_results_json.get('d', 'Not found')}")

            results = []
            # JSONレスポンスの解析 - SharePoint OData形式
            if isinstance(search_results_json, dict) and 'd' in search_results_json:
                d_content = search_results_json['d']
                logger.info(f"Type of 'd' content: {type(d_content)}")

                if isinstance(d_content, dict):
                    logger.info(f"Keys in 'd': {list(d_content.keys())}")

                    # OData v3形式
                    if 'PrimaryQueryResult' in d_content:
                        primary_results = d_content['PrimaryQueryResult']
                        relevant_results = primary_results.get('RelevantResults', {})
                    # クエリ結果が直接d以下にある場合
                    elif 'query' in d_content:
                        primary_results = d_content['query'].get('PrimaryQueryResult', {})
                        relevant_results = primary_results.get('RelevantResults', {})
                    else:
                        # その他の構造の場合、デバッグ情報を出力
                        logger.info(f"Unexpected response structure in 'd': {list(d_content.keys())}")
                        relevant_results = {}
                else:
                    logger.error(f"'d' is not a dict but {type(d_content)}: {d_content}")
                    relevant_results = {}

                # レスポンス解析の修正 - 'Rows'の中に'results'配列がある
                if 'Table' in relevant_results and 'Rows' in relevant_results['Table']:
                    rows = relevant_results['Table']['Rows'].get('results', [])
                    for row in rows:
                        cells = row.get('Cells', {}).get('results', [])
                        result_item = {}

                        # セル情報を解析
                        for cell in cells:
                            key = cell['Key']
                            value = cell['Value']

                            if key == 'Title':
                                result_item['title'] = value
                            elif key == 'Path':
                                result_item['path'] = value
                            elif key == 'Size':
                                result_item['size'] = value
                            elif key == 'LastModifiedTime':
                                result_item['modified'] = value
                            elif key == 'FileExtension':
                                result_item['extension'] = value
                            elif key == 'HitHighlightedSummary':
                                result_item['summary'] = value

                        if result_item:
                            results.append(result_item)

            logger.info(f"Found {len(results)} search results")
            return results

        except Exception as e:
            logger.error(f"Search failed: {str(e)}")
            return [{"error": f"Search failed: {str(e)}"}]
