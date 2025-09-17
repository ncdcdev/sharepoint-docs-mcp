"""
SharePoint検索機能モジュール
"""

import logging
from typing import Any

import requests

from .sharepoint_auth import SharePointCertificateAuth

logger = logging.getLogger(__name__)


class SharePointSearchClient:
    """SharePoint検索クライアント"""

    def __init__(self, site_url: str, auth: SharePointCertificateAuth):
        self.site_url = site_url.rstrip("/")
        self.auth = auth

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

        # 検索クエリの構築
        search_query = query

        # ファイル拡張子フィルターを追加
        if file_extensions:
            ext_filter = " OR ".join(
                [f"fileextension:{ext}" for ext in file_extensions]
            )
            search_query += f" AND ({ext_filter})"

        try:
            # アクセストークンを取得
            access_token = self.auth.get_access_token()

            # SharePoint REST APIの正しい構文（パラメータを単一引用符で囲む）
            params = {
                "querytext": f"'{search_query}'",
                "selectproperties": "'Title,Path,Size,LastModifiedTime,FileExtension,HitHighlightedSummary'",
            }

            search_url = f"{self.site_url}/_api/search/query"

            headers = {
                "Accept": "application/json;odata=verbose",
                "Authorization": f"Bearer {access_token}",
            }

            response = requests.get(
                search_url, params=params, headers=headers, timeout=30
            )
            response.raise_for_status()
            search_results_json = response.json()

            results = []
            # JSONレスポンスの解析
            if isinstance(search_results_json, dict) and "d" in search_results_json:
                d_content = search_results_json["d"]
                if isinstance(d_content, dict) and "query" in d_content:
                    primary_results = d_content["query"].get("PrimaryQueryResult", {})
                    relevant_results = primary_results.get("RelevantResults", {})

                    # SharePointレスポンス構造に合わせて解析
                    if (
                        "Table" in relevant_results
                        and "Rows" in relevant_results["Table"]
                    ):
                        rows = relevant_results["Table"]["Rows"].get("results", [])
                        for row in rows:
                            cells = row.get("Cells", {}).get("results", [])
                            result_item = {}

                            # セル情報を解析
                            for cell in cells:
                                key = cell.get("Key", "")
                                value = cell.get("Value", "")

                                if key == "Title":
                                    result_item["title"] = value
                                elif key == "Path":
                                    result_item["path"] = value
                                elif key == "Size":
                                    result_item["size"] = value
                                elif key == "LastModifiedTime":
                                    result_item["modified"] = value
                                elif key == "FileExtension":
                                    result_item["extension"] = value
                                elif key == "HitHighlightedSummary":
                                    result_item["summary"] = value

                            if result_item:
                                results.append(result_item)

            logger.info(f"Found {len(results)} search results")
            return results

        except Exception as e:
            logger.error(f"Search failed: {str(e)}")
            raise
