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

    def _make_request(
        self, url: str, params: dict[str, Any] | None = None
    ) -> dict[str, Any]:
        """SharePoint APIにリクエストを送信"""
        access_token = self.auth.get_access_token()
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json",
            "Content-Type": "application/json",
        }

        logger.info(f"Making request to SharePoint API: {url}")
        response = requests.get(url, headers=headers, params=params, timeout=30)
        response.raise_for_status()

        return response.json()

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
        # 検索APIエンドポイント
        search_url = f"{self.site_url}/_api/search/query"

        # 検索クエリの構築
        search_query = f"'{query}' AND path:{self.site_url}"

        # ファイル拡張子フィルターを追加
        if file_extensions:
            ext_filter = " OR ".join(
                [f"fileextension:{ext}" for ext in file_extensions]
            )
            search_query += f" AND ({ext_filter})"

        # 検索パラメータ
        params = {
            "querytext": search_query,
            "rowlimit": max_results,
            "selectproperties": "Title,Path,Size,LastModifiedTime,FileExtension,HitHighlightedSummary,Author",
            "trimduplicates": "false",
        }

        logger.info(f"Searching SharePoint with query: {search_query}")

        try:
            result = self._make_request(search_url, params)

            # 検索結果を解析
            search_results = []
            primary_results = (
                result.get("d", {}).get("query", {}).get("PrimaryQueryResult", {})
            )
            relevant_results = primary_results.get("RelevantResults", {})
            rows = relevant_results.get("Table", {}).get("Rows", {}).get("results", [])

            for row in rows:
                cells = row.get("Cells", {}).get("results", [])
                doc_info = {}

                # セルからプロパティを抽出
                for cell in cells:
                    key = cell.get("Key", "")
                    value = cell.get("Value", "")

                    if key == "Title":
                        doc_info["title"] = value
                    elif key == "Path":
                        doc_info["path"] = value
                    elif key == "Size":
                        doc_info["size"] = int(value) if value else 0
                    elif key == "LastModifiedTime":
                        doc_info["last_modified"] = value
                    elif key == "FileExtension":
                        doc_info["file_extension"] = value
                    elif key == "HitHighlightedSummary":
                        doc_info["summary"] = value
                    elif key == "Author":
                        doc_info["author"] = value

                if doc_info.get("path"):  # パスが存在する場合のみ結果に含める
                    search_results.append(doc_info)

            logger.info(f"Found {len(search_results)} documents")
            return search_results

        except requests.RequestException as e:
            logger.error(f"SharePoint search request failed: {e}")
            raise
        except Exception as e:
            logger.error(f"SharePoint search failed: {e}")
            raise
