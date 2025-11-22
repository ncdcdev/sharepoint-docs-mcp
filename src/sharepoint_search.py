"""
SharePoint検索機能モジュール
"""

import logging
import re
from typing import Any, Protocol
from urllib.parse import quote, unquote, urlparse

import requests

from src.config import config as global_config
from src.error_messages import handle_sharepoint_error

logger = logging.getLogger(__name__)


class AuthClient(Protocol):
    """認証クライアントのプロトコル（証明書認証/OAuth両対応）"""

    def get_access_token(self) -> str:
        """アクセストークンを取得"""
        ...


class SharePointSearchClient:
    """SharePoint検索クライアント"""

    def __init__(self, site_url: str, auth: AuthClient):
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
        search_query = self._build_search_query(query, global_config)
        logger.info(f"Built search query: {search_query}")

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

            # OneDrive検索を含む場合はベースURL、サイト固有検索の場合はサイトURLを使用
            if not global_config.is_site_specific:
                search_url = f"{global_config.base_url}/_api/search/query"
            else:
                search_url = f"{self.site_url}/_api/search/query"

            logger.info(f"Search URL: {search_url}")

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
                    total_rows = relevant_results.get("TotalRows", 0)
                    logger.info(f"Total rows from SharePoint: {total_rows}")

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
            raise handle_sharepoint_error(e, "search") from e

    def _build_search_query(self, query: str, config) -> str:
        """検索クエリを構築（OneDriveと複数サイト対応）"""
        search_query = query

        # サイトフィルターを構築
        site_filters = self._build_site_filters(config)

        if site_filters:
            search_query += f" AND ({site_filters})"

        return search_query

    def _build_site_filters(self, config) -> str:
        """サイトフィルターを構築"""
        filters = []

        # OneDriveフィルターを追加
        onedrive_filters = self._build_onedrive_filters(config)
        filters.extend(onedrive_filters)

        # SharePointサイトフィルターを追加
        sharepoint_filters = self._build_sharepoint_filters(config)
        filters.extend(sharepoint_filters)

        return " OR ".join(filters) if filters else ""

    def _build_onedrive_filters(self, config) -> list[str]:
        """OneDrive用のフィルターを構築"""
        if not config.include_onedrive:
            return []

        filters = []
        onedrive_targets = config.get_onedrive_targets()

        # OneDriveのベースURLを構築（-myサフィックス付きドメイン）
        onedrive_base_url = config.base_url.replace(
            ".sharepoint.com", "-my.sharepoint.com"
        )

        for target in onedrive_targets:
            onedrive_path = target["onedrive_path"]
            full_path = f"{onedrive_base_url}/{onedrive_path}"
            filters.append(f'path:"{full_path}"')

        return filters

    def _build_sharepoint_filters(self, config) -> list[str]:
        """SharePointサイト用のフィルターを構築"""
        filters = []

        # サイト指定がある場合
        if config.sites:
            for site_name in config.sites:
                site_url = f"{config.base_url}/sites/{site_name}"
                filters.append(f'site:"{site_url}"')

        return filters

    def download_file(self, file_path: str) -> bytes:
        """
        SharePointからファイルをダウンロード

        Args:
            file_path: ファイルのフルパス（search_documentsの結果から取得）

        Returns:
            ファイルの内容（bytes）
        """
        logger.info(f"Downloading file: {file_path}")

        try:
            # アクセストークンを取得
            access_token = self.auth.get_access_token()

            headers = {
                "Authorization": f"Bearer {access_token}",
                "Accept": "application/octet-stream",  # ファイルバイナリを要求
            }

            # SharePointのファイルパスからサーバー相対URLを抽出
            parsed_url = urlparse(file_path)
            server_relative_url = unquote(parsed_url.path)

            # ファイルのパスから適切なサイトURLを決定
            # OneDriveファイルかどうかを判定
            path_segments = server_relative_url.split("/")
            is_onedrive_file = (
                len(path_segments) >= 2 and path_segments[1] == "personal"
            )

            if is_onedrive_file:
                # OneDriveファイルの場合は個人用サイトのAPIエンドポイントを使用
                onedrive_base_url = global_config.base_url.replace(
                    ".sharepoint.com", "-my.sharepoint.com"
                )
                if len(path_segments) >= 3:
                    personal_site_name = path_segments[2]
                    api_base_url = f"{onedrive_base_url}/personal/{personal_site_name}"
                else:
                    # 通常は発生しないが、フォールバックとして-my.sharepoint.comドメインを使用
                    api_base_url = onedrive_base_url
            elif global_config.is_site_specific:
                # 特定サイト設定の場合はそのサイトのAPIを使用
                api_base_url = self.site_url
            else:
                # テナント全体設定の場合はファイルパスからサイトを特定
                if len(path_segments) >= 3 and path_segments[1] == "sites":
                    site_name = path_segments[2]
                    api_base_url = f"{global_config.base_url}/sites/{site_name}"
                else:
                    # サイト形式でない場合はベースURLを使用
                    api_base_url = global_config.base_url

            logger.info(f"Downloading from: {api_base_url}")

            # SharePointとOneDriveで異なるダウンロード方式を使用
            if is_onedrive_file:
                # OneDrive用：GetFileByServerRelativePath（特殊文字対応）を優先
                return self._download_onedrive_file(
                    api_base_url, server_relative_url, headers
                )
            else:
                # SharePoint用：GetFileByServerRelativeUrlを優先
                return self._download_sharepoint_file(
                    api_base_url, server_relative_url, headers
                )

        except Exception as e:
            logger.error(f"File download failed: {str(e)}")
            # OneDriveファイルかどうかを判定してエラーメッセージを調整
            raise handle_sharepoint_error(
                e, "download", is_onedrive_file=is_onedrive_file
            ) from e

    def _download_onedrive_file(
        self, api_base_url: str, server_relative_url: str, headers: dict
    ) -> bytes:
        """
        OneDriveファイルのダウンロード
        特殊文字対応のGetFileByServerRelativePathを優先し、失敗時にGetFileByServerRelativeUrlにフォールバック
        """
        # 方式1: GetFileByServerRelativePath（特殊文字に強い）
        try:
            # シングルクォートをエスケープ（SharePoint REST API仕様）
            escaped_path = server_relative_url.replace("'", "''")
            encoded_path = quote(escaped_path, safe="/")
            download_url = f"{api_base_url}/_api/web/GetFileByServerRelativePath(decodedUrl=@f)/$value?@f='{encoded_path}'"
            response = requests.get(download_url, headers=headers, timeout=60)
            response.raise_for_status()
            return response.content
        except Exception as e:
            logger.debug(f"GetFileByServerRelativePath failed: {str(e)}")

        # 方式2: GetFileByServerRelativeUrl（フォールバック）
        try:
            # シングルクォートをエスケープ（SharePoint REST API仕様）
            escaped_path = server_relative_url.replace("'", "''")
            download_url = f"{api_base_url}/_api/web/GetFileByServerRelativeUrl('{escaped_path}')/$value"
            response = requests.get(download_url, headers=headers, timeout=60)
            response.raise_for_status()
            return response.content
        except Exception as e:
            logger.error(f"All OneDrive download methods failed: {str(e)}")
            raise

    def _download_sharepoint_file(
        self, api_base_url: str, server_relative_url: str, headers: dict
    ) -> bytes:
        """
        SharePointファイルのダウンロード
        GetFileByServerRelativeUrlを優先し、失敗時にGetFileByServerRelativePathにフォールバック
        """
        # 方式1: GetFileByServerRelativeUrl（標準API）
        try:
            # シングルクォートをエスケープ（SharePoint REST API仕様）
            escaped_path = server_relative_url.replace("'", "''")
            download_url = f"{api_base_url}/_api/web/GetFileByServerRelativeUrl('{escaped_path}')/$value"
            response = requests.get(download_url, headers=headers, timeout=60)
            response.raise_for_status()
            return response.content
        except Exception as e:
            logger.debug(f"GetFileByServerRelativeUrl failed: {str(e)}")

        # 方式2: GetFileByServerRelativePath（フォールバック）
        try:
            # シングルクォートをエスケープ（SharePoint REST API仕様）
            escaped_path = server_relative_url.replace("'", "''")
            encoded_path = quote(escaped_path, safe="/")
            download_url = f"{api_base_url}/_api/web/GetFileByServerRelativePath(decodedUrl=@f)/$value?@f='{encoded_path}'"
            response = requests.get(download_url, headers=headers, timeout=60)
            response.raise_for_status()
            return response.content
        except Exception as e:
            logger.error(f"All SharePoint download methods failed: {str(e)}")
            raise

    def upload_file(
        self,
        file_content: bytes,
        file_name: str,
        folder_path: str,
        overwrite: bool = False,
    ) -> dict[str, Any]:
        """
        SharePointまたはOneDriveにファイルをアップロード

        Args:
            file_content: アップロードするファイルの内容（bytes）
            file_name: アップロード後のファイル名
            folder_path: アップロード先フォルダ（以下の形式をサポート）
                - 完全URL: "https://tenant.sharepoint.com/sites/Site/Folder"
                - サイト指定: "SiteName:/Folder/Path"
                - OneDrive指定: "@onedrive:user@domain.com:/Folder"
            overwrite: 既存ファイルの上書きを許可するか

        Returns:
            アップロード結果（検索結果と同じ形式）
        """
        logger.info(f"Uploading file: {file_name} to {folder_path}")

        # ファイル名の検証（パストラバーサル防止）
        if ".." in file_name or "/" in file_name or "\\" in file_name:
            raise ValueError(
                f"Invalid file name: {file_name}. "
                "File name must not contain path separators or '..'."
            )

        try:
            # アクセストークンを取得
            access_token = self.auth.get_access_token()

            # folder_pathを解析してAPIエンドポイントとサーバー相対パスを取得
            api_base_url, server_relative_folder = self._parse_folder_path(folder_path)

            logger.info(
                f"Uploading to: {api_base_url}, folder: {server_relative_folder}"
            )

            headers = {
                "Authorization": f"Bearer {access_token}",
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/octet-stream",
            }

            # シングルクォートをエスケープ（SharePoint REST API仕様）
            escaped_folder = server_relative_folder.replace("'", "''")
            overwrite_str = "true" if overwrite else "false"

            # アップロードURL構築
            upload_url = (
                f"{api_base_url}/_api/web/GetFolderByServerRelativeUrl('{escaped_folder}')"
                f"/Files/add(url='{file_name}',overwrite={overwrite_str})"
            )

            response = requests.post(
                upload_url, headers=headers, data=file_content, timeout=60
            )
            response.raise_for_status()
            result_json = response.json()

            # レスポンスを検索結果と同じ形式に変換
            file_info = result_json.get("d", {})
            server_relative_url = file_info.get("ServerRelativeUrl", "")

            # 完全URLを構築
            full_url = self._build_full_url(api_base_url, server_relative_url)

            result = {
                "title": file_info.get("Name", file_name),
                "path": full_url,
                "size": str(file_info.get("Length", len(file_content))),
                "modified": file_info.get("TimeLastModified", ""),
                "extension": file_name.split(".")[-1] if "." in file_name else "",
            }

            logger.info(f"File uploaded successfully: {result['path']}")
            return result

        except Exception as e:
            logger.error(f"File upload failed: {str(e)}")
            # OneDriveかどうかを判定してエラーメッセージを調整
            is_onedrive = folder_path.startswith("@onedrive:") or (
                "-my.sharepoint.com" in folder_path
            )
            raise handle_sharepoint_error(
                e, "upload", is_onedrive_file=is_onedrive
            ) from e

    def _parse_folder_path(self, folder_path: str) -> tuple[str, str]:
        """
        folder_pathを解析してAPIベースURLとサーバー相対パスを返す

        Args:
            folder_path: フォルダパス（複数形式対応）

        Returns:
            (api_base_url, server_relative_folder) のタプル
        """
        # OneDrive指定: @onedrive:user@domain.com:/Folder
        if folder_path.startswith("@onedrive:"):
            return self._parse_onedrive_path(folder_path)

        # 完全URL: https://tenant.sharepoint.com/...
        if folder_path.startswith("http://") or folder_path.startswith("https://"):
            return self._parse_full_url_path(folder_path)

        # サイト指定: SiteName:/Folder/Path
        if ":" in folder_path:
            return self._parse_site_path(folder_path)

        # それ以外はエラー
        raise ValueError(
            f"Invalid folder_path format: {folder_path}. "
            "Use 'SiteName:/Folder', '@onedrive:user@domain.com:/Folder', "
            "or full URL."
        )

    def _parse_onedrive_path(self, folder_path: str) -> tuple[str, str]:
        """OneDrive指定のパスを解析"""
        # @onedrive:user@domain.com:/Folder/Path
        match = re.match(r"@onedrive:([^:]+):(.+)", folder_path)
        if not match:
            raise ValueError(
                f"Invalid OneDrive path format: {folder_path}. "
                "Expected: @onedrive:user@domain.com:/Folder"
            )

        email = match.group(1)
        folder = match.group(2)

        # メールアドレスをOneDriveパスに変換
        # user@domain.com -> user_domain_com
        onedrive_user = email.replace("@", "_").replace(".", "_")

        # OneDriveのベースURL
        onedrive_base_url = global_config.base_url.replace(
            ".sharepoint.com", "-my.sharepoint.com"
        )
        api_base_url = f"{onedrive_base_url}/personal/{onedrive_user}"

        # サーバー相対パス
        server_relative_folder = f"/personal/{onedrive_user}{folder}"

        return api_base_url, server_relative_folder

    def _parse_full_url_path(self, folder_path: str) -> tuple[str, str]:
        """完全URLのパスを解析"""
        parsed_url = urlparse(folder_path)
        server_relative_folder = unquote(parsed_url.path)
        path_segments = server_relative_folder.split("/")

        # OneDriveかSharePointかを判定
        is_onedrive = len(path_segments) >= 2 and path_segments[1] == "personal"

        if is_onedrive:
            # OneDriveの場合
            onedrive_base_url = global_config.base_url.replace(
                ".sharepoint.com", "-my.sharepoint.com"
            )
            if len(path_segments) >= 3:
                personal_site_name = path_segments[2]
                api_base_url = f"{onedrive_base_url}/personal/{personal_site_name}"
            else:
                api_base_url = onedrive_base_url
        elif len(path_segments) >= 3 and path_segments[1] == "sites":
            # SharePointサイトの場合
            site_name = path_segments[2]
            api_base_url = f"{global_config.base_url}/sites/{site_name}"
        else:
            # それ以外はベースURLを使用
            api_base_url = global_config.base_url

        return api_base_url, server_relative_folder

    def _parse_site_path(self, folder_path: str) -> tuple[str, str]:
        """サイト指定のパスを解析"""
        # SiteName:/Folder/Path
        parts = folder_path.split(":", 1)
        if len(parts) != 2:
            raise ValueError(
                f"Invalid site path format: {folder_path}. "
                "Expected: SiteName:/Folder/Path"
            )

        site_name = parts[0]
        folder = parts[1]

        api_base_url = f"{global_config.base_url}/sites/{site_name}"
        server_relative_folder = f"/sites/{site_name}{folder}"

        return api_base_url, server_relative_folder

    def _build_full_url(self, api_base_url: str, server_relative_url: str) -> str:
        """APIベースURLとサーバー相対URLから完全URLを構築"""
        # api_base_urlからホスト部分を抽出
        parsed = urlparse(api_base_url)
        base = f"{parsed.scheme}://{parsed.netloc}"
        return f"{base}{server_relative_url}"
