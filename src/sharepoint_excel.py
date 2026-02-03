"""
SharePoint Excel操作モジュール
"""

import base64
import logging
from typing import Protocol
from urllib.parse import quote, urlparse

import requests

from src.error_messages import handle_sharepoint_error

logger = logging.getLogger(__name__)


class AuthClient(Protocol):
    """認証クライアントのプロトコル（証明書認証/OAuth両対応）"""

    def get_access_token(self) -> str:
        """アクセストークンを取得"""
        ...


class SharePointExcelClient:
    """SharePoint Excel操作クライアント"""

    def __init__(self, site_url: str, auth: AuthClient):
        self.site_url = site_url.rstrip("/")
        self.auth = auth

    def list_sheets(self, file_path: str) -> str:
        """
        Excelファイルのシート一覧を取得

        Args:
            file_path: Excelファイルのパス（検索結果から取得）

        Returns:
            XML形式のシート一覧
        """
        logger.info(f"Listing sheets for: {file_path}")

        try:
            excel_rest_url = self._build_excel_rest_url(
                file_path, "Sheets", format_type="atom"
            )
            access_token = self.auth.get_access_token()

            headers = {
                "Authorization": f"Bearer {access_token}",
                "Accept": "application/atom+xml",
            }

            response = requests.get(excel_rest_url, headers=headers, timeout=30)
            response.raise_for_status()

            logger.info("Successfully retrieved sheet list")
            return response.text

        except Exception as e:
            logger.error(f"Failed to list sheets: {str(e)}")
            raise handle_sharepoint_error(
                e,
                "excel_list_sheets",
                excel_context={"file_path": file_path, "sheet_name": None, "range_spec": None},
            ) from e

    def get_sheet_image(self, file_path: str, sheet_name: str) -> str:
        """
        シートのキャプチャ画像を取得

        Args:
            file_path: Excelファイルのパス
            sheet_name: シート名

        Returns:
            base64エンコードされた画像データ
        """
        logger.info(f"Getting image for sheet '{sheet_name}' in {file_path}")

        try:
            # シート名のシングルクォートをエスケープ
            escaped_sheet_name = sheet_name.replace("'", "''")
            resource = f"Sheets('{escaped_sheet_name}')"

            excel_rest_url = self._build_excel_rest_url(
                file_path, resource, format_type="image"
            )
            access_token = self.auth.get_access_token()

            headers = {
                "Authorization": f"Bearer {access_token}",
                "Accept": "image/png",
            }

            response = requests.get(excel_rest_url, headers=headers, timeout=30)
            response.raise_for_status()

            # バイナリデータをbase64エンコード
            image_base64 = base64.b64encode(response.content).decode("utf-8")
            logger.info("Successfully retrieved sheet image")
            return image_base64

        except Exception as e:
            logger.error(f"Failed to get sheet image: {str(e)}")
            raise handle_sharepoint_error(
                e,
                "excel_get_image",
                excel_context={"file_path": file_path, "sheet_name": sheet_name, "range_spec": None},
            ) from e

    def get_range_data(self, file_path: str, range_spec: str) -> str:
        """
        セル範囲のデータを取得

        Args:
            file_path: Excelファイルのパス
            range_spec: セル範囲（例: "Sheet1!A1:C10"）

        Returns:
            XML形式のセルデータ
        """
        logger.info(f"Getting range data '{range_spec}' from {file_path}")

        try:
            # 範囲指定のシングルクォートをエスケープ
            escaped_range = range_spec.replace("'", "''")
            resource = f"Ranges('{escaped_range}')"

            excel_rest_url = self._build_excel_rest_url(
                file_path, resource, format_type="atom"
            )
            access_token = self.auth.get_access_token()

            headers = {
                "Authorization": f"Bearer {access_token}",
                "Accept": "application/atom+xml",
            }

            response = requests.get(excel_rest_url, headers=headers, timeout=30)
            response.raise_for_status()

            logger.info("Successfully retrieved range data")
            return response.text

        except Exception as e:
            logger.error(f"Failed to get range data: {str(e)}")
            raise handle_sharepoint_error(
                e,
                "excel_get_range",
                excel_context={"file_path": file_path, "sheet_name": None, "range_spec": range_spec},
            ) from e

    def _build_excel_rest_url(
        self, file_path: str, resource: str, format_type: str
    ) -> str:
        """
        Excel REST API URLを構築

        Args:
            file_path: Excelファイルのパス
            resource: リソース（Sheets, Ranges, etc.）
            format_type: フォーマット（atom, image）

        Returns:
            完全なExcel REST API URL
        """
        # ファイルパスからサイトURLとライブラリパスを抽出
        parsed_url = urlparse(file_path)
        path_segments = parsed_url.path.split("/")

        # サイト名を検出
        site_name = None
        library_path = None

        for i, segment in enumerate(path_segments):
            if segment == "sites" and i + 1 < len(path_segments):
                site_name = path_segments[i + 1]
                # サイト名以降のパスをライブラリパスとする
                library_path = "/".join(path_segments[i + 2:])
                break

        if not site_name or not library_path:
            raise ValueError(f"Invalid file path format: {file_path}")

        # サイトURLを構築
        base_url = f"{parsed_url.scheme}://{parsed_url.netloc}"
        site_url = f"{base_url}/sites/{site_name}"

        # ライブラリパスをURLエンコード
        encoded_library_path = quote(library_path, safe="/")

        # Excel REST API URLを構築
        excel_rest_url = (
            f"{site_url}/_vti_bin/ExcelRest.aspx/"
            f"{encoded_library_path}/Model/{resource}?$format={format_type}"
        )

        logger.debug(f"Built Excel REST URL: {excel_rest_url}")
        return excel_rest_url
