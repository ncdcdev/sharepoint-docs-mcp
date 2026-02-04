"""
SharePoint Excel解析モジュール（ダウンロード+openpyxl方式）
"""

import json
import logging
from io import BytesIO
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell import Cell
from openpyxl.styles import Color

logger = logging.getLogger(__name__)


class SharePointExcelParser:
    """SharePoint Excelファイル解析クライアント"""

    def __init__(self, download_client):
        """
        Args:
            download_client: download_file(file_path) -> bytes メソッドを持つクライアント
        """
        self.download_client = download_client

    def search_cells(self, file_path: str, query: str) -> str:
        """
        セル内容を検索して該当位置を返す

        Args:
            file_path: Excelファイルのパス
            query: 検索キーワード

        Returns:
            JSON文字列（マッチしたセルの位置情報）
        """
        logger.info(f"Searching cells in Excel file: {file_path} (query={query})")

        try:
            # ファイルをダウンロード
            file_bytes = self.download_client.download_file(file_path)
            logger.info(f"Downloaded {len(file_bytes)} bytes")

            # BytesIOでメモリ上に展開
            file_stream = BytesIO(file_bytes)

            # openpyxlで読み込み
            workbook = load_workbook(file_stream, data_only=False, rich_text=True)

            matches = []
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                if sheet.dimensions:
                    for row in sheet.iter_rows():
                        for cell in row:
                            if cell.value is not None:
                                cell_value_str = str(cell.value)
                                if query in cell_value_str:
                                    matches.append({
                                        "sheet": sheet_name,
                                        "coordinate": cell.coordinate,
                                        "value": self._serialize_value(cell.value),
                                    })

            logger.info(f"Found {len(matches)} matches for query '{query}'")

            return json.dumps({
                "file_path": file_path,
                "mode": "search",
                "query": query,
                "match_count": len(matches),
                "matches": matches,
            }, ensure_ascii=False, indent=2)

        except Exception as e:
            logger.error(f"Failed to search cells in Excel file: {str(e)}")
            raise

    def parse_to_json(
        self,
        file_path: str,
        include_formatting: bool = False,
        sheet_name: str | None = None,
        cell_range: str | None = None,
    ) -> str:
        """
        Excelファイルを解析してJSON形式で返す

        Args:
            file_path: Excelファイルのパス
            include_formatting: 書式情報を含めるかどうか
                False (デフォルト): value, coordinate のみ
                True: value, coordinate, data_type, fill, merged, width, height を含む
            sheet_name: 特定シートのみ取得（Noneで全シート）
            cell_range: セル範囲指定（例: "A1:D10"）

        Returns:
            JSON文字列（全シート・全セルのデータ）
        """
        logger.info(
            f"Parsing Excel file: {file_path} "
            f"(include_formatting={include_formatting}, sheet={sheet_name}, range={cell_range})"
        )

        try:
            # ファイルをダウンロード
            file_bytes = self.download_client.download_file(file_path)
            logger.info(f"Downloaded {len(file_bytes)} bytes")

            # BytesIOでメモリ上に展開
            file_stream = BytesIO(file_bytes)

            # openpyxlで読み込み（data_only=Falseで数式も取得）
            workbook = load_workbook(file_stream, data_only=False, rich_text=True)

            # シートリストを取得
            if sheet_name:
                if sheet_name not in workbook.sheetnames:
                    raise ValueError(
                        f"Sheet '{sheet_name}' not found. "
                        f"Available sheets: {workbook.sheetnames}"
                    )
                sheets_to_parse = [sheet_name]
            else:
                sheets_to_parse = workbook.sheetnames

            # シートを解析
            result = {"file_path": file_path, "sheets": []}

            for name in sheets_to_parse:
                sheet = workbook[name]
                sheet_data = self._parse_sheet(sheet, include_formatting, cell_range)
                result["sheets"].append(sheet_data)

            logger.info(f"Parsed {len(result['sheets'])} sheets")
            return json.dumps(result, ensure_ascii=False, indent=2)

        except Exception as e:
            logger.error(f"Failed to parse Excel file: {str(e)}")
            raise

    def _parse_sheet(
        self,
        sheet,
        include_formatting: bool,
        cell_range: str | None = None,
    ) -> dict[str, Any]:
        """
        シートを解析してdict形式で返す

        Args:
            sheet: openpyxl Worksheet
            include_formatting: 書式情報を含めるかどうか
            cell_range: セル範囲指定（例: "A1:D10"）

        Returns:
            シートデータのdict
        """
        sheet_data = {
            "name": sheet.title,
            "dimensions": str(sheet.dimensions) if sheet.dimensions else None,
            "rows": [],
        }

        # セル範囲を取得
        if cell_range:
            # 指定された範囲のみを取得
            sheet_data["requested_range"] = cell_range
            range_data = sheet[cell_range]

            # 統一的にタプルのタプル形式に変換
            if isinstance(range_data, Cell):
                # 単一セルの場合
                rows_to_process = ((range_data,),)
            elif range_data and not isinstance(range_data[0], tuple):
                # 単一列/行の場合
                rows_to_process = (range_data,)
            else:
                # 通常の範囲の場合
                rows_to_process = range_data

            for row in rows_to_process:
                row_data = [self._parse_cell(cell, include_formatting) for cell in row]
                sheet_data["rows"].append(row_data)
        elif sheet.dimensions:
            for row in sheet.iter_rows():
                row_data = []
                for cell in row:
                    cell_data = self._parse_cell(cell, include_formatting)
                    row_data.append(cell_data)
                sheet_data["rows"].append(row_data)

        return sheet_data

    def _parse_cell(self, cell, include_formatting: bool) -> dict[str, Any]:
        """
        セルを解析してdict形式で返す

        Args:
            cell: openpyxl Cell
            include_formatting: 書式情報を含めるかどうか

        Returns:
            セルデータのdict
        """
        # 基本情報（常に含む）
        cell_data = {
            "value": self._serialize_value(cell.value),
            "coordinate": cell.coordinate,
        }

        # 書式情報（オプション）
        if include_formatting:
            cell_data["data_type"] = cell.data_type

            # 塗りつぶし色情報
            if cell.fill:
                cell_data["fill"] = {
                    "pattern_type": cell.fill.patternType,
                    "fg_color": self._color_to_hex(cell.fill.fgColor),
                    "bg_color": self._color_to_hex(cell.fill.bgColor),
                }

            # セル結合情報
            if hasattr(cell, "parent") and cell.parent:
                sheet = cell.parent
                for merged_range in sheet.merged_cells.ranges:
                    if cell.coordinate in merged_range:
                        cell_data["merged"] = {
                            "range": str(merged_range),
                            "is_top_left": cell.coordinate
                            == merged_range.start_cell.coordinate,
                        }
                        break

            # セルサイズ情報（MergedCellはcolumn_letterを持たない可能性がある）
            if hasattr(cell, "column_letter") and hasattr(cell, "row"):
                if cell.column_letter and cell.row:
                    sheet = cell.parent
                    # 列幅
                    if cell.column_letter in sheet.column_dimensions:
                        col_dim = sheet.column_dimensions[cell.column_letter]
                        cell_data["width"] = col_dim.width
                    # 行高
                    if cell.row in sheet.row_dimensions:
                        row_dim = sheet.row_dimensions[cell.row]
                        cell_data["height"] = row_dim.height

        return cell_data

    def _serialize_value(self, value: Any) -> Any:
        """
        セル値をJSONシリアライズ可能な形式に変換

        Args:
            value: セル値

        Returns:
            JSONシリアライズ可能な値
        """
        if value is None:
            return None

        # 基本的な型（JSONシリアライズ可能）はそのまま
        if isinstance(value, (str, int, float, bool)):
            return value

        # その他の型（datetime, timedelta等）は文字列に変換
        return str(value)

    def _color_to_hex(self, color: Color | None) -> str | None:
        """
        openpyxl Colorオブジェクトを16進数カラーコードに変換

        Args:
            color: openpyxl Color

        Returns:
            16進数カラーコード (例: "#FF0000") またはNone
        """
        if color is None:
            return None

        if color.type == "rgb":
            # RGB形式 (例: "FFFF0000" → "#FF0000")
            rgb = color.rgb
            if rgb and isinstance(rgb, str) and len(rgb) >= 6:
                return f"#{rgb[-6:]}"

        elif color.type == "theme":
            # テーマカラーは複雑なので、簡易的に処理
            return f"theme_{color.theme}"

        return None
