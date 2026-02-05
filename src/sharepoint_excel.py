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
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string

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
                                    matches.append(
                                        {
                                            "sheet": sheet_name,
                                            "coordinate": cell.coordinate,
                                            "value": self._serialize_value(cell.value),
                                        }
                                    )

            logger.info(f"Found {len(matches)} matches for query '{query}'")

            return json.dumps(
                {
                    "file_path": file_path,
                    "mode": "search",
                    "query": query,
                    "match_count": len(matches),
                    "matches": matches,
                },
                ensure_ascii=False,
                indent=2,
            )

        except Exception as e:
            logger.error(f"Failed to search cells in Excel file: {str(e)}")
            raise

    def parse_to_json(
        self,
        file_path: str,
        include_formatting: bool = False,
        sheet_name: str | None = None,
        cell_range: str | None = None,
        include_header: bool = True,
        metadata_only: bool = False,
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
            include_header: ヘッダー情報を自動追加して返すかどうか
                True (デフォルト): freeze_panesで固定された行をヘッダーとして認識し、
                     cell_range指定時にヘッダーが範囲外でも自動的に追加して
                     header_rows と data_rows に分けて返す
                False: rows にすべてのデータを含む（ヘッダー自動追加なし）
            metadata_only: メタデータのみを返すかどうか
                True: data_rows を空リストにする（header_rows とメタデータのみ返す）
                False (デフォルト): すべてのデータを含める

        Returns:
            JSON文字列（全シート・全セルのデータ）
        """
        logger.info(
            f"Parsing Excel file: {file_path} "
            f"(include_formatting={include_formatting}, sheet={sheet_name}, range={cell_range}, "
            f"metadata_only={metadata_only})"
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
                sheet_data = self._parse_sheet(
                    sheet, include_formatting, cell_range, include_header, metadata_only
                )
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
        include_header: bool = False,
        metadata_only: bool = False,
    ) -> dict[str, Any]:
        """
        シートを解析してdict形式で返す

        Args:
            sheet: openpyxl Worksheet
            include_formatting: 書式情報を含めるかどうか
            cell_range: セル範囲指定（例: "A1:D10"）
            include_header: ヘッダー情報を分離して返すかどうか
            metadata_only: メタデータのみを返すかどうか

        Returns:
            シートデータのdict
        """
        sheet_data = {
            "name": sheet.title,
            "dimensions": str(sheet.dimensions) if sheet.dimensions else None,
        }

        # freeze_panes情報の取得
        frozen_rows = 0
        frozen_cols = 0
        if include_header:
            frozen_rows, frozen_cols = self._parse_freeze_panes(sheet.freeze_panes)
            if sheet.freeze_panes:
                sheet_data["freeze_panes"] = sheet.freeze_panes
            sheet_data["frozen_rows"] = frozen_rows
            sheet_data["frozen_cols"] = frozen_cols

        # マージセル情報をキャッシュ（パフォーマンス最適化）
        merged_cell_map: dict[str, str] | None = None
        if include_formatting and sheet.merged_cells.ranges:
            merged_cell_map = {}
            for merged_range in sheet.merged_cells.ranges:
                for cell_coord in merged_range.cells:
                    # cell_coord is (row, col) tuple
                    col_letter = get_column_letter(cell_coord[1])
                    coord_str = f"{col_letter}{cell_coord[0]}"
                    merged_cell_map[coord_str] = str(merged_range)

        # セル範囲の拡張（include_headerがTrueで固定行がある場合）
        all_rows = []
        if cell_range:
            sheet_data["requested_range"] = cell_range

            # セル範囲を拡張してヘッダーを含める
            if include_header and frozen_rows > 0:
                header_range, data_range = self._expand_range_with_headers(
                    cell_range, frozen_rows, frozen_cols
                )

                # ヘッダー範囲がある場合は取得
                if header_range:
                    header_data = sheet[header_range]
                    header_rows = self._normalize_range_data(header_data)
                    for row in header_rows:
                        row_data = [
                            self._parse_cell(cell, include_formatting, merged_cell_map)
                            for cell in row
                        ]
                        all_rows.append(row_data)

                # データ範囲を取得
                range_data = sheet[data_range]
                data_rows = self._normalize_range_data(range_data)
                for row in data_rows:
                    row_data = [
                        self._parse_cell(cell, include_formatting, merged_cell_map)
                        for cell in row
                    ]
                    all_rows.append(row_data)
            else:
                # 通常のセル範囲取得
                range_data = sheet[cell_range]
                rows_to_process = self._normalize_range_data(range_data)
                for row in rows_to_process:
                    row_data = [
                        self._parse_cell(cell, include_formatting, merged_cell_map)
                        for cell in row
                    ]
                    all_rows.append(row_data)
        elif sheet.dimensions:
            # シート全体を取得
            for row in sheet.iter_rows():
                row_data = []
                for cell in row:
                    cell_data = self._parse_cell(
                        cell, include_formatting, merged_cell_map
                    )
                    row_data.append(cell_data)
                all_rows.append(row_data)

        # レスポンス形式の分岐
        if include_header:
            header_rows, data_rows = self._split_rows_by_header(all_rows, frozen_rows)
            sheet_data["header_rows"] = header_rows
            # metadata_onlyの場合はdata_rowsを空リストにする
            sheet_data["data_rows"] = [] if metadata_only else data_rows
        else:
            # metadata_onlyの場合はrowsを空リストにする
            sheet_data["rows"] = [] if metadata_only else all_rows

        return sheet_data

    def _parse_cell(
        self,
        cell,
        include_formatting: bool,
        merged_cell_map: dict[str, str] | None = None,
    ) -> dict[str, Any]:
        """
        セルを解析してdict形式で返す

        Args:
            cell: openpyxl Cell
            include_formatting: 書式情報を含めるかどうか
            merged_cell_map: マージセル座標からマージ範囲へのマップ（パフォーマンス最適化用）

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

            # セル結合情報（キャッシュを使用してO(1)で検索）
            if merged_cell_map and cell.coordinate in merged_cell_map:
                merged_range_str = merged_cell_map[cell.coordinate]
                # 左上セルかどうかを判定（マージ範囲の最初の座標と比較）
                range_start = merged_range_str.split(":")[0]
                cell_data["merged"] = {
                    "range": merged_range_str,
                    "is_top_left": cell.coordinate == range_start,
                }

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

    def _parse_freeze_panes(self, freeze_panes: str | None) -> tuple[int, int]:
        """
        freeze_panes文字列を解析して固定行数・列数を返す

        Args:
            freeze_panes: freeze_panes文字列（例: "B2", "A2", "B1", None）

        Returns:
            (frozen_rows, frozen_cols)のタプル
            例: "B2" → (1, 1)（行1と列Aが固定）
                "A2" → (1, 0)（行1のみ固定）
                "B1" → (0, 1)（列Aのみ固定）
                None → (0, 0)（固定なし）
        """
        if not freeze_panes:
            return (0, 0)

        try:
            # "B2" → ("B", 2)
            col_letter, row = coordinate_from_string(freeze_panes)
            # "B" → 2
            col_index = column_index_from_string(col_letter)

            # freeze_panes="B2"の場合、行2より前（行1）と列B（2列目）より前（列A）が固定
            frozen_rows = row - 1
            frozen_cols = col_index - 1

            return (frozen_rows, frozen_cols)
        except Exception as e:
            logger.warning(f"Failed to parse freeze_panes '{freeze_panes}': {e}")
            return (0, 0)

    def _expand_range_with_headers(
        self, cell_range: str, frozen_rows: int, frozen_cols: int
    ) -> tuple[str | None, str]:
        """
        cell_rangeを固定範囲を含むように拡張

        Args:
            cell_range: セル範囲（例: "A5:D10"）
            frozen_rows: 固定行数
            frozen_cols: 固定列数

        Returns:
            (header_range, data_range)のタプル
            header_range: ヘッダー範囲（固定行がない場合はNone）
            data_range: データ範囲（元のcell_range）
        """
        if frozen_rows == 0:
            return (None, cell_range)

        try:
            # セル範囲を解析
            if ":" in cell_range:
                start_cell, end_cell = cell_range.split(":")
            else:
                # 単一セルの場合
                start_cell = cell_range
                end_cell = cell_range

            start_col, start_row = coordinate_from_string(start_cell)
            end_col, _ = coordinate_from_string(end_cell)

            # 開始行が固定範囲内の場合は拡張不要
            if start_row <= frozen_rows:
                return (None, cell_range)

            # ヘッダー範囲を計算（行1からfrozen_rowsまで、列は元の範囲と同じ）
            header_range = f"{start_col}1:{end_col}{frozen_rows}"

            return (header_range, cell_range)
        except Exception as e:
            logger.warning(
                f"Failed to expand range '{cell_range}' with frozen_rows={frozen_rows}: {e}"
            )
            return (None, cell_range)

    def _split_rows_by_header(
        self, rows: list[list[dict[str, Any]]], frozen_rows: int
    ) -> tuple[list[list[dict[str, Any]]], list[list[dict[str, Any]]]]:
        """
        取得した行データをヘッダー行とデータ行に分割

        Args:
            rows: 行データのリスト
            frozen_rows: 固定行数

        Returns:
            (header_rows, data_rows)のタプル
        """
        if frozen_rows == 0:
            return ([], rows)

        if len(rows) <= frozen_rows:
            # すべてヘッダー
            return (rows, [])

        header_rows = rows[:frozen_rows]
        data_rows = rows[frozen_rows:]

        return (header_rows, data_rows)

    def _normalize_range_data(self, range_data):
        """
        openpyxlの範囲データを統一的なタプルのタプル形式に変換

        Args:
            range_data: sheet[range]の戻り値

        Returns:
            タプルのタプル形式の範囲データ
        """
        if isinstance(range_data, Cell):
            # 単一セルの場合
            return ((range_data,),)
        elif range_data and not isinstance(range_data[0], tuple):
            # 単一列/行の場合
            return (range_data,)
        else:
            # 通常の範囲の場合
            return range_data
