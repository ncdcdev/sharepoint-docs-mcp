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

from src.config import config

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
                現状は指定しても返却内容は変わらない。
                - value / coordinate は常に返す
                - 結合セルがある場合は merged を追加し、merged_ranges を返す
                ※ data_type / fill / width / height は返さない
            sheet_name: 特定シートのみ取得（Noneで全シート）
            cell_range: セル範囲指定（例: "A1:D10"）
            include_header: ヘッダー情報を自動追加して返すかどうか
                True (デフォルト): freeze_panesで固定された行をヘッダーとして認識し、
                     cell_range指定時にヘッダーが範囲外でも自動的に追加して
                     header_rows と data_rows に分けて返す
                False: rows にすべてのデータを含める（ヘッダー自動追加なし）
            metadata_only: メタデータのみを返すかどうか
                True: data_rows を空リストにする（header_rows とメタデータのみ返す）
                False (デフォルト): すべてのデータを含める

        Returns:
            JSON文字列（全シート・全セルのデータ）
        """
        logger.info(
            f"Parsing Excel file: {file_path} "
            f"(include_formatting={include_formatting}, sheet={sheet_name}, range={cell_range}, "
            f"include_header={include_header}, metadata_only={metadata_only})"
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
        include_header: bool = True,
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

        # freeze_panes情報の取得と検証
        frozen_rows = 0
        frozen_cols = 0
        if include_header:
            frozen_rows, frozen_cols = self._get_frozen_panes(sheet)

            # frozen_rows検証（DoS対策）
            if frozen_rows > config.excel_max_frozen_rows:
                logging.warning(
                    "ヘッダー行数が上限を超えたため、freeze_panesを無視して処理します",
                    config.excel_max_frozen_rows,
                    frozen_rows,
                )
                frozen_rows = 0
                frozen_cols = 0

            if frozen_rows > 0 or frozen_cols > 0:
                sheet_data["freeze_panes"] = self._format_freeze_panes(
                    frozen_rows, frozen_cols
                )
            sheet_data["frozen_rows"] = frozen_rows
            sheet_data["frozen_cols"] = frozen_cols

        # マージセル情報をキャッシュ（パフォーマンス最適化）
        # 結合セルの場合はmerged_rangesを渡す
        merged_cell_map: dict[str, str] | None = None
        merged_anchor_value_map: dict[str, Any] | None = None
        merged_ranges: list[dict[str, Any]] = []

        if sheet.merged_cells.ranges:
            merged_cell_map = {}
            merged_anchor_value_map = {}

            for merged_range in sheet.merged_cells.ranges:
                merged_range_str = str(merged_range)
                range_start = merged_range_str.split(":")[0]

                # アンカー値を決定（左上が空なら結合範囲内を走査）
                anchor_coord = range_start
                anchor_value = self._serialize_value(sheet[range_start].value)

                # セル座標 -> 結合範囲 のマップ（O(1)参照用）
                # ここは「面積」を全部展開するが、通常のExcelの結合範囲は現実的なサイズ。
                #       もし極端な結合が来たら、別途 input 制限（sheet.dimensions の上限）で守る。
                for cell_coord in merged_range.cells:
                    # cell_coord is (row, col) tuple
                    row_idx, col_idx = cell_coord
                    col_letter = get_column_letter(col_idx)
                    coord_str = f"{col_letter}{row_idx}"
                    merged_cell_map[coord_str] = merged_range_str

                    # 左上が空の場合は結合範囲内で最初の非空セルを探す
                    if anchor_value is None:
                        cell_value = self._serialize_value(
                            sheet.cell(row=row_idx, column=col_idx).value
                        )
                        if cell_value is not None:
                            anchor_value = cell_value
                            anchor_coord = coord_str

                # アンカー値を保存（結合セルの値埋め用）
                merged_anchor_value_map[merged_range_str] = anchor_value

                # 結合範囲そのものを返す（結合セルがある時だけ返す）
                merged_ranges.append(
                    {
                        "range": merged_range_str,
                        "anchor": {"coordinate": anchor_coord, "value": anchor_value},
                    }
                )

            # ここは「結合セルがある時だけ」返す
            if merged_ranges:
                sheet_data["merged_ranges"] = merged_ranges

        # セル範囲の拡張とデータサイズ検証
        all_rows = []
        if cell_range:
            sheet_data["requested_range"] = cell_range

            # データサイズ検証（DoS対策）
            range_rows, range_cols = self._calculate_range_size(cell_range)
            if (
                range_rows > config.excel_max_data_rows
                or range_cols > config.excel_max_data_cols
            ):
                raise ValueError(
                    f"データサイズ({range_rows}行 × {range_cols}列)が上限"
                    f"({config.excel_max_data_rows}行 × {config.excel_max_data_cols}列)を超えています。"
                    f"cell_rangeパラメータで必要な範囲を指定してください。"
                    f"例: cell_range='A1:Z1000'"
                )

            # セル範囲を拡張してヘッダーを含める
            if include_header and frozen_rows > 0:
                header_range, data_range = self._expand_range_with_headers(
                    cell_range, frozen_rows
                )

                # ヘッダー範囲がある場合は取得
                if header_range:
                    header_data = sheet[header_range]
                    header_rows_data = self._normalize_range_data(header_data)
                    all_rows.extend(
                        self._parse_rows(
                            header_rows_data,
                            include_formatting,
                            merged_cell_map,
                            merged_anchor_value_map,
                        )
                    )

                # データ範囲を取得（metadata_onlyの場合はスキップ）
                if not metadata_only:
                    range_data = sheet[data_range]
                    data_rows_data = self._normalize_range_data(range_data)
                    all_rows.extend(
                        self._parse_rows(
                            data_rows_data,
                            include_formatting,
                            merged_cell_map,
                            merged_anchor_value_map,
                        )
                    )
            else:
                # 通常のセル範囲取得（metadata_onlyの場合もヘッダーなしなので取得）
                if not metadata_only:
                    range_data = sheet[cell_range]
                    rows_to_process = self._normalize_range_data(range_data)
                    all_rows.extend(
                        self._parse_rows(
                            rows_to_process,
                            include_formatting,
                            merged_cell_map,
                            merged_anchor_value_map,
                        )
                    )
        elif sheet.dimensions:
            # シート全体を取得
            # データサイズ検証（DoS対策）
            sheet_rows, sheet_cols = self._calculate_range_size(sheet.dimensions)
            if (
                sheet_rows > config.excel_max_data_rows
                or sheet_cols > config.excel_max_data_cols
            ):
                raise ValueError(
                    f"シート全体のサイズ({sheet_rows}行 × {sheet_cols}列)が上限"
                    f"({config.excel_max_data_rows}行 × {config.excel_max_data_cols}列)を超えています。"
                    f"cell_rangeパラメータで必要な範囲を指定してください。"
                    f"例: cell_range='A1:Z1000'"
                )

            # metadata_onlyの場合はヘッダーのみ取得
            rows_to_process = None
            if metadata_only and include_header and frozen_rows > 0:
                # ヘッダー行のみ取得
                rows_to_process = tuple(sheet.iter_rows(max_row=frozen_rows))
            elif not metadata_only:
                # 全データを取得
                rows_to_process = tuple(sheet.iter_rows())

            if rows_to_process:
                all_rows.extend(
                    self._parse_rows(
                        rows_to_process,
                        include_formatting,
                        merged_cell_map,
                        merged_anchor_value_map,
                    )
                )

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
        merged_anchor_value_map: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """
        セルを解析してdict形式で返す

        Args:
            cell: openpyxl Cell
            include_formatting: 書式情報を含めるかどうか
            merged_cell_map: マージセル座標からマージ範囲へのマップ（パフォーマンス最適化用）
            merged_anchor_value_map: マージ範囲 -> アンカー値 のマップ（結合セルの値埋め用）

        Returns:
            セルデータのdict
        """
        # 基本情報（常に含む）
        cell_data = {
            "value": self._serialize_value(cell.value),
            "coordinate": cell.coordinate,
        }

        # セル結合情報（構造理解に必要）
        if merged_cell_map and cell.coordinate in merged_cell_map:
            merged_range_str = merged_cell_map[cell.coordinate]
            range_start = merged_range_str.split(":")[0]
            cell_data["merged"] = {
                "range": merged_range_str,
                "is_top_left": cell.coordinate == range_start,
            }

            # 結合セル内の空セルにも value を埋める（propagate）
            if cell_data["value"] is None and merged_anchor_value_map:
                anchor_value = merged_anchor_value_map.get(merged_range_str)
                if anchor_value is not None:
                    cell_data["value"] = anchor_value

        # 書式情報（オプション）
        # 現運用では fill/width/height/data_type は返さない。
        #       include_formatting を指定しても返却内容は変わらない。
        return cell_data

    def _parse_rows(
        self,
        rows: tuple[tuple[Cell, ...], ...],
        include_formatting: bool,
        merged_cell_map: dict[str, str] | None = None,
        merged_anchor_value_map: dict[str, Any] | None = None,
    ) -> list[list[dict[str, Any]]]:
        """
        行データを解析してリスト形式で返す（コード重複削減用ヘルパー）

        Args:
            rows: 行データのタプル
            include_formatting: 書式情報を含めるか
            merged_cell_map: マージセル情報
            merged_anchor_value_map: マージ範囲 -> アンカー値

        Returns:
            解析された行データのリスト
        """
        parsed_rows = []
        for row in rows:
            row_data = [
                self._parse_cell(
                    cell,
                    include_formatting,
                    merged_cell_map,
                    merged_anchor_value_map,
                )
                for cell in row
            ]
            parsed_rows.append(row_data)
        return parsed_rows

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

    def _calculate_range_size(self, range_str: str) -> tuple[int, int]:
        """
        セル範囲文字列から行数と列数を計算

        Args:
            range_str: セル範囲（例: "A1:D10" または "A1:XFD1048576"）

        Returns:
            (rows, cols)のタプル
        """
        try:
            if ":" in range_str:
                start_cell, end_cell = range_str.split(":")
            else:
                # 単一セルの場合
                return (1, 1)

            start_col, start_row = coordinate_from_string(start_cell)
            end_col, end_row = coordinate_from_string(end_cell)

            start_col_idx = column_index_from_string(start_col)
            end_col_idx = column_index_from_string(end_col)

            # 逆順序の範囲を検出（セキュリティ対策）
            if end_row < start_row or end_col_idx < start_col_idx:
                raise ValueError(
                    f"無効なセル範囲: '{range_str}'。"
                    f"範囲は正しい順序で指定してください（例: 'A1:Z100'）"
                )

            rows = end_row - start_row + 1
            cols = end_col_idx - start_col_idx + 1

            return (rows, cols)
        except Exception as e:
            logger.warning(f"Failed to calculate range size '{range_str}': {e}")
            return (0, 0)

    def _get_frozen_panes(self, sheet) -> tuple[int, int]:
        """
        シートのpane情報から固定行数・列数を返す（ySplit/xSplit使用）

        sheet.freeze_panes（= pane.topLeftCell）はスクロール位置に依存するため、
        正確な固定行数・列数を得るには pane.ySplit / pane.xSplit を直接参照する。

        Args:
            sheet: openpyxl Worksheet

        Returns:
            (frozen_rows, frozen_cols)のタプル
        """
        try:
            pane = sheet.sheet_view.pane
            if pane is None:
                return (0, 0)
            if pane.state not in ("frozen", "frozenSplit"):
                return (0, 0)
            frozen_rows = int(pane.ySplit) if pane.ySplit else 0
            frozen_cols = int(pane.xSplit) if pane.xSplit else 0
            return (frozen_rows, frozen_cols)
        except Exception as e:
            logger.warning(f"Failed to get frozen panes info: {e}")
            return (0, 0)

    def _format_freeze_panes(self, frozen_rows: int, frozen_cols: int) -> str:
        """
        固定行数・列数からfreeze_panes文字列表現を生成

        Args:
            frozen_rows: 固定行数
            frozen_cols: 固定列数

        Returns:
            freeze_panes文字列表現（例: "B4"）
        """
        col_letter = get_column_letter(frozen_cols + 1)
        return f"{col_letter}{frozen_rows + 1}"

    def _expand_range_with_headers(
        self, cell_range: str, frozen_rows: int
    ) -> tuple[str | None, str]:
        """
        cell_rangeを固定範囲を含むように拡張

        Args:
            cell_range: セル範囲（例: "A5:D10"）
            frozen_rows: 固定行数

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

    def _normalize_range_data(self, range_data: Any) -> tuple[tuple[Cell, ...], ...]:
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
        elif not range_data:
            # 空の場合
            return ()
        elif not isinstance(range_data[0], tuple):
            # 単一列/行の場合
            return (range_data,)
        else:
            # 通常の範囲の場合
            return range_data
