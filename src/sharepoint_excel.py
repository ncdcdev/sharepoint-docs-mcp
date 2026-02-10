"""
SharePoint Excel解析モジュール（ダウンロード+openpyxl方式）
"""

import difflib
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

    def search_cells(
        self,
        file_path: str,
        query: str,
        sheet_name: str | None = None,
    ) -> str:
        """
        セル内容を検索して該当位置を返す

        Args:
            file_path: Excelファイルのパス
            query: 検索キーワード
            sheet_name: 検索対象シート名（指定時はまずそのシートを検索し、マッチ0件なら全シート検索にフォールバック）

        Returns:
            JSON文字列（マッチしたセルの位置情報）
        """
        logger.info(
            f"Searching cells in Excel file: {file_path} (query={query}, sheet={sheet_name})"
        )

        try:
            # ファイルをダウンロード
            file_bytes = self.download_client.download_file(file_path)
            logger.info(f"Downloaded {len(file_bytes)} bytes")

            # BytesIOでメモリ上に展開
            file_stream = BytesIO(file_bytes)

            # openpyxlで読み込み
            workbook = load_workbook(file_stream, data_only=False, rich_text=True)

            matches = []

            # sheet_name 指定がある場合はそのシートを優先して検索
            if sheet_name:
                if sheet_name in workbook.sheetnames:
                    self._scan_sheet(workbook[sheet_name], sheet_name, query, matches)

                    # マッチが無ければ全シート走査にフォールバック
                    if len(matches) == 0:
                        for sn in workbook.sheetnames:
                            if sn == sheet_name:
                                continue
                            self._scan_sheet(workbook[sn], sn, query, matches)
                else:
                    # sheet_name が存在しない場合は「指定なし」と同じ扱いで全シート検索
                    for sn in workbook.sheetnames:
                        self._scan_sheet(workbook[sn], sn, query, matches)
            else:
                # 全シート検索
                for sn in workbook.sheetnames:
                    self._scan_sheet(workbook[sn], sn, query, matches)

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
        sheet_name: str | None = None,
        cell_range: str | None = None,
        include_frozen_rows: bool = True,
        include_cell_styles: bool = False,
    ) -> str:
        """
        Excelファイルを解析してJSON形式で返す

        Args:
            file_path: Excelファイルのパス
            sheet_name: 特定シートのみ取得（Noneで全シート）
            cell_range: セル範囲指定（例: "A1:D10"）
            include_frozen_rows: cell_range指定時に固定行（ヘッダー）を自動追加
                True（デフォルト）: frozen_rowsで指定された行を自動的に取得
                False: 指定されたcell_rangeのみを取得
            include_cell_styles: セルの色・サイズ情報（default: false）
                色分けデータ抽出時のみ使用。トークン消費+約20%

        Returns:
            JSON文字列
            - 各セルのデータ: value（値）、coordinate（座標）
            - 構造情報: シート名、dimensions（シート全体のセル範囲、例: "A1:D10"）
            - 構造情報: frozen_rows（固定行数）、frozen_cols（固定列数）
            - 条件付き構造情報: freeze_panes（存在する場合）、merged_ranges（結合セルが存在する場合）
            - スタイル情報（include_cell_styles=Trueの場合）: fill（背景色）、width（列幅）、height（行高さ）
        """
        logger.info(
            f"Parsing Excel file: {file_path} (sheet={sheet_name}, range={cell_range})"
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
            sheet_resolution: dict[str, Any] | None = None

            if sheet_name:
                resolved, candidates = self._resolve_sheet_name(
                    workbook.sheetnames, sheet_name
                )

                if resolved:
                    sheets_to_parse = [resolved]
                    if resolved != sheet_name:
                        sheet_resolution = {
                            "status": "resolved",
                            "requested": sheet_name,
                            "resolved": resolved,
                        }
                else:
                    # sheet_name が解決できない場合
                    # cell_range が指定されていれば、範囲が限定されるので全シートにフォールバック（やりすぎを避けつつ取りこぼし防止）
                    if cell_range:
                        sheets_to_parse = workbook.sheetnames
                        sheet_resolution = {
                            "status": "fallback_all_sheets",
                            "requested": sheet_name,
                            "resolved": None,
                            "candidates": candidates,
                            "reason": "sheet not found; fallback to all sheets because cell_range is specified",
                        }
                    else:
                        # ここで例外にせず、空の sheets を返して候補を出す（LLMが次の手を打てる）
                        sheets_to_parse = []
                        sheet_resolution = {
                            "status": "not_found",
                            "requested": sheet_name,
                            "resolved": None,
                            "candidates": candidates,
                        }
            else:
                sheets_to_parse = workbook.sheetnames

            # シートを解析
            result = {
                "file_path": file_path,
                "sheets": [],
            }

            # nullでない場合のみ追加
            if sheet_name is not None:
                result["requested_sheet"] = sheet_name
            if cell_range is not None:
                result["requested_range"] = cell_range

            if sheet_resolution:
                result["sheet_resolution"] = sheet_resolution
                result["available_sheets"] = workbook.sheetnames
                if sheet_resolution.get("status") != "resolved":
                    result["warning"] = (
                        "requested sheet_name was not found or ambiguous"
                    )

            for name in sheets_to_parse:
                sheet = workbook[name]
                sheet_data = self._parse_sheet(
                    sheet,
                    cell_range,
                    include_frozen_rows,
                    include_cell_styles,
                )
                result["sheets"].append(sheet_data)

            logger.info(f"Parsed {len(result['sheets'])} sheets")
            return json.dumps(result, ensure_ascii=False, indent=2)

        except Exception as e:
            logger.error(f"Failed to parse Excel file: {str(e)}")
            raise

    def _resolve_sheet_name(
        self,
        sheetnames: list[str],
        requested: str,
    ) -> tuple[str | None, list[str]]:
        """
        sheet_name を解決する
        - 完全一致 → そのまま
        - trim + casefold 一致が 1件 → そのシート名に解決
        - それ以外 → None と候補（曖昧一致 or 類似名）を返す
        """
        if requested in sheetnames:
            return (requested, [])

        req_norm = requested.strip().casefold()

        norm_map: dict[str, list[str]] = {}
        for sn in sheetnames:
            key = sn.strip().casefold()
            norm_map.setdefault(key, []).append(sn)

        # 正規化一致
        if req_norm in norm_map:
            candidates = norm_map[req_norm]
            if len(candidates) == 1:
                return (candidates[0], [])
            # 同一正規化で複数（曖昧）
            return (None, candidates)

        # 類似名候補（表示用）
        suggestions = difflib.get_close_matches(requested, sheetnames, n=3, cutoff=0.6)
        return (None, suggestions)

    def _scan_sheet(
        self,
        sheet,
        sheet_name_for_result: str,
        query: str,
        matches: list[dict[str, Any]],
    ) -> None:
        """
        シート内のセルを走査してqueryに一致するセルをmatchesに追加する
        """
        # 空シートを避ける意図
        if sheet.dimensions:
            # パフォーマンスのため_cellsを優先し、無い場合は公開APIにフォールバック
            if hasattr(sheet, "_cells"):
                # 実在セルのみを走査（高速）
                for cell in sheet._cells.values():
                    if cell.value is not None:
                        cell_value_str = str(cell.value)
                        if query in cell_value_str:
                            matches.append(
                                {
                                    "sheet": sheet_name_for_result,
                                    "coordinate": cell.coordinate,
                                    "value": self._serialize_value(cell.value),
                                }
                            )
            else:
                # openpyxl公開APIを使用（互換性確保）
                for row in sheet.iter_rows(values_only=False):
                    for cell in row:
                        if cell.value is not None:
                            cell_value_str = str(cell.value)
                            if query in cell_value_str:
                                matches.append(
                                    {
                                        "sheet": sheet_name_for_result,
                                        "coordinate": cell.coordinate,
                                        "value": self._serialize_value(cell.value),
                                    }
                                )

    def _calculate_header_range(self, cell_range: str, frozen_rows: int) -> str | None:
        """
        セル範囲に対してfrozen_rowsに基づくヘッダー範囲を計算

        Args:
            cell_range: セル範囲（例: "A5:D10"）
                       拡張後のeffective_rangeを渡すこと（軸拡張済み）
            frozen_rows: 固定行数

        Returns:
            ヘッダー範囲（例: "A1:D2"）またはNone

        早期リターン条件:
        - frozen_rows=0: ヘッダーなし
        - start_row == 1: 既に1行目から開始（ヘッダー全体を含む）

        部分的な重なり処理:
        - frozen_rows=2, cell_range="A2:B6" の場合
          → 不足分 "A1:B1" を返して、最終的に "A1:B6" になる
        """
        # frozen_rowsが0の場合はヘッダーなし
        if frozen_rows == 0:
            return None

        # セル範囲を解析
        # "A5:D10" -> start="A5", end="D10"
        if ":" in cell_range:
            start, end = cell_range.split(":")
        else:
            # 単一セル（例: "B5"）
            start = end = cell_range

        # 開始セルの座標を解析
        start_col_letter, start_row = coordinate_from_string(start)
        end_col_letter, _ = coordinate_from_string(end)

        # 既に1行目から開始している場合は追加不要（ヘッダー全体を含む）
        if start_row == 1:
            return None

        # 部分的な重なりがある場合は、不足している上部のヘッダー行を追加
        if start_row <= frozen_rows:
            # 1行目から(start_row-1)行目までを追加
            header_range = f"{start_col_letter}1:{end_col_letter}{start_row - 1}"
            return header_range

        # ヘッダー範囲を計算: {start_col}1:{end_col}{frozen_rows}
        header_range = f"{start_col_letter}1:{end_col_letter}{frozen_rows}"
        return header_range

    def _merge_ranges(self, range1: str, range2: str) -> str:
        """
        2つのセル範囲を結合して、最小の包含範囲を返す

        Args:
            range1: 範囲1（例: "A1:B2"）
            range2: 範囲2（例: "A4:B6"）

        Returns:
            結合された範囲（例: "A1:B6"）
        """
        # 範囲1を解析
        if ":" in range1:
            start1, end1 = range1.split(":")
        else:
            start1 = end1 = range1

        # 範囲2を解析
        if ":" in range2:
            start2, end2 = range2.split(":")
        else:
            start2 = end2 = range2

        # 座標を取得
        col1_start, row1_start = coordinate_from_string(start1)
        col1_end, row1_end = coordinate_from_string(end1)
        col2_start, row2_start = coordinate_from_string(start2)
        col2_end, row2_end = coordinate_from_string(end2)

        # 最小/最大の列を決定
        col_start_idx = min(
            column_index_from_string(col1_start), column_index_from_string(col2_start)
        )
        col_end_idx = max(
            column_index_from_string(col1_end), column_index_from_string(col2_end)
        )

        # 最小/最大の行を決定
        row_start = min(row1_start, row2_start)
        row_end = max(row1_end, row2_end)

        # 列インデックスを文字に変換
        col_start = get_column_letter(col_start_idx)
        col_end = get_column_letter(col_end_idx)

        return f"{col_start}{row_start}:{col_end}{row_end}"

    def _parse_sheet(
        self,
        sheet,
        cell_range: str | None = None,
        include_frozen_rows: bool = True,
        include_cell_styles: bool = False,
    ) -> dict[str, Any]:
        """
        シートを解析してdict形式で返す

        Args:
            sheet: openpyxl Worksheet
            cell_range: セル範囲指定（例: "A1:D10"）
            include_frozen_rows: cell_range指定時に固定行（ヘッダー）を自動追加
            include_cell_styles: セルのスタイル情報を含めるか

        Returns:
            シートデータのdict
        """
        sheet_data = {
            "name": sheet.title,
        }

        # dimensionsがNoneでない場合のみ追加
        if sheet.dimensions:
            sheet_data["dimensions"] = str(sheet.dimensions)

        # freeze_panes情報の取得と検証
        frozen_rows = 0
        frozen_cols = 0
        frozen_rows, frozen_cols = self._get_frozen_panes(sheet)

        # frozen_rows検証（DoS対策）
        # frozen_rowsは補助的なメタ情報なので、上限超過時はリセットして処理を続行
        if frozen_rows > config.excel_max_frozen_rows:
            logger.warning(
                "固定行数が上限(%d)を超えたため、freeze_panes情報を無視します。"
                "ファイルの解析は続行されますが、ヘッダー自動追加機能は利用できません。"
                " (frozen_rows=%d, sheet=%s)",
                config.excel_max_frozen_rows,
                frozen_rows,
                sheet.title,
            )
            frozen_rows = 0
            frozen_cols = 0  # freeze_panes全体を無視

        if frozen_rows > 0 or frozen_cols > 0:
            sheet_data["freeze_panes"] = self._format_freeze_panes(
                frozen_rows, frozen_cols
            )
        sheet_data["frozen_rows"] = frozen_rows
        sheet_data["frozen_cols"] = frozen_cols

        # セル範囲の正規化・拡張（cell_rangeがある場合）
        # マージセル情報のキャッシュに使用するため、先に計算する
        effective_range_for_merge = None
        header_range = None  # ヘッダー範囲（再利用のため事前に初期化）
        if cell_range:
            sheet_data["requested_range"] = cell_range
            effective_range = self._normalize_column_range(cell_range, sheet)
            expanded_range = self._expand_axis_range(effective_range)
            if expanded_range != effective_range:
                logger.info(
                    "Expanded axis range '%s' -> '%s' (sheet=%s)",
                    effective_range,
                    expanded_range,
                    sheet.title,
                )
                effective_range = expanded_range

            if effective_range != cell_range:
                logger.info(
                    "Normalized column range '%s' -> '%s' (sheet=%s)",
                    cell_range,
                    effective_range,
                    sheet.title,
                )
            sheet_data["effective_range"] = effective_range
            effective_range_for_merge = effective_range

            # ヘッダー自動追加の場合、マージセルキャッシュにもヘッダー範囲を含める
            if include_frozen_rows and frozen_rows > 0:
                header_range = self._calculate_header_range(
                    effective_range, frozen_rows
                )
                if header_range:
                    # ヘッダー範囲とデータ範囲を結合した範囲を計算
                    effective_range_for_merge = self._merge_ranges(
                        header_range, effective_range
                    )

        # データサイズ検証（DoS対策）
        # マージセルキャッシュ構築前に検証することで、巨大な範囲によるメモリ枯渇を防ぐ
        all_rows = []

        if cell_range:
            # effective_rangeは既に計算済み
            # データサイズ検証（DoS対策）
            range_rows, range_cols = self._calculate_range_size(effective_range)
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

        # データサイズ検証後にマージセル情報をキャッシュ（パフォーマンス最適化 + DoS対策）
        # 計算済みのeffective_range(effective_range_for_merge)を渡してキャッシュを構築し、
        # 戻り値としてmerged_ranges(結合セル範囲の一覧)を取得することで重複計算を回避
        merged_cell_map, merged_anchor_value_map, merged_ranges = (
            self._build_merged_cell_cache(sheet, effective_range_for_merge)
        )

        # ここは「結合セルがある時だけ」返す
        if merged_ranges:
            sheet_data["merged_ranges"] = merged_ranges

        # セルサイズのキャッシュを構築（パフォーマンス最適化）
        col_widths: dict[str, float] | None = None
        row_heights: dict[int, float] | None = None
        if include_cell_styles:
            col_widths = {}
            row_heights = {}
            for col_letter, dim in sheet.column_dimensions.items():
                if dim.width:
                    col_widths[col_letter] = dim.width
            for row_num, dim in sheet.row_dimensions.items():
                if dim.height:
                    row_heights[row_num] = dim.height

        # データ取得
        if cell_range:
            # ヘッダー自動追加（include_frozen_rows=Trueの場合）
            # header_rangeは既に計算済みなので再利用
            if header_range:
                # ヘッダー範囲を取得
                header_data = sheet[header_range]
                header_rows = self._normalize_range_data(header_data)
                all_rows.extend(
                    self._parse_rows(
                        header_rows,
                        include_cell_styles,
                        merged_cell_map,
                        merged_anchor_value_map,
                        col_widths,
                        row_heights,
                    )
                )

            # 通常のセル範囲取得（データ範囲）
            range_data = sheet[effective_range]
            rows_to_process = self._normalize_range_data(range_data)
            all_rows.extend(
                self._parse_rows(
                    rows_to_process,
                    include_cell_styles,
                    merged_cell_map,
                    merged_anchor_value_map,
                    col_widths,
                    row_heights,
                )
            )

        elif sheet.dimensions:
            # 全データを取得
            rows_to_process = tuple(sheet.iter_rows())

            if rows_to_process:
                all_rows.extend(
                    self._parse_rows(
                        rows_to_process,
                        include_cell_styles,
                        merged_cell_map,
                        merged_anchor_value_map,
                        col_widths,
                        row_heights,
                    )
                )

        sheet_data["rows"] = all_rows
        return sheet_data

    def _build_merged_cell_cache(
        self,
        sheet,
        effective_cell_range: str | None,
    ) -> tuple[
        dict[str, str] | None,
        dict[str, Any] | None,
        list[dict[str, Any]],
    ]:
        """
        マージセル情報をキャッシュして返す（パフォーマンス最適化）
        - 「今回返す予定の範囲」を先に確定し、その範囲と交差する結合だけを部分展開する
        - アンカー値は左上→無ければ結合範囲内の実在セルのみから最小(row,col)を選ぶ

        Args:
            sheet: openpyxl Worksheet
            effective_cell_range: 正規化・拡張済みのセル範囲（例: "A1:D10"）
                Noneの場合はsheet.dimensionsを使用

        Returns:
            (merged_cell_map, merged_anchor_value_map, merged_ranges)のタプル
        """
        merged_cell_map: dict[str, str] | None = None
        merged_anchor_value_map: dict[str, Any] | None = None
        merged_ranges: list[dict[str, Any]] = []

        # 今回返す予定の範囲（結合情報の部分展開に使用）
        # effective_cell_rangeがあればそれを使用、なければsheet.dimensionsを使用
        planned_range_for_merge = effective_cell_range or (
            str(sheet.dimensions) if sheet.dimensions else None
        )

        if not sheet.merged_cells.ranges or not planned_range_for_merge:
            return (None, None, [])

        # planned_range_for_merge から対象範囲の境界を計算
        if ":" in planned_range_for_merge:
            start_cell, end_cell = planned_range_for_merge.split(":", 1)
        else:
            start_cell = planned_range_for_merge
            end_cell = planned_range_for_merge

        start_cell = start_cell.replace("$", "")
        end_cell = end_cell.replace("$", "")

        start_col, start_row = coordinate_from_string(start_cell)
        end_col, end_row = coordinate_from_string(end_cell)

        start_col_idx = column_index_from_string(start_col)
        end_col_idx = column_index_from_string(end_col)

        target_min_row = min(start_row, end_row)
        target_max_row = max(start_row, end_row)
        target_min_col = min(start_col_idx, end_col_idx)
        target_max_col = max(start_col_idx, end_col_idx)

        merged_cell_map = {}
        merged_anchor_value_map = {}

        for merged_range in sheet.merged_cells.ranges:
            merged_range_str = str(merged_range)
            range_start = merged_range_str.split(":")[0]

            merged_min_row = merged_range.min_row
            merged_max_row = merged_range.max_row
            merged_min_col = merged_range.min_col
            merged_max_col = merged_range.max_col

            # 返す予定の範囲と交差しない結合は無視（部分展開）
            inter_min_row = max(merged_min_row, target_min_row)
            inter_max_row = min(merged_max_row, target_max_row)
            inter_min_col = max(merged_min_col, target_min_col)
            inter_max_col = min(merged_max_col, target_max_col)
            if inter_min_row > inter_max_row or inter_min_col > inter_max_col:
                continue

            # アンカー値を決定（左上が空なら結合範囲内の実在セルだけ走査）
            anchor_coord = range_start
            anchor_value = self._serialize_value(sheet[range_start].value)

            if anchor_value is None:
                best_rc: tuple[int, int] | None = None
                best_val: Any | None = None

                # 実在セル（sheet._cells）だけから、結合範囲内の最小(row,col)の値を選ぶ
                # 互換性のため_cellsの有無をチェックしてフォールバック
                if hasattr(sheet, "_cells"):
                    # プライベート属性を使った高速版
                    for (r, c), cell_obj in sheet._cells.items():
                        if (
                            merged_min_row <= r <= merged_max_row
                            and merged_min_col <= c <= merged_max_col
                        ):
                            cell_value = self._serialize_value(cell_obj.value)
                            if cell_value is not None:
                                if best_rc is None or (r, c) < best_rc:
                                    best_rc = (r, c)
                                    best_val = cell_value
                else:
                    # 公開APIを使ったフォールバック版
                    for row_idx in range(merged_min_row, merged_max_row + 1):
                        for col_idx in range(merged_min_col, merged_max_col + 1):
                            coord = f"{get_column_letter(col_idx)}{row_idx}"
                            cell = sheet[coord]
                            cell_value = self._serialize_value(cell.value)
                            if cell_value is not None:
                                if best_rc is None or (row_idx, col_idx) < best_rc:
                                    best_rc = (row_idx, col_idx)
                                    best_val = cell_value

                if best_rc is not None:
                    r, c = best_rc
                    anchor_value = best_val
                    anchor_coord = f"{get_column_letter(c)}{r}"

            # セル座標 -> 結合範囲 のマップ（返す予定の範囲と交差する部分だけ展開）
            for row_idx in range(inter_min_row, inter_max_row + 1):
                for col_idx in range(inter_min_col, inter_max_col + 1):
                    coord_str = f"{get_column_letter(col_idx)}{row_idx}"
                    merged_cell_map[coord_str] = merged_range_str

            # アンカー値を保存（結合セルの値埋め用）
            merged_anchor_value_map[merged_range_str] = anchor_value

            # 結合範囲そのものを返す（結合セルがある時だけ返す）
            merged_ranges.append(
                {
                    "range": merged_range_str,
                    "anchor": {"coordinate": anchor_coord, "value": anchor_value},
                }
            )

        if not merged_ranges:
            return (None, None, [])

        return (merged_cell_map, merged_anchor_value_map, merged_ranges)

    def _expand_axis_range(self, range_str: str) -> str:
        """
        指定されたセル範囲を「枠分離」ではなく「方向に拡張」する。
        - 単一列 (例: Z100:Z200) -> Z1:Z200
        - 単一行 (例: D200:Z200) -> A200:Z200
        - それ以外（矩形など）はそのまま
        """
        if not range_str:
            return range_str

        raw = range_str.strip()
        if ":" not in raw:
            try:
                col, row = coordinate_from_string(raw.replace("$", ""))
                return f"{col}1:{col}{row}"
            except Exception:
                return range_str

        start_cell, end_cell = raw.split(":", 1)
        start_cell = start_cell.replace("$", "")
        end_cell = end_cell.replace("$", "")

        start_col, start_row = coordinate_from_string(start_cell)
        end_col, end_row = coordinate_from_string(end_cell)

        # 列指定（同一列）: 逆順はそのまま（既存のrange検証で弾く）
        if start_col == end_col:
            if end_row < start_row:
                return range_str
            return f"{start_col}1:{end_col}{end_row}"

        # 行指定（同一行）: 逆順はそのまま（既存のrange検証で弾く）
        if start_row == end_row:
            if column_index_from_string(end_col) < column_index_from_string(start_col):
                return range_str
            return f"A{start_row}:{end_col}{end_row}"

        return range_str

    def _parse_cell(
        self,
        cell,
        include_cell_styles: bool = False,
        merged_cell_map: dict[str, str] | None = None,
        merged_anchor_value_map: dict[str, Any] | None = None,
        col_widths: dict[str, float] | None = None,
        row_heights: dict[int, float] | None = None,
    ) -> dict[str, Any]:
        """
        セルを解析してdict形式で返す

        Args:
            cell: openpyxl Cell
            include_cell_styles: セルのスタイル情報を含めるか（デフォルト: False）
            merged_cell_map: マージセル座標からマージ範囲へのマップ（パフォーマンス最適化用）
            merged_anchor_value_map: マージ範囲 -> アンカー値 のマップ（結合セルの値埋め用）
            col_widths: 列幅のキャッシュ（パフォーマンス最適化用）
            row_heights: 行高さのキャッシュ（パフォーマンス最適化用）

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

        # スタイル情報（include_cell_styles=Trueの場合のみ）
        if include_cell_styles:
            # 背景色情報
            if cell.fill and cell.fill.patternType:
                fill_info = {
                    "pattern_type": cell.fill.patternType,
                }
                fg_color = self._color_to_hex(cell.fill.fgColor)
                if fg_color:
                    fill_info["fg_color"] = fg_color
                bg_color = self._color_to_hex(cell.fill.bgColor)
                if bg_color:
                    fill_info["bg_color"] = bg_color
                cell_data["fill"] = fill_info

            # セルサイズ（列幅・行高さ）
            # MergedCellの場合は属性が存在しないため、hasattrでチェック
            if hasattr(cell, "column_letter") and hasattr(cell, "row"):
                if cell.column_letter and cell.row:
                    # キャッシュから列幅を取得（パフォーマンス最適化）
                    if col_widths and cell.column_letter in col_widths:
                        cell_data["width"] = col_widths[cell.column_letter]
                    # キャッシュから行高さを取得（パフォーマンス最適化）
                    if row_heights and cell.row in row_heights:
                        cell_data["height"] = row_heights[cell.row]

        return cell_data

    def _parse_rows(
        self,
        rows: tuple[tuple[Cell, ...], ...],
        include_cell_styles: bool = False,
        merged_cell_map: dict[str, str] | None = None,
        merged_anchor_value_map: dict[str, Any] | None = None,
        col_widths: dict[str, float] | None = None,
        row_heights: dict[int, float] | None = None,
    ) -> list[list[dict[str, Any]]]:
        """
        行データを解析してリスト形式で返す（コード重複削減用ヘルパー）

        Args:
            rows: 行データのタプル
            include_cell_styles: セルのスタイル情報を含めるか
            merged_cell_map: マージセル情報
            merged_anchor_value_map: マージ範囲 -> アンカー値
            col_widths: 列幅のキャッシュ（パフォーマンス最適化用）
            row_heights: 行高さのキャッシュ（パフォーマンス最適化用）

        Returns:
            解析された行データのリスト
        """
        parsed_rows = []
        for row in rows:
            row_data = [
                self._parse_cell(
                    cell,
                    include_cell_styles,
                    merged_cell_map,
                    merged_anchor_value_map,
                    col_widths,
                    row_heights,
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

    def _normalize_column_range(self, cell_range: str, sheet) -> str:
        """
        列のみ指定された範囲（例: "J:J" / "J"）を行番号付きに正規化する

        Args:
            cell_range: セル範囲
            sheet: openpyxl Worksheet

        Returns:
            正規化されたセル範囲
        """
        raw = cell_range.strip()
        if not raw:
            return cell_range

        # "J:J" のような列のみ指定
        if ":" in raw:
            start, end = raw.split(":", 1)
            start_col = start.replace("$", "")
            end_col = end.replace("$", "")
            if start_col.isalpha() and end_col.isalpha():
                start_col = start_col.upper()
                end_col = end_col.upper()
                # 逆順序の列を検出
                if column_index_from_string(end_col) < column_index_from_string(
                    start_col
                ):
                    raise ValueError(
                        f"無効なセル範囲: '{cell_range}'。"
                        f"範囲は正しい順序で指定してください（例: 'A1:Z100'）"
                    )
                max_row = sheet.max_row or 1
                return f"{start_col}1:{end_col}{max_row}"

        # "J" のような単一列指定
        col_only = raw.replace("$", "")
        if col_only.isalpha():
            col_only = col_only.upper()
            max_row = sheet.max_row or 1
            return f"{col_only}1:{col_only}{max_row}"

        return cell_range
