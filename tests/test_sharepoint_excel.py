import datetime
import json
from io import BytesIO
from unittest.mock import Mock

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.views import Pane

from src.sharepoint_excel import SharePointExcelParser


class TestSharePointExcelParser:
    """SharePoint Excel解析のテスト"""

    def setup_method(self):
        """テストメソッド実行前のセットアップ"""
        self.mock_download_client = Mock()

    def _create_test_excel(self) -> bytes:
        """テスト用のシンプルなExcelファイルを作成"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["A1"] = "Name"
        ws["B1"] = "Age"
        ws["A2"] = "John"
        ws["B2"] = 25

        # BytesIOに保存
        excel_bytes = BytesIO()
        wb.save(excel_bytes)
        excel_bytes.seek(0)
        return excel_bytes.getvalue()

    def _create_formatted_excel(self) -> bytes:
        """書式を含むテスト用Excelファイルを作成"""
        wb = Workbook()
        ws = wb.active
        ws.title = "FormattedSheet"

        # ヘッダー行に書式を設定
        ws["A1"] = "Name"
        ws["A1"].font = Font(name="Arial", size=12, bold=True, color="FF0000")
        ws["A1"].fill = PatternFill(
            start_color="FFFF00", end_color="FFFF00", fill_type="solid"
        )

        ws["B1"] = "Value"
        ws["B1"].font = Font(name="Arial", size=12, bold=True)

        # データ行
        ws["A2"] = "Item1"
        ws["B2"] = 100

        # BytesIOに保存
        excel_bytes = BytesIO()
        wb.save(excel_bytes)
        excel_bytes.seek(0)
        return excel_bytes.getvalue()

    def _create_multi_sheet_excel(self) -> bytes:
        """複数シートを含むテスト用Excelファイルを作成"""
        wb = Workbook()

        # 最初のシート
        ws1 = wb.active
        ws1.title = "Sheet1"
        ws1["A1"] = "Data1"

        # 2つ目のシート
        ws2 = wb.create_sheet("Sheet2")
        ws2["A1"] = "Data2"

        # BytesIOに保存
        excel_bytes = BytesIO()
        wb.save(excel_bytes)
        excel_bytes.seek(0)
        return excel_bytes.getvalue()

    def _create_merged_cells_excel(self) -> bytes:
        """結合セルを含むテスト用Excelファイルを作成"""
        wb = Workbook()
        ws = wb.active
        ws.title = "MergedSheet"

        # セルを結合
        ws.merge_cells("A1:B1")
        ws["A1"] = "Merged Header"

        ws["A2"] = "Data1"
        ws["B2"] = "Data2"

        # BytesIOに保存
        excel_bytes = BytesIO()
        wb.save(excel_bytes)
        excel_bytes.seek(0)
        return excel_bytes.getvalue()

    def _create_frozen_panes_excel(self, freeze_panes: str) -> bytes:
        """固定行・列を含むテスト用Excelファイルを作成"""
        wb = Workbook()
        ws = wb.active
        ws.title = "FrozenSheet"

        # ヘッダー行を作成
        ws["A1"] = "Header1"
        ws["B1"] = "Header2"
        ws["C1"] = "Header3"
        ws["D1"] = "Header4"

        # データ行を作成
        for row in range(2, 11):
            for col in range(1, 5):
                ws.cell(row=row, column=col, value=f"Data{row-1}_{col}")

        # freeze_panesを設定
        ws.freeze_panes = freeze_panes

        # BytesIOに保存
        excel_bytes = BytesIO()
        wb.save(excel_bytes)
        excel_bytes.seek(0)
        return excel_bytes.getvalue()

    def test_parse_simple_excel(self):
        """シンプルなExcelファイルの解析テスト（デフォルト：最小限の情報）"""
        excel_bytes = self._create_test_excel()
        self.mock_download_client.download_file.return_value = excel_bytes

        # 解析
        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json("/test/file.xlsx")

        # 検証
        result = json.loads(result_json)
        assert result["file_path"] == "/test/file.xlsx"
        assert len(result["sheets"]) == 1
        assert result["sheets"][0]["name"] == "Sheet1"
        assert len(result["sheets"][0]["rows"]) == 2

        # デフォルトでは value と coordinate のみ
        cell = result["sheets"][0]["rows"][0][0]
        assert cell["value"] == "Name"
        assert cell["coordinate"] == "A1"
        assert "data_type" not in cell
        assert "fill" not in cell
        assert "width" not in cell

        cell2 = result["sheets"][0]["rows"][1][1]
        assert cell2["value"] == 25
        assert cell2["coordinate"] == "B2"

    def test_parse_with_formatting(self):
        """include_formatting=Trueでも出力が変わらないことのテスト"""
        excel_bytes = self._create_formatted_excel()
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json("/test/formatted.xlsx", include_formatting=True)

        result = json.loads(result_json)
        assert result["sheets"][0]["name"] == "FormattedSheet"

        # include_formatting=Trueでも書式情報は追加されない
        header_cell = result["sheets"][0]["rows"][0][0]
        assert header_cell["value"] == "Name"
        assert "data_type" not in header_cell
        assert "fill" not in header_cell
        # fontとalignmentは含まれない
        assert "font" not in header_cell
        assert "alignment" not in header_cell

    def test_parse_multiple_sheets(self):
        """複数シートのExcelファイルの解析テスト"""
        excel_bytes = self._create_multi_sheet_excel()
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json("/test/multi.xlsx")

        result = json.loads(result_json)
        assert len(result["sheets"]) == 2
        assert result["sheets"][0]["name"] == "Sheet1"
        assert result["sheets"][1]["name"] == "Sheet2"
        assert result["sheets"][0]["rows"][0][0]["value"] == "Data1"
        assert result["sheets"][1]["rows"][0][0]["value"] == "Data2"

    def test_parse_merged_cells(self):
        """結合セルの解析テスト（結合情報は常に含まれる）"""
        excel_bytes = self._create_merged_cells_excel()
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json("/test/merged.xlsx", include_formatting=True)

        result = json.loads(result_json)
        assert result["sheets"][0]["name"] == "MergedSheet"

        # 結合セルの情報を確認（include_formattingに関係なく含まれる）
        merged_cell = result["sheets"][0]["rows"][0][0]
        assert merged_cell["value"] == "Merged Header"
        assert "merged" in merged_cell
        assert merged_cell["merged"]["range"] == "A1:B1"
        assert merged_cell["merged"]["is_top_left"] is True

    def test_download_error_handling(self):
        """ダウンロードエラーのハンドリングテスト"""
        self.mock_download_client.download_file.side_effect = Exception(
            "Download failed"
        )

        parser = SharePointExcelParser(self.mock_download_client)
        with pytest.raises(Exception) as exc_info:
            parser.parse_to_json("/test/file.xlsx")

        assert "Download failed" in str(exc_info.value)

    def test_invalid_excel_file(self):
        """無効なExcelファイルの処理テスト"""
        # 無効なバイトデータを返す
        self.mock_download_client.download_file.return_value = b"invalid excel data"

        parser = SharePointExcelParser(self.mock_download_client)
        with pytest.raises(Exception):
            parser.parse_to_json("/test/invalid.xlsx")

    def test_empty_excel_file(self):
        """空のExcelファイルの解析テスト"""
        wb = Workbook()
        ws = wb.active
        ws.title = "EmptySheet"
        # データを追加しない

        excel_bytes = BytesIO()
        wb.save(excel_bytes)
        excel_bytes.seek(0)

        self.mock_download_client.download_file.return_value = excel_bytes.getvalue()

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json("/test/empty.xlsx")

        result = json.loads(result_json)
        assert result["sheets"][0]["name"] == "EmptySheet"
        # openpyxlは空のシートでも最低1つのセル（A1）を持つ
        assert result["sheets"][0]["dimensions"] is not None
        # 空のシートでも行データが取得できる
        rows = result["sheets"][0]["rows"]
        assert isinstance(rows, list)

    def test_color_to_hex_rgb(self):
        """RGB色の16進数変換テスト"""
        excel_bytes = self._create_formatted_excel()
        self.mock_download_client.download_file.return_value = excel_bytes

        # _color_to_hex の単体動作を確認（include_formattingの有無とは無関係）
        parser = SharePointExcelParser(self.mock_download_client)
        wb = load_workbook(BytesIO(excel_bytes))
        cell = wb.active["A1"]
        hex_color = parser._color_to_hex(cell.fill.fgColor)
        if hex_color:
            assert hex_color.startswith("#")

    def test_parse_with_formulas(self):
        """数式を含むExcelファイルの解析テスト"""
        wb = Workbook()
        ws = wb.active
        ws.title = "FormulaSheet"

        ws["A1"] = 10
        ws["A2"] = 20
        ws["A3"] = "=A1+A2"  # 数式

        excel_bytes = BytesIO()
        wb.save(excel_bytes)
        excel_bytes.seek(0)

        self.mock_download_client.download_file.return_value = excel_bytes.getvalue()

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json("/test/formula.xlsx")

        result = json.loads(result_json)
        # 数式セルの値を確認（data_only=Falseなので数式文字列が入る）
        formula_cell = result["sheets"][0]["rows"][2][0]
        assert formula_cell["value"] == "=A1+A2"

    def test_default_response_is_minimal(self):
        """デフォルトレスポンスが最小限であることのテスト"""
        excel_bytes = self._create_formatted_excel()
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json("/test/formatted.xlsx")

        result = json.loads(result_json)
        cell = result["sheets"][0]["rows"][0][0]

        # デフォルトでは value と coordinate のみ
        assert "value" in cell
        assert "coordinate" in cell
        assert "data_type" not in cell
        assert "fill" not in cell
        assert "font" not in cell
        assert "alignment" not in cell
        assert "merged" not in cell
        assert "width" not in cell
        assert "height" not in cell

    def test_formatting_included_when_requested(self):
        """include_formatting=Trueでも追加の書式情報が含まれないことのテスト"""
        excel_bytes = self._create_formatted_excel()
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json("/test/formatted.xlsx", include_formatting=True)

        result = json.loads(result_json)
        cell = result["sheets"][0]["rows"][0][0]

        # include_formatting=True の場合でも追加フィールドはない
        assert "value" in cell
        assert "coordinate" in cell
        assert "data_type" not in cell
        assert "fill" not in cell
        # font と alignment は含まれない
        assert "font" not in cell
        assert "alignment" not in cell

    def test_datetime_serialization(self):
        """datetime型の値が正しくシリアライズされることのテスト"""
        wb = Workbook()
        ws = wb.active
        ws.title = "DateTimeSheet"

        # 各種datetime型の値を設定
        ws["A1"] = datetime.datetime(2024, 1, 15, 14, 30, 45)
        ws["A2"] = datetime.date(2024, 1, 15)
        ws["A3"] = datetime.time(14, 30, 45)
        ws["A4"] = datetime.timedelta(days=1, hours=2, minutes=30)

        excel_bytes = BytesIO()
        wb.save(excel_bytes)
        excel_bytes.seek(0)

        self.mock_download_client.download_file.return_value = excel_bytes.getvalue()

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json("/test/datetime.xlsx")

        # JSONパースが成功することを確認（datetime型が適切にシリアライズされている）
        result = json.loads(result_json)
        rows = result["sheets"][0]["rows"]

        # 文字列に変換されていることを確認
        assert rows[0][0]["value"] == "2024-01-15 14:30:45"
        # openpyxlはdateをdatetimeとして読み込む（時刻部分は00:00:00）
        assert rows[1][0]["value"] == "2024-01-15 00:00:00"
        assert rows[2][0]["value"] == "14:30:45"
        # timedeltaは文字列表現に変換される
        assert rows[3][0]["value"] == "1 day, 2:30:00"

    def test_search_cells_basic(self):
        """セル検索の基本テスト"""
        excel_bytes = self._create_test_excel()
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.search_cells("/test/file.xlsx", "John")

        result = json.loads(result_json)
        assert result["file_path"] == "/test/file.xlsx"
        assert result["mode"] == "search"
        assert result["query"] == "John"
        assert result["match_count"] == 1
        assert len(result["matches"]) == 1
        assert result["matches"][0]["sheet"] == "Sheet1"
        assert result["matches"][0]["coordinate"] == "A2"
        assert result["matches"][0]["value"] == "John"

    def test_search_cells_multiple_matches(self):
        """複数マッチする検索のテスト"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["A1"] = "売上報告"
        ws["A2"] = "月間売上"
        ws["B2"] = 1000
        ws["A3"] = "売上合計"
        ws["B3"] = "=SUM(B1:B2)"

        excel_bytes = BytesIO()
        wb.save(excel_bytes)
        excel_bytes.seek(0)

        self.mock_download_client.download_file.return_value = excel_bytes.getvalue()

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.search_cells("/test/file.xlsx", "売上")

        result = json.loads(result_json)
        assert result["match_count"] == 3
        assert len(result["matches"]) == 3
        coordinates = [m["coordinate"] for m in result["matches"]]
        assert "A1" in coordinates
        assert "A2" in coordinates
        assert "A3" in coordinates

    def test_search_cells_no_match(self):
        """マッチしない検索のテスト"""
        excel_bytes = self._create_test_excel()
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.search_cells("/test/file.xlsx", "NotFound")

        result = json.loads(result_json)
        assert result["match_count"] == 0
        assert result["matches"] == []

    def test_search_cells_multiple_sheets(self):
        """複数シートにまたがる検索のテスト"""
        excel_bytes = self._create_multi_sheet_excel()
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.search_cells("/test/file.xlsx", "Data")

        result = json.loads(result_json)
        assert result["match_count"] == 2
        assert len(result["matches"]) == 2
        sheets = [m["sheet"] for m in result["matches"]]
        assert "Sheet1" in sheets
        assert "Sheet2" in sheets

    def test_parse_specific_sheet(self):
        """特定シートのみ取得するテスト"""
        excel_bytes = self._create_multi_sheet_excel()
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json("/test/file.xlsx", sheet_name="Sheet2")

        result = json.loads(result_json)
        assert len(result["sheets"]) == 1
        assert result["sheets"][0]["name"] == "Sheet2"
        assert result["sheets"][0]["rows"][0][0]["value"] == "Data2"

    def test_parse_nonexistent_sheet(self):
        """存在しないシート名を指定した場合の解決情報テスト"""
        excel_bytes = self._create_test_excel()
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json("/test/file.xlsx", sheet_name="NonExistent")

        result = json.loads(result_json)
        assert result["requested_sheet"] == "NonExistent"
        assert result["sheets"] == []
        assert result["sheet_resolution"]["status"] == "not_found"
        assert result["sheet_resolution"]["requested"] == "NonExistent"
        assert result["sheet_resolution"]["resolved"] is None
        assert result["available_sheets"] == ["Sheet1"]
        assert result["warning"] == "requested sheet_name was not found or ambiguous"

    def test_parse_cell_range(self):
        """セル範囲指定のテスト"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        # 5x5のデータを作成
        for row in range(1, 6):
            for col in range(1, 6):
                ws.cell(row=row, column=col, value=f"R{row}C{col}")

        excel_bytes = BytesIO()
        wb.save(excel_bytes)
        excel_bytes.seek(0)

        self.mock_download_client.download_file.return_value = excel_bytes.getvalue()

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json(
            "/test/file.xlsx", sheet_name="Sheet1", cell_range="B2:D4"
        )

        result = json.loads(result_json)
        assert result["sheets"][0]["requested_range"] == "B2:D4"
        # 3行3列のデータが取得される
        assert len(result["sheets"][0]["rows"]) == 3
        assert len(result["sheets"][0]["rows"][0]) == 3
        # 最初のセルはB2（R2C2）
        assert result["sheets"][0]["rows"][0][0]["value"] == "R2C2"
        assert result["sheets"][0]["rows"][0][0]["coordinate"] == "B2"
        # 最後のセルはD4（R4C4）
        assert result["sheets"][0]["rows"][2][2]["value"] == "R4C4"
        assert result["sheets"][0]["rows"][2][2]["coordinate"] == "D4"

    def test_parse_single_cell_range(self):
        """単一セル範囲指定のテスト"""
        excel_bytes = self._create_test_excel()
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json(
            "/test/file.xlsx", sheet_name="Sheet1", cell_range="A1"
        )

        result = json.loads(result_json)
        assert result["sheets"][0]["requested_range"] == "A1"
        assert len(result["sheets"][0]["rows"]) == 1
        assert len(result["sheets"][0]["rows"][0]) == 1
        assert result["sheets"][0]["rows"][0][0]["value"] == "Name"

    def test_parse_sheet_and_range_combined(self):
        """シート名と範囲を組み合わせたテスト"""
        excel_bytes = self._create_multi_sheet_excel()
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json(
            "/test/file.xlsx", sheet_name="Sheet1", cell_range="A1"
        )

        result = json.loads(result_json)
        assert len(result["sheets"]) == 1
        assert result["sheets"][0]["name"] == "Sheet1"
        assert result["sheets"][0]["rows"][0][0]["value"] == "Data1"

    def test_parse_with_freeze_panes_both(self):
        """freeze_panes="B2"（行と列の両方固定）のテスト"""
        excel_bytes = self._create_frozen_panes_excel("B2")
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json("/test/frozen.xlsx")

        result = json.loads(result_json)
        sheet = result["sheets"][0]

        # freeze_panes情報を確認
        assert sheet["freeze_panes"] == "B2"
        assert sheet["frozen_rows"] == 1
        assert sheet["frozen_cols"] == 1

        rows = sheet["rows"]
        assert len(rows) == 10
        assert rows[0][0]["value"] == "Header1"
        assert rows[0][1]["value"] == "Header2"
        assert rows[1][0]["value"] == "Data1_1"
        assert rows[1][1]["value"] == "Data1_2"

    def test_parse_with_freeze_panes_rows_only(self):
        """freeze_panes="A2"（行のみ固定）のテスト"""
        excel_bytes = self._create_frozen_panes_excel("A2")
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json("/test/frozen.xlsx")

        result = json.loads(result_json)
        sheet = result["sheets"][0]

        # freeze_panes情報を確認
        assert sheet["freeze_panes"] == "A2"
        assert sheet["frozen_rows"] == 1
        assert sheet["frozen_cols"] == 0

        rows = sheet["rows"]
        assert len(rows) == 10
        assert rows[0][0]["value"] == "Header1"
        assert rows[1][0]["value"] == "Data1_1"

    def test_parse_with_no_freeze_panes(self):
        """freeze_panes=Noneのテスト"""
        excel_bytes = self._create_test_excel()
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json("/test/file.xlsx")

        result = json.loads(result_json)
        sheet = result["sheets"][0]

        # freeze_panes情報を確認
        assert "freeze_panes" not in sheet
        assert sheet["frozen_rows"] == 0
        assert sheet["frozen_cols"] == 0

        rows = sheet["rows"]
        assert len(rows) == 2
        assert rows[0][0]["value"] == "Name"

    def test_parse_range_with_overlapping_headers(self):
        """cell_range内にヘッダーが含まれる場合のテスト"""
        excel_bytes = self._create_frozen_panes_excel("B2")
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json(
            "/test/frozen.xlsx", cell_range="A1:D5"
        )

        result = json.loads(result_json)
        sheet = result["sheets"][0]

        rows = sheet["rows"]
        assert len(rows) == 5
        assert rows[0][0]["value"] == "Header1"
        assert rows[1][0]["value"] == "Data1_1"
        assert rows[4][0]["value"] == "Data4_1"

    def test_parse_range_with_non_overlapping_headers(self):
        """ヘッダーがcell_range外にある場合のテスト"""
        excel_bytes = self._create_frozen_panes_excel("B2")
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json(
            "/test/frozen.xlsx", cell_range="A5:D10"
        )

        result = json.loads(result_json)
        sheet = result["sheets"][0]

        # requested_rangeは元のまま
        assert sheet["requested_range"] == "A5:D10"

        rows = sheet["rows"]
        assert len(rows) == 6
        assert rows[0][0]["value"] == "Data4_1"
        assert rows[0][0]["coordinate"] == "A5"

    def test_freeze_panes_scrolled_position_does_not_affect_frozen_rows(self):
        """スクロール後に保存されたファイルでfrozen_rowsが正しく取得されるテスト

        Excel上で3行固定して行450付近にスクロールして保存すると、
        pane.topLeftCell="A450"になるが、pane.ySplit=3は不変。
        旧実装ではsheet.freeze_panes（=topLeftCell）を解析していたため
        frozen_rows=449と誤判定していたバグを検証する。
        """
        wb = Workbook()
        ws = wb.active
        ws.title = "ScrolledSheet"

        # ヘッダー3行 + データ行
        for row in range(1, 11):
            ws.cell(row=row, column=1, value=f"Row{row}")

        # 3行固定を設定
        ws.freeze_panes = "A4"

        # スクロール位置を変更（pane.topLeftCellを直接操作）
        pane = ws.sheet_view.pane
        pane.topLeftCell = "A450"

        excel_bytes = BytesIO()
        wb.save(excel_bytes)
        excel_bytes.seek(0)

        self.mock_download_client.download_file.return_value = excel_bytes.getvalue()

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json("/test/scrolled.xlsx")

        result = json.loads(result_json)
        sheet = result["sheets"][0]

        # pane.ySplit=3なので、frozen_rowsは3であるべき（449ではない）
        assert sheet["frozen_rows"] == 3
        assert sheet["frozen_cols"] == 0
        assert sheet["freeze_panes"] == "A4"

        rows = sheet["rows"]
        assert len(rows) == 10
        assert rows[0][0]["value"] == "Row1"
        assert rows[2][0]["value"] == "Row3"

    def test_split_pane_is_ignored(self):
        """split pane（state="split"）は固定行として認識されないテスト"""
        wb = Workbook()
        ws = wb.active
        ws.title = "SplitSheet"

        ws["A1"] = "Header"
        ws["A2"] = "Data"

        # split paneを設定（frozenではなくsplit）
        ws.sheet_view.pane = Pane(ySplit=3, xSplit=0, state="split")

        excel_bytes = BytesIO()
        wb.save(excel_bytes)
        excel_bytes.seek(0)

        self.mock_download_client.download_file.return_value = excel_bytes.getvalue()

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json("/test/split.xlsx")

        result = json.loads(result_json)
        sheet = result["sheets"][0]

        # split paneは固定行として認識されない
        assert sheet["frozen_rows"] == 0
        assert sheet["frozen_cols"] == 0
        assert "freeze_panes" not in sheet
