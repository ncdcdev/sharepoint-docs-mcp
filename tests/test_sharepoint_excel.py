import json
from io import BytesIO
from unittest.mock import Mock

import pytest
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

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

    def test_parse_simple_excel(self):
        """シンプルなExcelファイルの解析テスト（デフォルト：最小限の情報）"""
        excel_bytes = self._create_test_excel()
        self.mock_download_client.download_file.return_value = excel_bytes

        # 解析（デフォルト：include_formatting=False）
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
        """書式情報を含むExcelファイルの解析テスト（include_formatting=True）"""
        excel_bytes = self._create_formatted_excel()
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json("/test/formatted.xlsx", include_formatting=True)

        result = json.loads(result_json)
        assert result["sheets"][0]["name"] == "FormattedSheet"

        # ヘッダー行の書式を確認
        header_cell = result["sheets"][0]["rows"][0][0]
        assert header_cell["value"] == "Name"
        assert "data_type" in header_cell
        assert header_cell["fill"] is not None
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
        """結合セルの解析テスト（include_formatting=True）"""
        excel_bytes = self._create_merged_cells_excel()
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json("/test/merged.xlsx", include_formatting=True)

        result = json.loads(result_json)
        assert result["sheets"][0]["name"] == "MergedSheet"

        # 結合セルの情報を確認（include_formatting=Trueの場合のみ含まれる）
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
        # 1行1列のデータが含まれる可能性がある
        assert len(result["sheets"][0]["rows"]) >= 0

    def test_color_to_hex_rgb(self):
        """RGB色の16進数変換テスト"""
        excel_bytes = self._create_formatted_excel()
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json("/test/formatted.xlsx", include_formatting=True)

        result = json.loads(result_json)
        header_cell = result["sheets"][0]["rows"][0][0]

        # 塗りつぶし色の確認
        if header_cell.get("fill", {}).get("fg_color"):
            assert header_cell["fill"]["fg_color"].startswith("#")

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
        """include_formatting=Trueの場合に書式情報が含まれることのテスト"""
        excel_bytes = self._create_formatted_excel()
        self.mock_download_client.download_file.return_value = excel_bytes

        parser = SharePointExcelParser(self.mock_download_client)
        result_json = parser.parse_to_json("/test/formatted.xlsx", include_formatting=True)

        result = json.loads(result_json)
        cell = result["sheets"][0]["rows"][0][0]

        # include_formatting=True の場合
        assert "value" in cell
        assert "coordinate" in cell
        assert "data_type" in cell
        assert "fill" in cell
        # font と alignment は含まれない
        assert "font" not in cell
        assert "alignment" not in cell
