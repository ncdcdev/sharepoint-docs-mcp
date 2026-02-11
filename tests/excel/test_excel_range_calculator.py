"""
ExcelRangeCalculatorのテスト
"""

import pytest

from src.excel import ExcelRangeCalculator


class TestExcelRangeCalculator:
    """ExcelRangeCalculator（範囲計算）のテスト"""

    # calculate_header_range のテスト

    def test_calculate_header_range_with_frozen_rows(self):
        """frozen_rows > 0 でヘッダー範囲が計算されること"""
        # frozen_rows=2, cell_range="A5:D10" -> "A1:D2"
        result = ExcelRangeCalculator.calculate_header_range("A5:D10", 2)
        assert result == "A1:D2"

    def test_calculate_header_range_frozen_rows_zero(self):
        """frozen_rows=0 の場合はNoneが返ること"""
        result = ExcelRangeCalculator.calculate_header_range("A5:D10", 0)
        assert result is None

    def test_calculate_header_range_already_includes_row1(self):
        """cell_rangeが既に1行目を含む場合はNoneが返ること"""
        result = ExcelRangeCalculator.calculate_header_range("A1:D10", 2)
        assert result is None

    def test_calculate_header_range_partial_overlap(self):
        """部分的な重なりがある場合は不足分のみ返すこと"""
        # frozen_rows=2, cell_range="A2:B6" -> 不足分 "A1:B1"
        result = ExcelRangeCalculator.calculate_header_range("A2:B6", 2)
        assert result == "A1:B1"

    def test_calculate_header_range_single_cell(self):
        """単一セルの場合も正しく処理されること"""
        # frozen_rows=2, cell_range="B5" -> "B1:B2"
        result = ExcelRangeCalculator.calculate_header_range("B5", 2)
        assert result == "B1:B2"

    # merge_ranges のテスト

    def test_merge_ranges_basic(self):
        """2つの範囲が正しく結合されること"""
        result = ExcelRangeCalculator.merge_ranges("A1:B2", "A4:B6")
        assert result == "A1:B6"

    def test_merge_ranges_overlapping(self):
        """重なる範囲が正しく結合されること"""
        result = ExcelRangeCalculator.merge_ranges("A1:C5", "B3:D7")
        assert result == "A1:D7"

    def test_merge_ranges_single_cells(self):
        """単一セル同士の結合が正しく処理されること"""
        result = ExcelRangeCalculator.merge_ranges("A1", "C3")
        assert result == "A1:C3"

    def test_merge_ranges_different_columns(self):
        """異なる列範囲の結合が正しく処理されること"""
        result = ExcelRangeCalculator.merge_ranges("A1:A5", "C1:C5")
        assert result == "A1:C5"

    # expand_axis_range のテスト

    def test_expand_axis_range_single_cell(self):
        """単一セルが列範囲に拡張されること"""
        result = ExcelRangeCalculator.expand_axis_range("C5")
        assert result == "C1:C5"

    def test_expand_axis_range_single_column(self):
        """単一列が1行目まで拡張されること"""
        result = ExcelRangeCalculator.expand_axis_range("Z100:Z200")
        assert result == "Z1:Z200"

    def test_expand_axis_range_single_row(self):
        """単一行がA列まで拡張されること"""
        result = ExcelRangeCalculator.expand_axis_range("D200:Z200")
        assert result == "A200:Z200"

    def test_expand_axis_range_rectangle_unchanged(self):
        """矩形範囲はそのままであること"""
        result = ExcelRangeCalculator.expand_axis_range("B2:D5")
        assert result == "B2:D5"

    def test_expand_axis_range_empty_string(self):
        """空文字列はそのまま返すこと"""
        result = ExcelRangeCalculator.expand_axis_range("")
        assert result == ""

    def test_expand_axis_range_reverse_order_unchanged(self):
        """逆順序の範囲はそのまま返すこと（検証は別途）"""
        # 逆順序はexpandせず、後続の検証で弾かれる
        result = ExcelRangeCalculator.expand_axis_range("Z100:Z50")
        assert result == "Z100:Z50"

    def test_expand_axis_range_with_dollar_signs(self):
        """$記号付きの範囲が正しく処理されること"""
        result = ExcelRangeCalculator.expand_axis_range("$C$5")
        assert result == "C1:C5"

    # calculate_range_size のテスト

    def test_calculate_range_size_basic(self):
        """基本的な範囲サイズが計算されること"""
        rows, cols = ExcelRangeCalculator.calculate_range_size("A1:D10")
        assert rows == 10
        assert cols == 4

    def test_calculate_range_size_single_cell(self):
        """単一セルのサイズが(1, 1)であること"""
        rows, cols = ExcelRangeCalculator.calculate_range_size("B5")
        assert rows == 1
        assert cols == 1

    def test_calculate_range_size_single_row(self):
        """単一行のサイズが正しく計算されること"""
        rows, cols = ExcelRangeCalculator.calculate_range_size("A1:Z1")
        assert rows == 1
        assert cols == 26

    def test_calculate_range_size_single_column(self):
        """単一列のサイズが正しく計算されること"""
        rows, cols = ExcelRangeCalculator.calculate_range_size("A1:A100")
        assert rows == 100
        assert cols == 1

    def test_calculate_range_size_reverse_order_raises(self):
        """逆順序の範囲で(0, 0)が返ること（互換性維持）"""
        rows, cols = ExcelRangeCalculator.calculate_range_size("D10:A1")
        assert rows == 0
        assert cols == 0

    def test_calculate_range_size_reverse_column_raises(self):
        """逆順序の列で(0, 0)が返ること（互換性維持）"""
        rows, cols = ExcelRangeCalculator.calculate_range_size("D1:A10")
        assert rows == 0
        assert cols == 0

    def test_calculate_range_size_reverse_row_raises(self):
        """逆順序の行で(0, 0)が返ること（互換性維持）"""
        rows, cols = ExcelRangeCalculator.calculate_range_size("A10:D1")
        assert rows == 0
        assert cols == 0

    # normalize_column_range のテスト

    def test_normalize_column_range_single_column(self):
        """単一列指定が正規化されること"""
        result = ExcelRangeCalculator.normalize_column_range("J", 100)
        assert result == "J1:J100"

    def test_normalize_column_range_single_column_with_dollar(self):
        """$記号付き単一列が正規化されること"""
        result = ExcelRangeCalculator.normalize_column_range("$J", 100)
        assert result == "J1:J100"

    def test_normalize_column_range_lowercase(self):
        """小文字の列が大文字に変換されること"""
        result = ExcelRangeCalculator.normalize_column_range("j", 100)
        assert result == "J1:J100"

    def test_normalize_column_range_column_range(self):
        """列範囲指定が正規化されること"""
        result = ExcelRangeCalculator.normalize_column_range("J:K", 50)
        assert result == "J1:K50"

    def test_normalize_column_range_column_range_with_dollar(self):
        """$記号付き列範囲が正規化されること"""
        result = ExcelRangeCalculator.normalize_column_range("$J:$K", 50)
        assert result == "J1:K50"

    def test_normalize_column_range_reverse_order_raises(self):
        """逆順序の列範囲でValueErrorが発生すること"""
        with pytest.raises(ValueError) as exc_info:
            ExcelRangeCalculator.normalize_column_range("K:J", 50)
        assert "無効なセル範囲" in str(exc_info.value)
        assert "K:J" in str(exc_info.value)

    def test_normalize_column_range_already_normalized(self):
        """既に正規化済みの範囲はそのまま返すこと"""
        result = ExcelRangeCalculator.normalize_column_range("A1:B10", 100)
        assert result == "A1:B10"

    def test_normalize_column_range_single_cell(self):
        """単一セル指定はそのまま返すこと"""
        result = ExcelRangeCalculator.normalize_column_range("C5", 100)
        assert result == "C5"

    def test_normalize_column_range_empty_string(self):
        """空文字列はそのまま返すこと"""
        result = ExcelRangeCalculator.normalize_column_range("", 100)
        assert result == ""

    def test_normalize_column_range_whitespace(self):
        """空白のみの文字列はそのまま返すこと"""
        result = ExcelRangeCalculator.normalize_column_range("  ", 100)
        assert result == "  "

    def test_normalize_column_range_max_row_one(self):
        """max_row=1の場合も正しく処理されること"""
        result = ExcelRangeCalculator.normalize_column_range("A", 1)
        assert result == "A1:A1"

        result = ExcelRangeCalculator.normalize_column_range("A:C", 1)
        assert result == "A1:C1"
