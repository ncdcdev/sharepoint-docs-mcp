"""
Excelスタイル抽出ユーティリティ

セルスタイル（色・サイズ）の抽出と変換を担当するヘルパークラス
"""

from typing import Any

from openpyxl.styles import Color


class ExcelStyleExtractor:
    """セルスタイル情報の抽出と変換（全て staticmethod）"""

    @staticmethod
    def color_to_hex(color: Color | None) -> str | None:
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

    @staticmethod
    def build_cell_size_cache(sheet) -> tuple[dict[str, float], dict[int, float]]:
        """
        列幅・行高さのキャッシュを構築（パフォーマンス最適化）

        Args:
            sheet: openpyxl Worksheet

        Returns:
            (col_widths, row_heights)のタプル
            - col_widths: 列文字 -> 幅のマップ
            - row_heights: 行番号 -> 高さのマップ
        """
        col_widths: dict[str, float] = {}
        row_heights: dict[int, float] = {}

        for col_letter, dim in sheet.column_dimensions.items():
            if dim.width:
                col_widths[col_letter] = dim.width

        for row_num, dim in sheet.row_dimensions.items():
            if dim.height:
                row_heights[row_num] = dim.height

        return (col_widths, row_heights)

    @staticmethod
    def extract_cell_styles(
        cell,
        col_widths: dict[str, float] | None,
        row_heights: dict[int, float] | None,
    ) -> dict[str, Any]:
        """
        セルからスタイル情報を抽出

        Args:
            cell: openpyxl Cell
            col_widths: 列幅のキャッシュ
            row_heights: 行高さのキャッシュ

        Returns:
            スタイル情報のdict（fill, width, heightなど）
        """
        styles: dict[str, Any] = {}

        # 背景色情報
        if cell.fill and cell.fill.patternType:
            fill_info = {
                "pattern_type": cell.fill.patternType,
            }
            fg_color = ExcelStyleExtractor.color_to_hex(cell.fill.fgColor)
            if fg_color:
                fill_info["fg_color"] = fg_color
            bg_color = ExcelStyleExtractor.color_to_hex(cell.fill.bgColor)
            if bg_color:
                fill_info["bg_color"] = bg_color
            styles["fill"] = fill_info

        # セルサイズ（列幅・行高さ）
        # MergedCellの場合は属性が存在しないため、hasattrでチェック
        if hasattr(cell, "column_letter") and hasattr(cell, "row"):
            if cell.column_letter and cell.row:
                # キャッシュから列幅を取得（パフォーマンス最適化）
                if col_widths and cell.column_letter in col_widths:
                    styles["width"] = col_widths[cell.column_letter]
                # キャッシュから行高さを取得（パフォーマンス最適化）
                if row_heights and cell.row in row_heights:
                    styles["height"] = row_heights[cell.row]

        return styles
