"""
Excel固定行列（freeze_panes）管理ユーティリティ

固定行列情報の取得と変換を担当するヘルパークラス
"""

import logging

from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)


class ExcelPaneManager:
    """固定行列情報の取得と変換（全て staticmethod）"""

    @staticmethod
    def get_frozen_panes(sheet) -> tuple[int, int]:
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

    @staticmethod
    def format_freeze_panes(frozen_rows: int, frozen_cols: int) -> str:
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

    @staticmethod
    def validate_frozen_rows(frozen_rows: int, max_limit: int) -> tuple[bool, int]:
        """
        固定行数をDoS対策上限で検証

        Args:
            frozen_rows: 固定行数
            max_limit: 上限値

        Returns:
            (is_valid, validated_frozen_rows)のタプル
            - is_valid: 上限超過の場合のみFalse、それ以外はTrue
            - validated_frozen_rows: 負の値は0に丸める、上限超過は0、それ以外は元の値
        """
        # 負の値は無効として0に丸める
        if frozen_rows < 0:
            return (True, 0)
        if frozen_rows > max_limit:
            return (False, 0)
        return (True, frozen_rows)
