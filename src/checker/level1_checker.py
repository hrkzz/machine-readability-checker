import re
from typing import Tuple, cast

from src.processor.context import TableContext
from src.checker.base_checker import BaseLevel1Checker
from .common import (
    get_excel_column_letter,
    MAX_EXAMPLES,
)
from src.checker.handler.format_handler import FormatHandler


class Level1Checker(BaseLevel1Checker):
    """
    Level1チェッカー
    CSV、XLS、XLSXの全ファイル形式に対応
    """
    
    def __init__(self):
        super().__init__()
        self.logger.add("logs/level1_checker.log", rotation="10 MB", retention="30 days", level="DEBUG")
        self.handler = FormatHandler()
    
    def get_supported_file_types(self) -> set:
        return {".csv", ".xls", ".xlsx"}
    
    def check_valid_file_format(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """ファイル形式の妥当性チェック"""
        return self.handler.check_file_format(filepath)
    
    def check_no_images_or_objects(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """画像・オブジェクトの存在チェック"""
        return self.handler.check_images_objects(filepath)
    
    def check_one_table_per_sheet(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """1シート1テーブルのチェック"""
        return self.handler.check_multiple_tables(ctx, workbook, filepath)
    
    def check_no_hidden_rows_or_columns(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """非表示行・列のチェック"""
        return self.handler.check_hidden_rows_columns(ctx, workbook, filepath)
    
    def check_no_notes_outside_table(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """表外注釈のチェック（全形式共通）"""
        if not ctx.upper_annotations.empty or not ctx.lower_annotations.empty:
            column_rows = ctx.row_indices.get("column_rows")
            data_end = ctx.row_indices.get("data_end")

            if column_rows is None or data_end is None:
                return False, "注釈行のチェックに必要な情報が不足しています"

            top = column_rows[0] if isinstance(column_rows, list) else cast(int, column_rows)
            bottom = cast(int, data_end) + 2
            return False, f"注釈行が検出されました（{top}行目より前、または{bottom}行目以降）"
        return True, "表外の注釈や備考はありません"
    
    def check_no_merged_cells(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """結合セルのチェック"""
        return self.handler.check_merged_cells(ctx, workbook, filepath)
    
    def check_no_format_based_semantics(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """書式による意味付けのチェック"""
        return self.handler.check_format_semantics(ctx, workbook, filepath)
    
    def check_no_whitespace_formatting(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """空白による体裁調整のチェック"""
        return self.handler.check_whitespace_formatting(ctx, workbook, filepath)
    
    def check_single_data_per_cell(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """1セル1データのチェック（全形式共通）"""
        pattern = re.compile(r"[\n,;/]")
        problems = []

        data_start = ctx.row_indices.get("data_start")
        if data_start is None:
            return False, "データ開始位置が不明です"

        start: int = cast(int, data_start)
        for row_idx_raw, row in ctx.data.iterrows():
            row_idx: int = cast(int, row_idx_raw)
            for col_idx, val in enumerate(row):
                if isinstance(val, str) and pattern.search(val):
                    excel_row = row_idx + 1 + start
                    excel_col_letter = get_excel_column_letter(col_idx + 1)
                    coord = f"{excel_col_letter}{excel_row}"
                    problems.append(f"{coord}: {repr(val)}")

        if problems:
            return False, f"複数データセルが検出されました（例: {problems[:MAX_EXAMPLES]}）"
        return True, "各セルに1データのみです"
    
    def check_no_platform_dependent_characters(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """機種依存文字のチェック"""
        return self.handler.check_platform_dependent_characters(ctx, workbook, filepath) 