from typing import Tuple
import pandas as pd
import warnings

from src.processor.context import TableContext
from src.checker.base_checker import BaseLevel3Checker
from src.checker.common import is_likely_long_format
from src.checker.handler.format_handler import FormatHandler

# PerformanceWarningを抑制
warnings.simplefilter("ignore", pd.errors.PerformanceWarning)


class Level3Checker(BaseLevel3Checker):
    """
    Level3チェッカー
    CSV、XLS、XLSXの全ファイル形式に対応
    """
    
    def __init__(self):
        super().__init__()
        self.logger.add("logs/level3_checker.log", rotation="10 MB", retention="30 days", level="DEBUG")
        self.handler = FormatHandler()
    
    def get_supported_file_types(self) -> set:
        return {".csv", ".xls", ".xlsx"}
    
    def check_code_format_for_choices(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """選択肢のコード化チェック（全形式共通）"""
        df = ctx.data
        candidate_cols = []
        processed_columns = 0

        for col in df.columns:
            try:
                series = df[col]
                if isinstance(series, pd.DataFrame):
                    continue

                processed_columns += 1
                unique_vals = series.dropna().unique()
                if len(unique_vals) < 10:
                    if any(not str(val).isdigit() for val in unique_vals):
                        candidate_cols.append(col)
            except Exception as e:
                return False, f"列 '{col}' でエラー発生: {e}"

        if candidate_cols:
            return False, f"コード表記ではない可能性のある列: {candidate_cols}"
        return True, "選択肢はコード表記されています"
    
    def check_codebook_exists(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """コード表の存在チェック（形式別処理）"""
        return self.handler.check_codebook_exists(ctx, workbook, filepath)
    
    def check_question_master_exists(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """設問マスターの存在チェック（形式別処理）"""
        return self.handler.check_question_master_exists(ctx, workbook, filepath)
    
    def check_metadata_presence(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """メタデータの存在チェック（形式別処理）"""
        return self.handler.check_metadata_presence(ctx, workbook, filepath)
    
    def check_long_format_if_many_columns(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """long format形式のチェック（全形式共通）"""
        col_count = len(ctx.data.columns)
        if is_likely_long_format(ctx.data):
            return True, "縦型（long format）とみなされます"
        return False, "wide型であり、long型形式ではありません" 