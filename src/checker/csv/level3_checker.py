from typing import Tuple
import pandas as pd
from pathlib import Path
from loguru import logger

from src.processor.context import TableContext
from src.checker.base.base_checker import BaseLevel3Checker
from src.checker.utils.common import is_likely_long_format


class CSVLevel3Checker(BaseLevel3Checker):
    """
    CSV専用のLevel3チェッカー
    """
    
    def __init__(self):
        super().__init__()
        self.logger.add("logs/csv_checker3.log", rotation="10 MB", retention="30 days", level="DEBUG")
    
    def get_supported_file_types(self) -> set:
        return {".csv"}
    
    def check_code_format_for_choices(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """CSV専用の選択肢のコード化チェック"""
        df = ctx.data
        candidate_cols = []

        for col in df.columns:
            try:
                series = df[col]
                if isinstance(series, pd.DataFrame):
                    continue

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
        """CSV専用のコード表の存在チェック"""
        # CSVファイルではコード表は別ファイルまたは別の手段で提供される
        return False, "CSVファイルのためコード表チェックをスキップします。コード表は別途提供してください"
    
    def check_question_master_exists(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """CSV専用の設問マスターの存在チェック"""
        # CSVファイルでは設問マスターは別ファイルまたは別の手段で提供される
        return False, "CSVファイルのため設問マスターチェックをスキップします。設問マスター（変数定義表）は別途提供してください"
    
    def check_metadata_presence(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """CSV専用のメタデータの存在チェック"""
        # CSVファイルではメタデータは別ファイルまたは別の手段で提供される
        return False, "CSVファイルのためメタデータチェックをスキップします。調査概要やメタデータは別途提供してください"
    
    def check_long_format_if_many_columns(self, ctx: TableContext, workbook: object, filepath: str) -> Tuple[bool, str]:
        """CSV専用のlong format形式のチェック"""
        if is_likely_long_format(ctx.data):
            return True, "縦型（long format）とみなされます"
        return False, "wide型であり、long型形式ではありません" 